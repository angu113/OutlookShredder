using System.Text.RegularExpressions;
using Azure.Identity;
using Microsoft.Graph;
using Microsoft.Graph.Models;

namespace OutlookShredder.Proxy.Services;

/// <summary>
/// Reads messages and attachments from a user mailbox via Microsoft Graph API
/// using the same app-only (client credential) auth as SharePointService.
///
/// Azure AD app requires:  Mail.ReadWrite  (Application permission, admin consented)
/// Config:  SharePoint:TenantId / ClientId / ClientSecret  (reused from SharePoint config)
///          Mail:MailboxAddress  — the UPN/email of the mailbox to monitor
/// </summary>
public class MailService
{
    private readonly IConfiguration _config;
    private readonly ILogger<MailService> _log;
    private GraphServiceClient? _graph;

    private const string ProcessedCategory = "RFQ-Processed";

    private static readonly Regex _htmlTag    = new(@"<[^>]+>",  RegexOptions.Compiled);
    private static readonly Regex _whitespace = new(@"\s{2,}",   RegexOptions.Compiled);

    public MailService(IConfiguration config, ILogger<MailService> log)
    {
        _config = config;
        _log    = log;
    }

    private GraphServiceClient GetGraph()
    {
        if (_graph is not null) return _graph;

        var tenantId     = _config["SharePoint:TenantId"]     ?? throw new InvalidOperationException("SharePoint:TenantId not set");
        var clientId     = _config["SharePoint:ClientId"]     ?? throw new InvalidOperationException("SharePoint:ClientId not set");
        var clientSecret = _config["SharePoint:ClientSecret"] ?? throw new InvalidOperationException("SharePoint:ClientSecret not set");

        var credential = new ClientSecretCredential(tenantId, clientId, clientSecret);
        _graph = new GraphServiceClient(credential, ["https://graph.microsoft.com/.default"]);
        return _graph;
    }

    /// <summary>
    /// Returns inbox messages received at or after <paramref name="since"/> that have NOT
    /// been tagged with the "RFQ-Processed" category. Filtering is done server-side.
    /// </summary>
    public async Task<List<Message>> GetMessagesAsync(string mailbox, DateTimeOffset since)
    {
        var filter = $"receivedDateTime ge {since.UtcDateTime:yyyy-MM-ddTHH:mm:ssZ}" +
                     $" and not (categories/any(c: c eq '{ProcessedCategory}'))";

        _log.LogDebug("[Mail] Querying inbox for {Mailbox} since {Since}", mailbox, since);

        var result = await GetGraph()
            .Users[mailbox]
            .MailFolders["inbox"]
            .Messages
            .GetAsync(req =>
            {
                req.QueryParameters.Filter  = filter;
                req.QueryParameters.Select  = ["id", "subject", "from", "receivedDateTime",
                                               "body", "hasAttachments", "bodyPreview", "categories"];
                req.QueryParameters.Top     = 50;
                req.QueryParameters.Orderby = ["receivedDateTime desc"];
            });

        return result?.Value ?? [];
    }

    // ── Dashboard: retrieve email body text ──────────────────────────────────

    /// <summary>
    /// Finds the message sent from <paramref name="emailFrom"/> at approximately
    /// <paramref name="receivedAt"/> (±2 minutes) and returns its plain-text body.
    /// </summary>
    public async Task<(string Subject, string BodyText)?> GetBodyAsync(
        string emailFrom, DateTimeOffset receivedAt)
    {
        var messages = await FindMessagesAsync(emailFrom, receivedAt,
            ["id", "subject", "body", "receivedDateTime"]);

        var msg = messages.FirstOrDefault();
        if (msg is null) return null;

        // Return the raw body content — Shredder's CleanBody handles HTML stripping
        // and entity decoding so structure (paragraphs, line breaks) is preserved.
        var text = msg.Body?.Content ?? string.Empty;

        return (msg.Subject ?? "(no subject)", text);
    }

    // ── Dashboard: retrieve attachment bytes ──────────────────────────────────

    /// <summary>
    /// Finds the message matching the sender / timestamp and returns the bytes
    /// of the named attachment so the WPF client can open it with the default viewer.
    /// </summary>
    public async Task<(string ContentType, byte[] Bytes, string FileName)?> GetAttachmentAsync(
        string emailFrom, DateTimeOffset receivedAt, string filename)
    {
        // Try narrow window first; widen if no match found (handles timestamp drift
        // when the stored ReceivedAt comes from a different mailbox than the proxy searches).
        var messages = await FindMessagesAsync(emailFrom, receivedAt,
            ["id", "hasAttachments", "receivedDateTime"]);

        if (!messages.Any(m => m.HasAttachments == true))
            messages = await FindMessagesAsync(emailFrom, receivedAt,
                ["id", "hasAttachments", "receivedDateTime"], windowSeconds: 600);

        var mailbox = GetMailbox();

        // Multiple emails from the same sender can fall within the ±2 min window.
        // Scan each candidate (that has attachments) until we find the one with the named file.
        foreach (var msg in messages.Where(m => m.HasAttachments == true && m.Id is not null))
        {
            var listResult = await GetGraph()
                .Users[mailbox]
                .Messages[msg.Id!]
                .Attachments
                .GetAsync();

            // Match on the base Attachment class — the collection response returns the base type.
            var meta = listResult?.Value?
                .FirstOrDefault(a => string.Equals(a.Name, filename, StringComparison.OrdinalIgnoreCase));

            if (meta?.Id is null) continue;

            // Fetch by ID: returns a properly-typed FileAttachment with ContentBytes.
            var fa = await GetGraph().Users[mailbox].Messages[msg.Id!].Attachments[meta.Id].GetAsync()
                as FileAttachment;

            if (fa?.ContentBytes is null) continue;

            return (fa.ContentType ?? "application/octet-stream", fa.ContentBytes, fa.Name ?? filename);
        }

        return null;
    }

    // ── Shared message search ─────────────────────────────────────────────────

    private string GetMailbox() =>
        _config["Mail:MailboxAddress"]
        ?? throw new InvalidOperationException("Mail:MailboxAddress not configured in appsettings / User Secrets");

    private async Task<List<Message>> FindMessagesAsync(
        string emailFrom, DateTimeOffset receivedAt, string[] select,
        int windowSeconds = 30)
    {
        var dtFrom = receivedAt.AddSeconds(-windowSeconds).UtcDateTime.ToString("yyyy-MM-ddTHH:mm:ssZ");
        var dtTo   = receivedAt.AddSeconds(+windowSeconds).UtcDateTime.ToString("yyyy-MM-ddTHH:mm:ssZ");

        var filter = $"from/emailAddress/address eq '{emailFrom}'" +
                     $" and receivedDateTime ge {dtFrom}" +
                     $" and receivedDateTime le {dtTo}";

        var result = await GetGraph()
            .Users[GetMailbox()]
            .Messages
            .GetAsync(req =>
            {
                req.QueryParameters.Filter = filter;
                req.QueryParameters.Select = select;
                req.QueryParameters.Top    = 5;
            });

        // Return sorted by closeness to the requested timestamp so callers
        // that use FirstOrDefault() get the best match first.
        return (result?.Value ?? [])
            .OrderBy(m => Math.Abs((m.ReceivedDateTime!.Value - receivedAt).TotalSeconds))
            .ToList();
    }

    /// <summary>
    /// Returns all attachments for a message. Caller should filter by ContentType.
    /// </summary>
    public async Task<List<Attachment>> GetAttachmentsAsync(string mailbox, string messageId)
    {
        var result = await GetGraph()
            .Users[mailbox]
            .Messages[messageId]
            .Attachments
            .GetAsync();

        return result?.Value ?? [];
    }

    /// <summary>
    /// Removes "RFQ-Processed" (and "Unknown") categories from every message in the inbox
    /// so the next poll cycle will reprocess all emails.
    /// Returns the count of messages that were un-marked.
    /// </summary>
    public async Task<int> UnmarkAllAsync(string mailbox)
    {
        var unmarkCount = 0;

        // Fetch batches of marked messages and clear them until none remain.
        while (true)
        {
            var result = await GetGraph()
                .Users[mailbox]
                .MailFolders["inbox"]
                .Messages
                .GetAsync(req =>
                {
                    req.QueryParameters.Filter = $"categories/any(c: c eq '{ProcessedCategory}')";
                    req.QueryParameters.Select = ["id", "categories"];
                    req.QueryParameters.Top    = 50;
                });

            var messages = result?.Value ?? [];
            if (messages.Count == 0) break;

            foreach (var msg in messages)
            {
                if (msg.Id is null) continue;
                var stripped = (msg.Categories ?? [])
                    .Where(c => !string.Equals(c, ProcessedCategory, StringComparison.OrdinalIgnoreCase)
                             && !string.Equals(c, "Unknown",         StringComparison.OrdinalIgnoreCase))
                    .ToList();
                try
                {
                    await GetGraph()
                        .Users[mailbox]
                        .Messages[msg.Id]
                        .PatchAsync(new Message { Categories = stripped });
                    unmarkCount++;
                }
                catch (Exception ex)
                {
                    _log.LogWarning(ex, "[Mail] Could not un-mark message {Id}", msg.Id);
                }
            }
        }

        return unmarkCount;
    }

    /// <summary>
    /// Stamps "RFQ-Processed" (and optionally one extra category) on a message
    /// so it is excluded from future polls.
    /// Errors are logged but not rethrown — a missed stamp means the message will be retried,
    /// which is safer than crashing the poll loop.
    /// </summary>
    public async Task MarkProcessedAsync(string mailbox, string messageId, string? additionalCategory = null)
    {
        var categories = additionalCategory is not null
            ? new List<string> { ProcessedCategory, additionalCategory }
            : new List<string> { ProcessedCategory };
        try
        {
            await GetGraph()
                .Users[mailbox]
                .Messages[messageId]
                .PatchAsync(new Message { Categories = categories });

            _log.LogDebug("[Mail] Stamped categories [{Categories}] on message {Id}",
                string.Join(", ", categories), messageId);
        }
        catch (Exception ex)
        {
            _log.LogWarning(ex, "[Mail] Could not stamp categories on message {Id} — will retry next poll", messageId);
        }
    }
}
