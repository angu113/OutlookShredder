using System.Text.RegularExpressions;
using Azure.Identity;
using Microsoft.Graph;
using Microsoft.Graph.Models;
using Microsoft.Graph.Users.Item.SendMail;
using OutlookShredder.Proxy.Models;

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
    private const string ClaimingCategory  = "RFQ-Claiming";

    // Job numbers are always auto-generated as uppercase alphanumeric (e.g. "RFQ [A1B2C3]").
    // Do NOT broaden this pattern — the strict format is intentional and guaranteed by the generator.
    private static readonly Regex _rfqSubjectRegex = new(
        @"^RFQ\s+\[([A-Za-z0-9]+)\]",
        RegexOptions.IgnoreCase | RegexOptions.Compiled);

    // Strips leading reply/forward prefixes (RE:, FW:, FWD:) that Outlook adds when
    // emails are forwarded into the RFQOut folder. Handles nested prefixes e.g. "RE: FW: RFQ [...]".
    private static readonly Regex _subjectPrefixRegex = new(
        @"^(RE|FW|FWD)\s*:\s*",
        RegexOptions.IgnoreCase | RegexOptions.Compiled);

    // Strips [EXTERNAL], [EXT], [CAUTION] etc. added by email security gateways.
    private static readonly Regex _externalTagRegex = new(
        @"^\[.*?\]\s*",
        RegexOptions.IgnoreCase | RegexOptions.Compiled);

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

    // ── RFQ Send ──────────────────────────────────────────────────────────────

    /// <summary>
    /// Sends an RFQ email via Microsoft Graph using app-only auth.
    /// Requires Mail.Send application permission on the Azure AD app.
    /// </summary>
    public async Task SendRfqEmailAsync(
        string subject,
        string body,
        IEnumerable<string> bccAddresses)
    {
        var from    = _config["Mail:FromAddress"]    ?? throw new InvalidOperationException("Mail:FromAddress not configured");
        var replyTo = _config["Mail:ReplyToAddress"] ?? throw new InvalidOperationException("Mail:ReplyToAddress not configured");

        var bcc = bccAddresses
            .Where(a => !string.IsNullOrWhiteSpace(a))
            .Select(a => new Recipient { EmailAddress = new EmailAddress { Address = a } })
            .ToList();

        await GetGraph().Users[from].SendMail.PostAsync(new()
        {
            Message = new Message
            {
                Subject     = subject,
                Body        = new ItemBody { ContentType = BodyType.Text, Content = body },
                From        = new Recipient { EmailAddress = new EmailAddress { Address = from } },
                ReplyTo     = [new Recipient { EmailAddress = new EmailAddress { Address = replyTo } }],
                BccRecipients = bcc,
            }
        });

        _log.LogInformation("[RFQ Send] Sent '{Subject}' BCC to {Count} recipient(s)", subject, bcc.Count);
    }

    // ── RFQ Import: scan a named folder ──────────────────────────────────────

    /// <summary>
    /// Scans a named mail folder in <paramref name="mailbox"/> and returns raw email
    /// data for all messages whose subject matches "RFQ [JobNo]".
    /// Shredder parses the body locally into line items.
    /// </summary>
    public async Task<List<RfqScanEmailDto>> ScanRfqFolderAsync(string mailbox, string folderName, int days = 0)
    {
        // Resolve folder ID from display name.
        var folders = await GetGraph().Users[mailbox].MailFolders
            .GetAsync(req =>
            {
                req.QueryParameters.Filter = $"displayName eq '{folderName}'";
                req.QueryParameters.Select = ["id", "displayName"];
            });

        var folder = folders?.Value?.FirstOrDefault()
            ?? throw new InvalidOperationException(
                $"Folder \"{folderName}\" not found in mailbox {mailbox}. " +
                "Create the folder in Outlook and copy RFQ sent emails into it.");

        // Build optional date filter — Graph supports $filter on sentDateTime for messages.
        var since = days > 0 ? DateTimeOffset.UtcNow.AddDays(-days) : (DateTimeOffset?)null;
        var dateFilter = since.HasValue
            ? $"sentDateTime ge {since.Value:yyyy-MM-ddTHH:mm:ssZ}"
            : null;

        // Fetch messages — filter by date server-side when possible; subject filter client-side.
        var firstPage = await GetGraph().Users[mailbox].MailFolders[folder.Id!].Messages
            .GetAsync(req =>
            {
                req.QueryParameters.Select  = ["subject", "sender", "toRecipients",
                                               "ccRecipients", "bccRecipients", "sentDateTime", "body"];
                req.QueryParameters.Top     = 500;
                req.QueryParameters.Orderby = ["sentDateTime desc"];
                if (dateFilter is not null)
                    req.QueryParameters.Filter = dateFilter;
            });

        var allMessages = new List<Microsoft.Graph.Models.Message>();
        var pageIterator = Microsoft.Graph.PageIterator<
                Microsoft.Graph.Models.Message,
                Microsoft.Graph.Models.MessageCollectionResponse>
            .CreatePageIterator(GetGraph(), firstPage!, msg =>
            {
                // Stop paging once we pass the date window (messages are ordered newest-first)
                if (since.HasValue && msg.SentDateTime.HasValue && msg.SentDateTime.Value < since.Value)
                    return false;
                allMessages.Add(msg);
                return true;
            });
        await pageIterator.IterateAsync();

        var result = new List<RfqScanEmailDto>();

        _log.LogInformation("[RFQ Scan] {Mailbox}/{Folder}: {Total} messages fetched",
            mailbox, folderName, allMessages.Count);

        foreach (var msg in allMessages)
        {
            var rawSubject = msg.Subject ?? "";
            var sentAt     = msg.SentDateTime?.ToString("yyyy-MM-dd HH:mm") ?? "unknown";
            var sender     = msg.Sender?.EmailAddress?.Address ?? "unknown";

            // Strip any leading RE:/FW:/FWD: prefixes and [EXTERNAL]/[EXT]/[CAUTION] gateway
            // tags before matching. Loop handles combinations e.g. "RE: [EXTERNAL] RFQ [...]".
            var subject = rawSubject;
            bool stripped;
            do
            {
                stripped = false;
                if (_subjectPrefixRegex.IsMatch(subject)) { subject = _subjectPrefixRegex.Replace(subject, "", 1); stripped = true; }
                if (_externalTagRegex.IsMatch(subject))   { subject = _externalTagRegex.Replace(subject, "", 1);   stripped = true; }
            }
            while (stripped);

            var m = _rfqSubjectRegex.Match(subject);

            if (!m.Success)
            {
                // Diagnose why the subject didn't match: help identify format variations.
                string reason;
                if (string.IsNullOrWhiteSpace(subject))
                    reason = "empty subject";
                else if (!subject.StartsWith("RFQ", StringComparison.OrdinalIgnoreCase))
                    reason = $"does not start with 'RFQ' (starts with '{subject[..Math.Min(20, subject.Length)]}')";
                else if (!subject.Contains('['))
                    reason = "missing '[' — expected format: RFQ [JOBNO]";
                else if (!subject.Contains(']'))
                    reason = "missing ']' — expected format: RFQ [JOBNO]";
                else
                    reason = $"subject has '[...]' but content doesn't match [A-Za-z0-9]+ — actual: '{subject}'";

                _log.LogWarning("[RFQ Scan] SKIP | {Sent} | {Sender} | Subject: {Subject} | Reason: {Reason}",
                    sentAt, sender, rawSubject, reason);
                continue;
            }

            var rfqId = m.Groups[1].Value.ToUpperInvariant();

            var requester = mailbox;
            var senderName = msg.Sender?.EmailAddress?.Name;
            if (!string.IsNullOrWhiteSpace(senderName)) requester = senderName;

            var recipients = new List<string>();
            foreach (var r in msg.ToRecipients ?? [])
                if (r.EmailAddress?.Address is string a) recipients.Add(a);
            foreach (var r in msg.CcRecipients ?? [])
                if (r.EmailAddress?.Address is string a) recipients.Add(a);
            foreach (var r in msg.BccRecipients ?? [])
                if (r.EmailAddress?.Address is string a) recipients.Add(a);

            var bodyText    = msg.Body?.Content ?? "";
            var contentType = msg.Body?.ContentType == BodyType.Html ? "html" : "text";

            result.Add(new RfqScanEmailDto
            {
                RfqId           = rfqId,
                Subject         = subject,
                SentAt          = msg.SentDateTime?.UtcDateTime ?? DateTime.UtcNow,
                Requester       = requester,
                EmailRecipients = string.Join("\n", recipients),
                MailboxSource   = mailbox,
                BodyText        = bodyText,
                ContentType     = contentType,
            });
        }

        return result;
    }

    /// <summary>
    /// Returns inbox messages already tagged with "RFQ-Processed" (newest first).
    /// Used to populate the Reprocess Supplier Emails scan list.
    /// </summary>
    public async Task<List<ProcessedEmailDto>> GetProcessedMessagesAsync(string mailbox, int top = 200)
    {
        var result = await GetGraph()
            .Users[mailbox]
            .MailFolders["inbox"]
            .Messages
            .GetAsync(req =>
            {
                req.QueryParameters.Filter  = $"categories/any(c: c eq '{ProcessedCategory}')";
                req.QueryParameters.Select  = ["id", "subject", "from", "receivedDateTime",
                                               "bodyPreview", "categories"];
                req.QueryParameters.Top     = top;
                req.QueryParameters.Orderby = ["receivedDateTime desc"];
            });

        return (result?.Value ?? [])
            .Select(m => new ProcessedEmailDto
            {
                MessageId  = m.Id ?? "",
                Subject    = m.Subject ?? "(no subject)",
                From       = m.From?.EmailAddress?.Address ?? "",
                ReceivedAt = m.ReceivedDateTime?.UtcDateTime ?? DateTime.MinValue,
                Preview    = m.BodyPreview ?? "",
                IsUnknown  = m.Categories?.Contains("Unknown", StringComparer.OrdinalIgnoreCase) == true,
            })
            .ToList();
    }

    /// <summary>
    /// Fetches a single full message (with body content) by ID.
    /// Used by the reprocess flow to fetch body + attachment metadata before calling Claude.
    /// </summary>
    public async Task<Message?> GetMessageByIdAsync(string mailbox, string messageId)
    {
        return await GetGraph()
            .Users[mailbox]
            .Messages[messageId]
            .GetAsync(req =>
            {
                req.QueryParameters.Select = ["id", "subject", "from", "receivedDateTime",
                                              "body", "hasAttachments", "bodyPreview", "categories"];
            });
    }

    /// <summary>
    /// Returns inbox messages received at or after <paramref name="since"/> that have NOT
    /// been tagged with the "RFQ-Processed" category. Filtering is done server-side.
    /// </summary>
    public async Task<List<Message>> GetMessagesAsync(string mailbox, DateTimeOffset since)
    {
        var filter = $"receivedDateTime ge {since.UtcDateTime:yyyy-MM-ddTHH:mm:ssZ}" +
                     $" and not (categories/any(c: c eq '{ProcessedCategory}'))" +
                     $" and not (categories/any(c: c eq '{ClaimingCategory}'))";

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
            ["id", "subject", "body", "receivedDateTime"], windowSeconds: 120);

        var msg = messages.FirstOrDefault();
        if (msg is null)
        {
            _log.LogWarning("[Mail] GetBodyAsync: no message found for {From} near {At} (±2 min)", emailFrom, receivedAt);
            return null;
        }

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

        _log.LogWarning("[Mail] GetAttachmentAsync: '{File}' not found for {From} near {At} (±10 min)", filename, emailFrom, receivedAt);
        return null;
    }

    // ── Shared message search ─────────────────────────────────────────────────

    /// <summary>
    /// Finds the Graph message ID for an email identified by sender address and received time.
    /// Uses a ±5-minute window. Returns null if no match is found or Outlook is not configured.
    /// </summary>
    public async Task<string?> FindMessageIdAsync(string emailFrom, DateTimeOffset receivedAt)
    {
        var messages = await FindMessagesAsync(emailFrom, receivedAt,
            ["id", "receivedDateTime"], windowSeconds: 300);
        return messages.FirstOrDefault()?.Id;
    }

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
    /// For large attachments (>~3 MB) the list endpoint returns null ContentBytes;
    /// those are fetched individually by ID so the bytes are always populated.
    /// </summary>
    public async Task<List<Attachment>> GetAttachmentsAsync(string mailbox, string messageId)
    {
        var result = await GetGraph()
            .Users[mailbox]
            .Messages[messageId]
            .Attachments
            .GetAsync();

        var items = result?.Value ?? [];

        // Fetch individual attachments whose ContentBytes are missing (large-file case).
        for (int i = 0; i < items.Count; i++)
        {
            if (items[i] is FileAttachment fa && fa.ContentBytes is null && fa.Id is not null)
            {
                try
                {
                    var full = await GetGraph()
                        .Users[mailbox]
                        .Messages[messageId]
                        .Attachments[fa.Id]
                        .GetAsync() as FileAttachment;

                    if (full is not null)
                        items[i] = full;
                }
                catch (Exception ex)
                {
                    _log.LogWarning(ex, "[Mail] Could not fetch attachment {Id} by ID — ContentBytes will be null", fa.Id);
                }
            }
        }

        return items;
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
    /// Removes "RFQ-Processed" (and PO-specific categories) from inbox messages in
    /// <paramref name="mailbox"/> that were received within the last <paramref name="days"/> days
    /// and have a PO-related subject ("Purchase Order #HSK-PO").
    /// This lets the next poll cycle re-run PO extraction on those emails.
    /// </summary>
    public async Task<int> UnmarkPoEmailsAsync(string mailbox, int days)
    {
        // Unmark from both inbox (replies received) and sentitems (outbound POs).
        var inbox = await UnmarkPoEmailsInFolderAsync(mailbox, "inbox",     days, useReceivedDateTime: true);
        var sent  = await UnmarkPoEmailsInFolderAsync(mailbox, "sentitems", days, useReceivedDateTime: false);
        return inbox + sent;
    }

    private async Task<int> UnmarkPoEmailsInFolderAsync(
        string mailbox, string folder, int days, bool useReceivedDateTime)
    {
        var since      = DateTimeOffset.UtcNow.AddDays(-days);
        var dateField  = useReceivedDateTime ? "receivedDateTime" : "sentDateTime";
        var filter     = $"{dateField} ge {since.UtcDateTime:yyyy-MM-ddTHH:mm:ssZ}" +
                         $" and categories/any(c: c eq '{ProcessedCategory}')";

        var unmarkCount = 0;
        string? nextLink = null;

        do
        {
            MessageCollectionResponse? result;
            if (nextLink is null)
            {
                result = await GetGraph()
                    .Users[mailbox]
                    .MailFolders[folder]
                    .Messages
                    .GetAsync(req =>
                    {
                        req.QueryParameters.Filter = filter;
                        req.QueryParameters.Select = ["id", "subject", "categories"];
                        req.QueryParameters.Top    = 50;
                    });
            }
            else
            {
                result = await new Microsoft.Graph.Users.Item.MailFolders.Item.Messages.MessagesRequestBuilder(
                    nextLink, GetGraph().RequestAdapter).GetAsync();
            }

            var messages = result?.Value ?? [];
            nextLink     = result?.OdataNextLink;

            foreach (var msg in messages)
            {
                if (msg.Id is null) continue;

                // Client-side subject filter — only unmark actual PO emails
                var subject      = msg.Subject ?? "";
                var cleanSubject = System.Text.RegularExpressions.Regex
                    .Replace(subject, @"^(RE:|FW:|FWD:|\[EXTERNAL\])\s*", "",
                        System.Text.RegularExpressions.RegexOptions.IgnoreCase |
                        System.Text.RegularExpressions.RegexOptions.Compiled).Trim();
                if (!cleanSubject.StartsWith("Purchase Order #HSK-PO", StringComparison.OrdinalIgnoreCase))
                    continue;

                var stripped = (msg.Categories ?? [])
                    .Where(c => !string.Equals(c, ProcessedCategory,  StringComparison.OrdinalIgnoreCase)
                             && !string.Equals(c, "PO-Processed",     StringComparison.OrdinalIgnoreCase)
                             && !string.Equals(c, "PO-NoExtract",     StringComparison.OrdinalIgnoreCase)
                             && !string.Equals(c, ClaimingCategory,   StringComparison.OrdinalIgnoreCase))
                    .ToList();
                try
                {
                    await GetGraph()
                        .Users[mailbox]
                        .Messages[msg.Id]
                        .PatchAsync(new Message { Categories = stripped });
                    unmarkCount++;
                    _log.LogInformation("[Mail] Un-marked PO email {Id} in {Folder}: {Subject}", msg.Id, folder, msg.Subject);
                }
                catch (Exception ex)
                {
                    _log.LogWarning(ex, "[Mail] Could not un-mark PO message {Id} in {Folder}", msg.Id, folder);
                }
            }
        }
        while (nextLink is not null);

        return unmarkCount;
    }

    /// <summary>
    /// Returns Sent Items messages sent at or after <paramref name="since"/> whose subject
    /// starts with "Purchase Order #HSK-PO" and have NOT yet been tagged with "RFQ-Processed".
    /// Used by the mail poller to pick up outbound POs for extraction.
    /// </summary>
    public async Task<List<Message>> GetUnprocessedSentPoMessagesAsync(string mailbox, DateTimeOffset since)
    {
        var filter = $"sentDateTime ge {since.UtcDateTime:yyyy-MM-ddTHH:mm:ssZ}" +
                     $" and not (categories/any(c: c eq '{ProcessedCategory}'))" +
                     $" and not (categories/any(c: c eq '{ClaimingCategory}'))";

        var result = await GetGraph()
            .Users[mailbox]
            .MailFolders["sentitems"]
            .Messages
            .GetAsync(req =>
            {
                req.QueryParameters.Filter  = filter;
                req.QueryParameters.Select  = ["id", "subject", "from", "sentDateTime",
                                               "body", "hasAttachments", "bodyPreview", "categories"];
                req.QueryParameters.Top     = 50;
                req.QueryParameters.Orderby = ["sentDateTime desc"];
            });

        // Client-side filter for PO subject — sentitems may contain many outbound emails.
        return (result?.Value ?? [])
            .Where(m =>
            {
                var s = System.Text.RegularExpressions.Regex
                    .Replace(m.Subject ?? "", @"^(RE:|FW:|FWD:|\[EXTERNAL\])\s*", "",
                        System.Text.RegularExpressions.RegexOptions.IgnoreCase).Trim();
                return s.StartsWith("Purchase Order #HSK-PO", StringComparison.OrdinalIgnoreCase);
            })
            .ToList();
    }

    /// <summary>
    /// Stamps "RFQ-Claiming" on a message immediately so other proxy instances skip it
    /// during concurrent polls. The claim is replaced by "RFQ-Processed" once done.
    /// Errors are swallowed — a missed claim just means another proxy may duplicate work,
    /// which is acceptable since the SharePoint dedup prevents duplicate writes.
    /// </summary>
    public async Task MarkClaimingAsync(string mailbox, string messageId)
    {
        try
        {
            await GetGraph()
                .Users[mailbox]
                .Messages[messageId]
                .PatchAsync(new Message { Categories = [ClaimingCategory] });

            _log.LogDebug("[Mail] Claimed message {Id}", messageId);
        }
        catch (Exception ex)
        {
            _log.LogDebug(ex, "[Mail] Could not claim message {Id} — another proxy may process it concurrently", messageId);
        }
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
