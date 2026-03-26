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
    /// Stamps the "RFQ-Processed" category on a message so it is excluded from future polls.
    /// Errors are logged but not rethrown — a missed stamp means the message will be retried,
    /// which is safer than crashing the poll loop.
    /// </summary>
    public async Task MarkProcessedAsync(string mailbox, string messageId)
    {
        try
        {
            await GetGraph()
                .Users[mailbox]
                .Messages[messageId]
                .PatchAsync(new Message { Categories = [ProcessedCategory] });

            _log.LogDebug("[Mail] Stamped '{Category}' on message {Id}", ProcessedCategory, messageId);
        }
        catch (Exception ex)
        {
            _log.LogWarning(ex, "[Mail] Could not stamp '{Category}' on message {Id} — will retry next poll",
                ProcessedCategory, messageId);
        }
    }
}
