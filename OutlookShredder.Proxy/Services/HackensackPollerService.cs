using System.Text.RegularExpressions;
using MailKit;
using MailKit.Net.Imap;
using MailKit.Search;
using MimeKit;
using OutlookShredder.Proxy.Models;

namespace OutlookShredder.Proxy.Services;

/// <summary>
/// Polls the hackensack@metalsupermarkets.com inbox via IMAP/OAuth2 using
/// a delegated token (device-code flow). Processes new messages through the
/// same AI extraction pipeline as MailPollerService.
///
/// Config keys (appsettings.json / secrets):
///   HackensackMail:PollIntervalSeconds      — default 30
///   HackensackMail:LookbackHours            — default 24
///   HackensackMail:ExtractBodyWithoutJobRef — default false
/// </summary>
public class HackensackPollerService : BackgroundService
{
    private readonly DelegatedTokenProvider         _tokens;
    private readonly AiServiceFactory               _aiFactory;
    private readonly SharePointService              _sp;
    private readonly ProductCatalogService          _catalog;
    private readonly RfqNotificationService         _notifications;
    private readonly ShrConvInRouter                _shrRouter;
    private readonly IConfiguration                 _config;
    private readonly ILogger<HackensackPollerService> _log;
    private const string MailboxAddress = "hackensack@metalsupermarkets.com";

    private static readonly Regex JobRefRegex =
        new(@"\[(HQ[A-Z0-9]{6}|[A-Z0-9]{6})\]", RegexOptions.Compiled | RegexOptions.IgnoreCase);

    private static readonly Regex JobRefBareRegex =
        new(@"\bRFQ\s+(?>(?:Number\b\s*)?)(HQ[A-Z0-9]{6}|(?!\d{6})[A-Z0-9]{6})\b|\bJob\s*Ref(?:erence|\b)(?>(?:\s+Number\b)?\s*[:#]?\s*)(HQ[A-Z0-9]{6}|(?!\d{6})[A-Z0-9]{6})\b",
            RegexOptions.Compiled | RegexOptions.IgnoreCase);

    private static readonly Regex HtmlLineBreakRegex =
        new(@"<br\s*/?>|</p>|</div>|</tr>|</li>|</h[1-6]>", RegexOptions.Compiled | RegexOptions.IgnoreCase);

    private static readonly Regex HtmlTagRegex    = new(@"<[^>]+>",  RegexOptions.Compiled);
    private static readonly Regex HorizWhitespace = new(@"[ \t]{2,}", RegexOptions.Compiled);
    private static readonly Regex ExcessNewlines   = new(@"\n{3,}",   RegexOptions.Compiled);

    private const string ProcessedFlag = "RFQ-Processed";

    public HackensackPollerService(
        DelegatedTokenProvider              tokens,
        AiServiceFactory                    aiFactory,
        SharePointService                   sp,
        ProductCatalogService               catalog,
        RfqNotificationService              notifications,
        ShrConvInRouter                     shrRouter,
        IConfiguration                      config,
        ILogger<HackensackPollerService>    log)
    {
        _tokens        = tokens;
        _aiFactory     = aiFactory;
        _sp            = sp;
        _catalog       = catalog;
        _notifications = notifications;
        _shrRouter     = shrRouter;
        _config        = config;
        _log           = log;
    }

    protected override async Task ExecuteAsync(CancellationToken ct)
    {
        var interval      = int.TryParse(_config["HackensackMail:PollIntervalSeconds"], out var i) ? i : 30;
        var lookbackHours = int.TryParse(_config["HackensackMail:LookbackHours"],       out var l) ? l : 24;
        var extractBody   = bool.TryParse(_config["HackensackMail:ExtractBodyWithoutJobRef"], out var e) && e;

        if (!_tokens.IsConfigured)
        {
            _log.LogInformation("[HackensackMail] No MSAL cache — polling disabled until hackensack-msal-cache.bin is present");
            return;
        }

        _log.LogInformation("[HackensackMail] Starting — polling every {Interval}s, lookback {Hours}h", interval, lookbackHours);

        await Task.Delay(TimeSpan.FromSeconds(10), ct);

        while (!ct.IsCancellationRequested)
        {
            try
            {
                await PollAsync(lookbackHours, extractBody, ct);
            }
            catch (OperationCanceledException) { break; }
            catch (Exception ex)
            {
                _log.LogError(ex, "[HackensackMail] Poll cycle error");
            }

            await Task.Delay(TimeSpan.FromSeconds(interval), ct);
        }
    }

    private async Task PollAsync(int lookbackHours, bool extractBodyWithoutJobRef, CancellationToken ct)
    {
        var token = await _tokens.GetAccessTokenAsync(ct);
        var since = DateTimeOffset.UtcNow.AddHours(-lookbackHours);

        using var imap = new ImapClient();
        await imap.ConnectAsync("outlook.office365.com", 993, MailKit.Security.SecureSocketOptions.SslOnConnect, ct);
        await imap.AuthenticateAsync(new MailKit.Security.SaslMechanismOAuth2(MailboxAddress, token), ct);

        var inbox = imap.Inbox;
        await inbox.OpenAsync(FolderAccess.ReadWrite, ct);

        var query = SearchQuery.DeliveredAfter(since.UtcDateTime)
            .And(SearchQuery.NotKeyword(ProcessedFlag));

        var uids = await inbox.SearchAsync(query, ct);

        if (uids.Count > 0)
            _log.LogInformation("[HackensackMail] {Count} unprocessed message(s) in {Hours}h window", uids.Count, lookbackHours);

        foreach (var uid in uids)
        {
            if (ct.IsCancellationRequested) break;

            var msg = await inbox.GetMessageAsync(uid, ct);

            await inbox.StoreAsync(uid,
                new StoreFlagsRequest(StoreAction.Add, new[] { ProcessedFlag }), ct);

            await ProcessImapMessageAsync(msg, uid.ToString(), extractBodyWithoutJobRef, ct);
        }

        await imap.DisconnectAsync(true, ct);
    }

    private async Task ProcessImapMessageAsync(
        MimeMessage msg, string uid, bool extractBodyWithoutJobRef, CancellationToken ct)
    {
        var subject   = msg.Subject ?? "(no subject)";
        var fromBox   = msg.From.Mailboxes.FirstOrDefault();
        var from      = fromBox?.Address ?? "unknown";
        var received  = msg.Date.ToString("o");
        var messageId = msg.MessageId ?? uid;

        // Prefer plain text; fall back to cleaned HTML.
        var body = msg.TextBody ?? "";
        if (string.IsNullOrEmpty(body) && msg.HtmlBody is { } html)
        {
            body = HtmlLineBreakRegex.Replace(html, "\n");
            body = HtmlTagRegex.Replace(body, " ");
            body = System.Net.WebUtility.HtmlDecode(body);
            body = HorizWhitespace.Replace(body, " ");
            body = ExcessNewlines.Replace(body, "\n\n");
        }
        else
        {
            body = System.Net.WebUtility.HtmlDecode(body);
        }
        body = body.Trim();

        // Only non-image file attachments count for processing decisions.
        var fileAttachments = msg.Attachments.OfType<MimePart>()
            .Where(a => !a.ContentType.MimeType.StartsWith("image/", StringComparison.OrdinalIgnoreCase))
            .ToList();
        var hasAtt = fileAttachments.Count > 0;

        var searchText = subject + " " + body;
        var jobRefs = JobRefRegex.Matches(searchText)
            .Select(m => m.Groups[1].Value.ToUpperInvariant())
            .Distinct().ToList();

        if (jobRefs.Count == 0)
        {
            jobRefs = JobRefBareRegex.Matches(searchText)
                .Select(m => (m.Groups[1].Success ? m.Groups[1] : m.Groups[2]).Value.ToUpperInvariant())
                .Distinct().ToList();
        }

        var shrResult = await _shrRouter.TryRouteAsync(
            searchText:     searchText,
            fromAddr:       from,
            subject:        subject,
            body:           body,
            messageId:      messageId,
            hasAttachments: hasAtt,
            receivedAt:     msg.Date);

        if (shrResult.ShrRfqId is not null &&
            !jobRefs.Contains(shrResult.ShrRfqId, StringComparer.OrdinalIgnoreCase))
            jobRefs.Insert(0, shrResult.ShrRfqId);

        bool hasJobRef = jobRefs.Count > 0;
        bool sendToAi  = hasJobRef || hasAtt || extractBodyWithoutJobRef;

        _log.LogInformation("[HackensackMail] Processing: \"{Subject}\" from {From} refs=[{Refs}] attachments={Att}",
            subject, from, string.Join(", ", jobRefs), hasAtt);

        var bodySnippet = body[..Math.Min(body.Length, 2000)];

        if (!sendToAi)
        {
            var req = new ExtractRequest
            {
                Content      = string.Empty,
                EmailBody    = body,
                SourceType   = "body",
                JobRefs      = jobRefs,
                EmailSubject = subject,
                EmailFrom    = from,
                ReceivedAt   = received,
            };
            await _sp.WriteProductRowAsync(new RfqExtraction(),
                new ProductLine { SupplierProductComments = "No job reference or attachment — placeholder row." },
                req, "body", null, 0, messageId);
            return;
        }

        if (hasAtt)
        {
            foreach (var attachment in fileAttachments)
            {
                var attName = attachment.FileName ?? "attachment";
                var attType = attachment.ContentType.MimeType;

                if (attachment.Content is null) continue;   // no decodable content — skip
                using var mem = new MemoryStream();
                await attachment.Content.DecodeToAsync(mem);
                var attBytes = Convert.ToBase64String(mem.ToArray());

                var req = new ExtractRequest
                {
                    Content       = attBytes,
                    Base64Data    = attBytes,
                    ContentType   = attType,
                    FileName      = attName,
                    SourceType    = "attachment",
                    EmailBody     = body,
                    BodyContext   = bodySnippet,
                    EmailSubject  = subject,
                    EmailFrom     = from,
                    ReceivedAt    = received,
                    HasAttachment = true,
                    JobRefs       = jobRefs,
                };
                await RunExtractionAsync(req, "attachment", attName, messageId, ct);
            }
        }

        if (!hasAtt || extractBodyWithoutJobRef)
        {
            var req = new ExtractRequest
            {
                Content       = body,
                SourceType    = "body",
                EmailBody     = body,
                EmailSubject  = subject,
                EmailFrom     = from,
                ReceivedAt    = received,
                HasAttachment = hasAtt,
                JobRefs       = jobRefs,
            };
            await RunExtractionAsync(req, "body", null, messageId, ct);
        }
    }

    private async Task RunExtractionAsync(
        ExtractRequest req, string source, string? fileName, string messageId, CancellationToken ct)
    {
        if (req.RliItems.Count == 0)
        {
            var validRef = req.JobRefs
                .Select(r => r.Trim('[', ']'))
                .FirstOrDefault(r => !string.IsNullOrEmpty(r) && r != "000000" && r != "WHOIS");

            if (validRef is not null)
            {
                try
                {
                    var rli = await _sp.ReadRfqLineItemsByRfqIdAsync(validRef);
                    if (rli.Count > 0) req.RliItems = rli;
                }
                catch (Exception ex)
                {
                    _log.LogWarning(ex, "[HackensackMail] RLI fetch failed for [{Ref}]", validRef);
                }
            }
        }

        try
        {
            var extraction = await _aiFactory.GetService().ExtractRfqAsync(req, ct);

            var products = extraction?.Products ?? [];
            if (products.Count == 0)
            {
                products = [new ProductLine { SupplierProductComments = "No products could be extracted from this email." }];
                extraction ??= new RfqExtraction();
            }
            else
            {
                products = ProductDeduplicator.Deduplicate(products, source, false, _log);
            }

            var rowList = new List<SpWriteResult>(products.Count);
            foreach (var (p, i) in products.Select((p, i) => (p, i)))
            {
                if (ct.IsCancellationRequested) break;
                rowList.Add(await _sp.WriteProductRowAsync(extraction!, p, req, source, fileName, i, messageId));
            }

            var anyInserted = rowList.Any(r => r.Success && !r.Updated && !r.SupplierUnknown);
            var anyUpdated  = rowList.Any(r => r.Success &&  r.Updated && !r.SupplierUnknown);

            if (anyInserted || anyUpdated)
            {
                _notifications.NotifyRfqProcessed(new RfqProcessedNotification
                {
                    EventType    = "SR",
                    SupplierName = rowList.Zip(products)
                                         .FirstOrDefault(x => x.First.Success && !x.First.SupplierUnknown)
                                         .First?.SupplierName,
                    RfqId = rowList.FirstOrDefault(r => r.Success && !string.IsNullOrEmpty(r.RfqId))?.RfqId
                            ?? req.JobRefs.FirstOrDefault()?.Trim('[', ']'),
                    MessageId = messageId,
                    Products  = rowList.Zip(products)
                                      .Where(x => x.First.Success && !x.First.SupplierUnknown)
                                      .Select(x => new RfqNotificationProduct
                                      {
                                          Name       = x.First.ProductName,
                                          TotalPrice = x.Second.TotalPrice,
                                      }).ToList(),
                });
            }
        }
        catch (Exception ex)
        {
            _log.LogError(ex, "[HackensackMail] Extraction error for message {MsgId}", messageId);
        }
    }
}
