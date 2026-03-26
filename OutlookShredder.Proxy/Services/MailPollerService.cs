using System.Text.RegularExpressions;
using Microsoft.Graph.Models;
using OutlookShredder.Proxy.Models;

namespace OutlookShredder.Proxy.Services;

/// <summary>
/// Background service that polls a configured mailbox on a timer.
/// For each new message it:
///   1. Sends the email body to Claude for RFQ extraction
///   2. Sends any PDF/DOCX attachments to Claude
///   3. Writes extracted product rows to SharePoint
///   4. Stamps "RFQ-Processed" category on the message via Graph so it is
///      excluded from future polls at the server-side query level.
///
/// Config keys (add to User Secrets or appsettings.Development.json):
///   Mail:MailboxAddress       — UPN of the inbox to watch, e.g. "rfq@example.com"
///   Mail:PollIntervalSeconds  — how often to poll (default: 30)
///   Mail:LookbackHours        — rolling lookback window per poll (default: 24)
/// </summary>
public class MailPollerService : BackgroundService
{
    private readonly IConfiguration       _config;
    private readonly MailService          _mail;
    private readonly ClaudeService        _claude;
    private readonly SharePointService    _sp;
    private readonly ILogger<MailPollerService> _log;

    private static readonly Regex JobRefRegex =
        new(@"\[([A-Z0-9]{6})\]", RegexOptions.Compiled | RegexOptions.IgnoreCase);

    private static readonly Regex HtmlTagRegex =
        new(@"<[^>]+>", RegexOptions.Compiled);

    private static readonly Regex WhitespaceRegex =
        new(@"\s{2,}", RegexOptions.Compiled);

    public MailPollerService(
        IConfiguration          config,
        MailService             mail,
        ClaudeService           claude,
        SharePointService       sp,
        ILogger<MailPollerService> log)
    {
        _config = config;
        _mail   = mail;
        _claude = claude;
        _sp     = sp;
        _log    = log;
    }

    protected override async Task ExecuteAsync(CancellationToken stoppingToken)
    {
        var mailbox         = _config["Mail:MailboxAddress"]
            ?? throw new InvalidOperationException(
                "Mail:MailboxAddress is not configured. " +
                "Add it to User Secrets or appsettings.Development.json.");

        var intervalSeconds = int.TryParse(_config["Mail:PollIntervalSeconds"], out var s) ? s : 30;
        var lookbackHours   = double.TryParse(_config["Mail:LookbackHours"], out var h) ? h : 24.0;
        var interval        = TimeSpan.FromSeconds(intervalSeconds);

        _log.LogInformation("[Mail] Poller started — mailbox={Mailbox} interval={Interval}s lookback={Lookback}h",
            mailbox, intervalSeconds, lookbackHours);

        while (!stoppingToken.IsCancellationRequested)
        {
            var since = DateTimeOffset.UtcNow.AddHours(-lookbackHours);
            try
            {
                await PollAsync(mailbox, since, stoppingToken);
            }
            catch (Exception ex) when (!stoppingToken.IsCancellationRequested)
            {
                _log.LogError(ex, "[Mail] Poll cycle failed");
            }

            await Task.Delay(interval, stoppingToken).ConfigureAwait(false);
        }
    }

    // ── Poll one cycle ────────────────────────────────────────────────────────

    private async Task PollAsync(string mailbox, DateTimeOffset since, CancellationToken ct)
    {
        var messages = await _mail.GetMessagesAsync(mailbox, since);

        if (messages.Count == 0)
        {
            _log.LogDebug("[Mail] No unprocessed messages");
            return;
        }

        _log.LogInformation("[Mail] {Count} unprocessed message(s) found", messages.Count);

        foreach (var msg in messages)
        {
            if (ct.IsCancellationRequested) break;
            await ProcessMessageAsync(mailbox, msg, ct);
        }
    }

    // ── Process one message ───────────────────────────────────────────────────

    private async Task ProcessMessageAsync(string mailbox, Message msg, CancellationToken ct)
    {
        var subject  = msg.Subject ?? "(no subject)";
        var fromAddr = msg.From?.EmailAddress?.Address ?? "unknown";
        var received = msg.ReceivedDateTime?.ToString("o") ?? DateTime.UtcNow.ToString("o");

        var rawBody = msg.Body?.Content ?? string.Empty;
        if (msg.Body?.ContentType == BodyType.Html)
            rawBody = HtmlTagRegex.Replace(rawBody, " ");
        var body = WhitespaceRegex.Replace(rawBody, " ").Trim();

        var jobRefs = JobRefRegex.Matches(subject + " " + body)
            .Select(m => m.Groups[1].Value.ToUpperInvariant())
            .Distinct()
            .ToList();

        // Gate: skip emails without a [XXXXXX] job reference to avoid unnecessary Claude calls.
        if (jobRefs.Count == 0)
        {
            _log.LogDebug("[Mail] No job reference in \"{Subject}\" — skipping extraction, marking processed", subject);
            if (msg.Id is not null) await _mail.MarkProcessedAsync(mailbox, msg.Id);
            return;
        }

        _log.LogInformation("[Mail] Processing: \"{Subject}\" from {From} refs=[{Refs}]",
            subject, fromAddr, string.Join(", ", jobRefs));

        // 1. Extract from body
        await RunExtractionAsync(new ExtractRequest
        {
            Content       = body[..Math.Min(body.Length, 12_000)],
            SourceType    = "body",
            JobRefs       = jobRefs,
            EmailSubject  = subject,
            EmailFrom     = fromAddr,
            ReceivedAt    = received,
            HasAttachment = msg.HasAttachments ?? false,
        }, "body", null);

        // 2. Extract from PDF / DOCX attachments
        if (msg.HasAttachments == true && !ct.IsCancellationRequested && msg.Id is not null)
        {
            var attachments = await _mail.GetAttachmentsAsync(mailbox, msg.Id);
            foreach (var att in attachments)
            {
                if (ct.IsCancellationRequested) break;
                if (att is not FileAttachment fa) continue;

                var contentType = (fa.ContentType ?? "").ToLowerInvariant();
                if (!contentType.Contains("pdf") &&
                    !contentType.Contains("wordprocessingml") &&
                    !contentType.Contains("msword")) continue;

                if (fa.ContentBytes is null) continue;

                await RunExtractionAsync(new ExtractRequest
                {
                    Content      = string.Empty,
                    SourceType   = "attachment",
                    FileName     = fa.Name,
                    ContentType  = fa.ContentType,
                    Base64Data   = Convert.ToBase64String(fa.ContentBytes),
                    BodyContext  = body[..Math.Min(body.Length, 2_000)],
                    JobRefs      = jobRefs,
                    EmailSubject = subject,
                    EmailFrom    = fromAddr,
                    ReceivedAt   = received,
                }, "attachment", fa.Name);
            }
        }

        if (msg.Id is not null) await _mail.MarkProcessedAsync(mailbox, msg.Id);
    }

    // ── Claude → SharePoint ───────────────────────────────────────────────────

    private async Task RunExtractionAsync(ExtractRequest req, string source, string? fileName)
    {
        try
        {
            var extraction = await _claude.ExtractAsync(req);

            if (extraction is null || extraction.Products.Count == 0)
            {
                _log.LogInformation("[Mail] No products extracted from {Source}", source);
                return;
            }

            _log.LogInformation("[Mail] Extracted {Count} product(s) from {Source}",
                extraction.Products.Count, source);

            for (int i = 0; i < extraction.Products.Count; i++)
            {
                var row = await _sp.WriteProductRowAsync(
                    extraction, extraction.Products[i], req, source, fileName, i);

                if (row.Success)
                    _log.LogInformation("[Mail] SP row written: '{Product}' -> {Url}",
                        extraction.Products[i].ProductName, row.SpWebUrl);
                else
                    _log.LogWarning("[Mail] SP write failed for '{Product}': {Error}",
                        extraction.Products[i].ProductName, row.Error);
            }
        }
        catch (Exception ex)
        {
            _log.LogError(ex, "[Mail] Extraction failed for {Source} ({File})", source, fileName ?? "body");
        }
    }
}
