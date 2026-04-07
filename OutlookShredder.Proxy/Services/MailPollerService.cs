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
///   Mail:MaxEmailsPerMinute   — max Claude API calls per minute (default: 5)
/// </summary>
public class MailPollerService : BackgroundService
{
    private readonly IConfiguration            _config;
    private readonly MailService               _mail;
    private readonly ClaudeService             _claude;
    private readonly SharePointService         _sp;
    private readonly RfqNotificationService    _notifications;
    private readonly ILogger<MailPollerService> _log;

    private static readonly Regex JobRefRegex =
        new(@"\[([A-Z0-9]{6})\]", RegexOptions.Compiled | RegexOptions.IgnoreCase);

    private static readonly Regex HtmlLineBreakRegex =
        new(@"<br\s*/?>|</p>|</div>|</tr>|</li>|</h[1-6]>", RegexOptions.Compiled | RegexOptions.IgnoreCase);

    private static readonly Regex HtmlTagRegex =
        new(@"<[^>]+>", RegexOptions.Compiled);

    private static readonly Regex HorizontalWhitespaceRegex =
        new(@"[ \t]{2,}", RegexOptions.Compiled);

    private static readonly Regex ExcessiveNewlineRegex =
        new(@"\n{3,}", RegexOptions.Compiled);

    // Sliding-window rate limiter — tracks timestamps of recent Claude calls.
    private readonly Queue<DateTimeOffset> _claudeCallTimestamps = new();
    private readonly SemaphoreSlim         _rateLimitLock        = new(1, 1);

    // Signalled by TriggerReprocessAllAsync to request an immediate full scan.
    private readonly SemaphoreSlim _reprocessTrigger = new(0, 1);

    /// <summary>
    /// Reprocesses a specific set of already-processed messages by fetching each one from
    /// Graph and running the full extraction pipeline (Claude + SharePoint write + re-stamp).
    /// Called synchronously by the reprocess-selected endpoint; awaited before responding.
    /// </summary>
    public async Task ReprocessMessagesAsync(IEnumerable<string> messageIds, CancellationToken ct)
    {
        var mailbox             = _config["Mail:MailboxAddress"]
            ?? throw new InvalidOperationException("Mail:MailboxAddress not configured");
        var maxPerMinute        = int.TryParse(_config["Mail:MaxEmailsPerMinute"],        out var r)  ? Math.Max(1, r) : 30;
        var bodyContextChars    = int.TryParse(_config["Mail:BodyContextChars"],          out var bc) ? bc             : 2_000;
        var extractBodyNoJobRef = bool.TryParse(_config["Mail:ExtractBodyWithoutJobRef"], out var eb) && eb;

        foreach (var messageId in messageIds)
        {
            if (ct.IsCancellationRequested) break;

            var msg = await _mail.GetMessageByIdAsync(mailbox, messageId);
            if (msg is null)
            {
                _log.LogWarning("[Reprocess] Message {Id} not found — skipping", messageId);
                continue;
            }

            _log.LogInformation("[Reprocess] Reprocessing \"{Subject}\" from {From}",
                msg.Subject, msg.From?.EmailAddress?.Address);
            await ProcessMessageAsync(mailbox, msg, maxPerMinute, bodyContextChars, extractBodyNoJobRef, ct);
        }
    }

    /// <summary>
    /// Triggers an immediate full scan of all unprocessed inbox messages (no lookback limit).
    /// Returns immediately; the scan runs on the background poller thread.
    /// </summary>
    public void TriggerReprocessAll()
    {
        // Release the semaphore (max 1 — extra calls are no-ops).
        if (_reprocessTrigger.CurrentCount == 0)
            _reprocessTrigger.Release();
    }

    public MailPollerService(
        IConfiguration             config,
        MailService                mail,
        ClaudeService              claude,
        SharePointService          sp,
        RfqNotificationService     notifications,
        ILogger<MailPollerService> log)
    {
        _config        = config;
        _mail          = mail;
        _claude        = claude;
        _sp            = sp;
        _notifications = notifications;
        _log           = log;
    }

    protected override async Task ExecuteAsync(CancellationToken stoppingToken)
    {
        var mailbox         = _config["Mail:MailboxAddress"]
            ?? throw new InvalidOperationException(
                "Mail:MailboxAddress is not configured. " +
                "Add it to User Secrets or appsettings.Development.json.");

        var intervalSeconds          = int.TryParse(_config["Mail:PollIntervalSeconds"],      out var s)  ? s          : 30;
        var lookbackHours            = double.TryParse(_config["Mail:LookbackHours"],          out var h)  ? h          : 24.0;
        var maxPerMinute             = int.TryParse(_config["Mail:MaxEmailsPerMinute"],        out var r)  ? Math.Max(1, r) : 30;
        var bodyContextChars         = int.TryParse(_config["Mail:BodyContextChars"],          out var bc) ? bc         : 2_000;
        var extractBodyWithoutJobRef = bool.TryParse(_config["Mail:ExtractBodyWithoutJobRef"], out var eb) && eb;
        var interval                 = TimeSpan.FromSeconds(intervalSeconds);

        _log.LogInformation(
            "[Mail] Poller started — mailbox={Mailbox} interval={Interval}s lookback={Lookback}h rateLimit={Rate}/min extractBodyNoRef={ExtractNoRef}",
            mailbox, intervalSeconds, lookbackHours, maxPerMinute, extractBodyWithoutJobRef);

        bool firstCycle = true;
        while (!stoppingToken.IsCancellationRequested)
        {
            // Full scan when: first cycle on startup OR triggered via TriggerReprocessAll().
            // Subsequent regular cycles use the rolling lookback window.
            bool fullScan = firstCycle || _reprocessTrigger.CurrentCount > 0;
            if (fullScan && !firstCycle) await _reprocessTrigger.WaitAsync(stoppingToken);

            var since = fullScan
                ? DateTimeOffset.MinValue
                : DateTimeOffset.UtcNow.AddHours(-lookbackHours);

            if (fullScan)
                _log.LogInformation("[Mail] {Reason} — scanning all unprocessed messages (no lookback limit)",
                    firstCycle ? "Startup cycle" : "Triggered reprocess");

            try
            {
                await PollAsync(mailbox, since, maxPerMinute, bodyContextChars, extractBodyWithoutJobRef, stoppingToken);
            }
            catch (Exception ex) when (!stoppingToken.IsCancellationRequested)
            {
                _log.LogError(ex, "[Mail] Poll cycle failed");
            }

            firstCycle = false;
            await Task.Delay(interval, stoppingToken).ConfigureAwait(false);
        }
    }

    // ── Poll one cycle ────────────────────────────────────────────────────────

    private async Task PollAsync(
        string mailbox, DateTimeOffset since, int maxPerMinute,
        int bodyContextChars, bool extractBodyWithoutJobRef, CancellationToken ct)
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
            await ProcessMessageAsync(mailbox, msg, maxPerMinute, bodyContextChars, extractBodyWithoutJobRef, ct);
        }
    }

    // ── Rate limiter ──────────────────────────────────────────────────────────

    /// <summary>
    /// Waits until a Claude API call is allowed under the sliding-window rate limit.
    /// Records the call timestamp so subsequent calls count against the same window.
    /// </summary>
    private async Task AcquireRateSlotAsync(int maxPerMinute, CancellationToken ct)
    {
        var window = TimeSpan.FromMinutes(1);

        while (true)
        {
            await _rateLimitLock.WaitAsync(ct);
            try
            {
                var now    = DateTimeOffset.UtcNow;
                var cutoff = now - window;

                // Evict timestamps older than one minute.
                while (_claudeCallTimestamps.Count > 0 && _claudeCallTimestamps.Peek() <= cutoff)
                    _claudeCallTimestamps.Dequeue();

                if (_claudeCallTimestamps.Count < maxPerMinute)
                {
                    _claudeCallTimestamps.Enqueue(now);
                    return;   // slot acquired
                }

                // Calculate exactly how long until the oldest call leaves the window.
                var waitUntil = _claudeCallTimestamps.Peek() + window;
                var precise   = waitUntil - now + TimeSpan.FromMilliseconds(50); // 50 ms buffer
                _log.LogDebug("[Mail] Rate limit reached ({Max}/min) — waiting {Ms}ms", maxPerMinute, (int)precise.TotalMilliseconds);
            }
            finally
            {
                _rateLimitLock.Release();
            }

            // Sleep precisely until the next slot opens rather than a fixed back-off.
            // Re-read under lock for accuracy; use a short minimum to avoid tight spin.
            TimeSpan sleepFor;
            await _rateLimitLock.WaitAsync(ct);
            try
            {
                var now2   = DateTimeOffset.UtcNow;
                var cutoff = now2 - window;
                while (_claudeCallTimestamps.Count > 0 && _claudeCallTimestamps.Peek() <= cutoff)
                    _claudeCallTimestamps.Dequeue();
                sleepFor = _claudeCallTimestamps.Count >= maxPerMinute
                    ? (_claudeCallTimestamps.Peek() + window - now2 + TimeSpan.FromMilliseconds(50))
                    : TimeSpan.Zero;
            }
            finally { _rateLimitLock.Release(); }

            if (sleepFor > TimeSpan.Zero)
                await Task.Delay(sleepFor, ct);
        }
    }

    // ── Process one message ───────────────────────────────────────────────────

    private async Task ProcessMessageAsync(
        string mailbox, Message msg, int maxPerMinute,
        int bodyContextChars, bool extractBodyWithoutJobRef, CancellationToken ct)
    {
        var subject  = msg.Subject ?? "(no subject)";
        var fromAddr = msg.From?.EmailAddress?.Address ?? "unknown";
        var received = msg.ReceivedDateTime?.ToString("o") ?? DateTime.UtcNow.ToString("o");

        // Attach email context to every log line emitted while processing this message.
        using var logScope = _log.BeginScope(new Dictionary<string, object?>
        {
            ["EmailSubject"] = subject,
            ["EmailFrom"]    = fromAddr,
            ["MessageId"]    = msg.Id,
        });

        var rawBody = msg.Body?.Content ?? string.Empty;
        if (msg.Body?.ContentType == BodyType.Html)
        {
            rawBody = HtmlLineBreakRegex.Replace(rawBody, "\n");
            rawBody = HtmlTagRegex.Replace(rawBody, " ");
            rawBody = System.Net.WebUtility.HtmlDecode(rawBody);
            rawBody = HorizontalWhitespaceRegex.Replace(rawBody, " ");
            rawBody = ExcessiveNewlineRegex.Replace(rawBody, "\n\n");
        }
        else
        {
            rawBody = System.Net.WebUtility.HtmlDecode(rawBody);
        }
        var body = rawBody.Trim();

        var jobRefs = JobRefRegex.Matches(subject + " " + body)
            .Select(m => m.Groups[1].Value.ToUpperInvariant())
            .Distinct()
            .ToList();

        // Decide whether to call Claude.
        // Claude is always used when a job reference is present or there is an attachment
        // (which may be a quote PDF/DOCX even without a reference in the body).
        // Body-only emails with no job reference bypass Claude and get a direct placeholder row
        // unless ExtractBodyWithoutJobRef is explicitly enabled in config.
        bool hasJobRef    = jobRefs.Count > 0;
        bool hasAttachment = msg.HasAttachments == true;
        bool sendToClaude = hasJobRef || hasAttachment || extractBodyWithoutJobRef;

        if (!hasJobRef && !hasAttachment)
            _log.LogInformation("[Mail] No job ref or attachment in \"{Subject}\" from {From} — {Action}",
                subject, fromAddr, sendToClaude ? "sending to Claude" : "writing direct row under [000000]");
        else
            _log.LogInformation("[Mail] Processing: \"{Subject}\" from {From} refs=[{Refs}] attachments={Att}",
                subject, fromAddr, string.Join(", ", jobRefs), hasAttachment);

        var bodySnippet = body[..Math.Min(body.Length, bodyContextChars)];

        bool supplierUnknown = false;

        if (!sendToClaude)
        {
            // Fast path: no job ref and no attachment — write a placeholder row directly
            // without spending a Claude API call. SP will assign [000000] or [WHOIS] as appropriate.
            var req = new ExtractRequest
            {
                Content      = string.Empty,
                EmailBody    = body,
                SourceType   = "body",
                JobRefs      = jobRefs,
                EmailSubject = subject,
                EmailFrom    = fromAddr,
                ReceivedAt   = received,
            };
            var extraction = new RfqExtraction();
            var placeholder = new ProductLine
            {
                SupplierProductComments = "Email recorded without extraction — no job reference or quote attachment detected."
            };
            var row = await _sp.WriteProductRowAsync(extraction, placeholder, req, "body", null, 0);
            supplierUnknown = row.SupplierUnknown;
            if (row.Success) _notifications.NotifyRfqProcessed();
        }
        else if (!hasAttachment || msg.Id is null)
        {
            // No attachments — extract pricing from body; always write at least one row.
            supplierUnknown = await RunExtractionAsync(new ExtractRequest
            {
                Content       = body[..Math.Min(body.Length, 12_000)],
                EmailBody     = body,
                SourceType    = "body",
                JobRefs       = jobRefs,
                EmailSubject  = subject,
                EmailFrom     = fromAddr,
                ReceivedAt    = received,
                HasAttachment = false,
            }, "body", null, maxPerMinute, ct);
        }
        else
        {
            // Has attachments — prefer attachment data for pricing; skip body-only extraction
            // so only one row is written per email.  Body is stored in EmailBody column.
            var attachments = await _mail.GetAttachmentsAsync(mailbox, msg.Id);
            bool processedAny = false;

            foreach (var att in attachments)
            {
                if (ct.IsCancellationRequested) break;
                if (att is not FileAttachment fa) continue;

                var contentType = (fa.ContentType ?? "").ToLowerInvariant();
                if (!contentType.Contains("pdf") &&
                    !contentType.Contains("wordprocessingml") &&
                    !contentType.Contains("msword")) continue;

                if (fa.ContentBytes is null) continue;

                supplierUnknown = await RunExtractionAsync(new ExtractRequest
                {
                    Content      = string.Empty,
                    EmailBody    = body,
                    SourceType   = "attachment",
                    FileName     = fa.Name,
                    ContentType  = fa.ContentType,
                    Base64Data   = Convert.ToBase64String(fa.ContentBytes),
                    BodyContext  = bodySnippet,
                    JobRefs      = jobRefs,
                    EmailSubject = subject,
                    EmailFrom    = fromAddr,
                    ReceivedAt   = received,
                    HasAttachment = true,
                }, "attachment", fa.Name, maxPerMinute, ct);

                processedAny = true;
            }

            // No recognisable attachment format — fall back to body extraction.
            if (!processedAny && !ct.IsCancellationRequested)
            {
                supplierUnknown = await RunExtractionAsync(new ExtractRequest
                {
                    Content       = body[..Math.Min(body.Length, 12_000)],
                    EmailBody     = body,
                    SourceType    = "body",
                    JobRefs       = jobRefs,
                    EmailSubject  = subject,
                    EmailFrom     = fromAddr,
                    ReceivedAt    = received,
                    HasAttachment = true,
                }, "body", null, maxPerMinute, ct);
            }
        }

        if (msg.Id is not null)
        {
            var extra = supplierUnknown ? "Unknown" : null;
            if (supplierUnknown)
                _log.LogInformation("[Mail] Supplier unrecognised in \"{Subject}\" — stamping 'Unknown' category", subject);
            await _mail.MarkProcessedAsync(mailbox, msg.Id, extra);
        }
    }

    // ── Claude → SharePoint ───────────────────────────────────────────────────

    // Returns true if the supplier was not found in the reference list.
    private async Task<bool> RunExtractionAsync(
        ExtractRequest req, string source, string? fileName, int maxPerMinute, CancellationToken ct)
    {
        await AcquireRateSlotAsync(maxPerMinute, ct);

        try
        {
            var extraction = await _claude.ExtractAsync(req);

            // Always write at least one row — even when nothing useful could be extracted —
            // so every processed email has a visible record in SharePoint.
            var products = extraction?.Products ?? [];
            if (products.Count == 0)
            {
                products = [new ProductLine { SupplierProductComments = "No products could be extracted from this email." }];
                extraction ??= new RfqExtraction();
                _log.LogInformation("[Mail] No products extracted from {Source} — writing placeholder row", source);
            }
            else
            {
                _log.LogInformation("[Mail] Extracted {Count} product(s) from {Source}", products.Count, source);
            }

            // Write product rows sequentially so they all share one SupplierResponse row.
            // Concurrent writes caused a race: every call independently found no existing SR
            // and created its own, producing duplicate SR records per multi-product email.
            var rowList = new List<SpWriteResult>(products.Count);
            foreach (var (p, i) in products.Select((p, i) => (p, i)))
            {
                if (ct.IsCancellationRequested) break;
                rowList.Add(await _sp.WriteProductRowAsync(extraction!, p, req, source, fileName, i));
            }
            var rows = rowList.ToArray();

            bool anyUnknown    = false;
            bool anySuccessful = false;
            for (int i = 0; i < rows.Length; i++)
            {
                var row = rows[i];
                if (row.SupplierUnknown)
                    anyUnknown = true;
                else if (row.Success)
                {
                    anySuccessful = true;
                    _log.LogInformation("[Mail] SP row {Action}: '{Product}' -> {Url}",
                        row.Updated ? "updated" : "inserted", products[i].ProductName, row.SpWebUrl);
                }
                else
                    _log.LogWarning("[Mail] SP upsert failed for '{Product}': {Error}",
                        products[i].ProductName, row.Error);
            }

            if (anySuccessful)
            {
                var notification = new Models.RfqProcessedNotification
                {
                    EventType    = "SR",
                    SupplierName = rows.Zip(products)
                                       .FirstOrDefault(x => x.First.Success && !x.First.SupplierUnknown)
                                       .First?.SupplierName,
                    RfqId = req.JobRefs.FirstOrDefault()?.Trim('[', ']'),
                    Products = rows.Zip(products)
                                   .Where(x => x.First.Success)
                                   .Select(x => new Models.RfqNotificationProduct
                                   {
                                       Name       = x.First.ProductName,
                                       TotalPrice = x.Second.TotalPrice,
                                   }).ToList(),
                };
                _notifications.NotifyRfqProcessed(notification);
            }

            return anyUnknown;
        }
        catch (Exception ex)
        {
            _log.LogError(ex, "[Mail] Extraction failed for {Source} ({File})", source, fileName ?? "body");
            return false;
        }
    }
}
