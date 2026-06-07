using System.Text.RegularExpressions;
using Microsoft.Graph.Models;
using OutlookShredder.Proxy.Models;

namespace OutlookShredder.Proxy.Services;

/// <summary>
/// Background service that polls a configured mailbox on a timer.
/// For each new message it:
///   1. Sends the email body to the AI for RFQ extraction
///   2. Sends any PDF/DOCX attachments to the AI
///   3. Writes extracted product rows to SharePoint
///   4. Stamps "RFQ-Processed" category on the message via Graph so it is
///      excluded from future polls at the server-side query level.
///
/// Config keys (add to User Secrets or appsettings.Development.json):
///   Mail:MailboxAddress       — UPN of the inbox to watch, e.g. "rfq@example.com"
///   Mail:PollIntervalSeconds  — how often to poll (default: 30)
///   Mail:LookbackHours        — rolling lookback window per poll (default: 24)
///   Mail:MaxEmailsPerMinute   — max AI API calls per minute (default: 5)
/// </summary>
public class MailPollerService : BackgroundService
{
    private readonly IConfiguration             _config;
    private readonly MailService                _mail;
    private readonly AiServiceFactory           _aiFactory;
    private readonly SharePointService          _sp;
    private readonly ProductCatalogService      _catalog;
    private readonly RfqNotificationService     _notifications;
    private readonly RfqSummaryQueue            _summaryQueue;
    private readonly SliCacheService            _sliCache;
    private readonly ShrConvInRouter            _shrRouter;
    private readonly CatalogAnalysisService     _analysis;
    private readonly ILogger<MailPollerService> _log;

    // Accepts two formats:
    //   HQ:     HQ + 6 alphanumeric (e.g. HQBX9EWM) — 8 chars
    //   Legacy: 6 alphanumeric (e.g. UA2ZJC or AW0001) — new initials+Crockford4 IDs fit here
    // HQ alt is listed first so 8-char matches aren't truncated to 6.
    private static readonly Regex JobRefRegex =
        new(@"\[(HQ[A-Z0-9]{6}|[A-Z0-9]{6})\]", RegexOptions.Compiled | RegexOptions.IgnoreCase);

    // Fallback: bare ID appearing after "RFQ" or "Job Ref" (no brackets). Same two formats.
    // Used only when JobRefRegex finds nothing, to catch supplier replies that strip brackets.
    //
    // "Ref" is required (not optional) to prevent "job number" matching "number" as an ID.
    // Two false-positive classes fixed via atomic groups:
    //   ERENCE — "JOB REFERENCE" (no ID follows): (?:erence|\b) forces the engine to consume
    //     the full "erence" suffix or stop at a word boundary; neither path lets it un-consume
    //     "erence" and re-offer it as a 6-char ID.
    //   NUMBER — "Job Reference Number" or "RFQ Number" where "Number" is a 6-char label word:
    //     (?>(?:Number\b\s*)?) and (?>(?:\s+Number\b)?\s*...) atomically absorb the optional
    //     "Number" label so it can never fall through into the ID capture group.
    private static readonly Regex JobRefBareRegex =
        new(@"\bRFQ\s+(?>(?:Number\b\s*)?)(HQ[A-Z0-9]{6}|(?!\d{6})[A-Z0-9]{6})\b|\bJob\s*Ref(?:erence|\b)(?>(?:\s+Number\b)?\s*[:#]?\s*)(HQ[A-Z0-9]{6}|(?!\d{6})[A-Z0-9]{6})\b",
            RegexOptions.Compiled | RegexOptions.IgnoreCase);

    // SHR tracking token ([SHR:{rfqId}]) routing lives in ShrConvInRouter so the
    // add-in extract endpoint can honour it too — see Services/ShrConvInRouter.cs.

    // Subject prefixes to strip before PO detection (RE:, FW:, [EXTERNAL], etc.)
    private static readonly Regex SubjectPrefixRegex =
        new(@"^(\s*(RE|FW|FWD)\s*:\s*|\s*\[.*?\]\s*)+", RegexOptions.Compiled | RegexOptions.IgnoreCase);

    private static readonly Regex HtmlLineBreakRegex =
        new(@"<br\s*/?>|</p>|</div>|</tr>|</li>|</h[1-6]>", RegexOptions.Compiled | RegexOptions.IgnoreCase);

    private static readonly Regex HtmlTagRegex =
        new(@"<[^>]+>", RegexOptions.Compiled);

    private static readonly Regex HorizontalWhitespaceRegex =
        new(@"[ \t]{2,}", RegexOptions.Compiled);

    private static readonly Regex ExcessiveNewlineRegex =
        new(@"\n{3,}", RegexOptions.Compiled);

    // Strips the quoted original-message thread from a supplier reply body before sending
    // to AI. Outlook appends "-----Original Message-----" when replying inline; without
    // stripping, Claude reads the RFQ product list from the original message and creates
    // ghost rows for all requested items even when the supplier provided no prices.
    private static readonly Regex QuotedThreadRegex =
        new(@"[\r\n][\s]*-{4,}[\s]*[Oo]riginal [Mm]essage[\s]*-{4,}", RegexOptions.Compiled);

    // Sliding-window rate limiter — tracks timestamps of recent AI calls.
    private readonly Queue<DateTimeOffset> _aiCallTimestamps = new();
    private readonly SemaphoreSlim         _rateLimitLock    = new(1, 1);

    // Signalled by TriggerReprocessAllAsync to request an immediate full scan.
    private readonly SemaphoreSlim _reprocessTrigger = new(0, 1);

    // ── Observable status (read by MailStatusController) ─────────────────────

    // Reprocess batch progress (reset at start of each batch).
    private volatile int  _reprocessTotal     = 0;
    private volatile int  _reprocessCompleted = 0;
    private volatile int  _reprocessFailed    = 0;
    private volatile bool _reprocessActive    = false;

    // Messages currently being processed (messageId → subject + from + start time).
    private readonly System.Collections.Concurrent.ConcurrentDictionary<string, (string Subject, string From, DateTimeOffset StartedAt)>
        _inFlight = new();

    // Application shutdown token — set once by ExecuteAsync so reprocess batches survive
    // HTTP client disconnections (which cancel the per-request CancellationToken).
    private CancellationToken _shutdownToken = CancellationToken.None;

    // Last poll cycle summary.
    private volatile int        _lastPollFound   = 0;
    private DateTimeOffset?     _lastPollAt      = null;

    /// <summary>Snapshot of current processing status for the /api/mail/status endpoint.</summary>
    public MailStatus GetStatus()
    {
        int callsInWindow;
        int maxPerMinute;
        _rateLimitLock.Wait();
        try
        {
            var cutoff = DateTimeOffset.UtcNow - TimeSpan.FromMinutes(1);
            while (_aiCallTimestamps.Count > 0 && _aiCallTimestamps.Peek() <= cutoff)
                _aiCallTimestamps.Dequeue();
            callsInWindow = _aiCallTimestamps.Count;
            maxPerMinute  = int.TryParse(_config["Mail:MaxEmailsPerMinute"], out var r) ? Math.Max(1, r) : 100;
        }
        finally { _rateLimitLock.Release(); }

        return new MailStatus(
            Poller: new PollerStatus(
                Running:                true,
                LastPollAt:             _lastPollAt,
                MessagesFoundLastCycle: _lastPollFound),
            Reprocess: new ReprocessStatus(
                Active:          _reprocessActive,
                Total:           _reprocessTotal,
                Completed:       _reprocessCompleted,
                Failed:          _reprocessFailed,
                PercentComplete: _reprocessTotal > 0
                                 ? Math.Round(_reprocessCompleted * 100.0 / _reprocessTotal, 1)
                                 : 0),
            RateLimit: new RateLimitStatus(
                CallsInLastMinute: callsInWindow,
                MaxPerMinute:      maxPerMinute,
                SlotsAvailable:    Math.Max(0, maxPerMinute - callsInWindow)),
            InFlight: _inFlight
                .Select(kv => new InFlightItem(
                    kv.Value.Subject,
                    kv.Value.From,
                    kv.Value.StartedAt.ToLocalTime().ToString("HH:mm:ss")))
                .ToList());
    }

    /// <summary>
    /// Reprocesses a specific set of already-processed messages by fetching each one from
    /// Graph and running the full extraction pipeline (AI + SharePoint write + re-stamp).
    /// Called synchronously by the reprocess-selected endpoint; awaited before responding.
    /// </summary>
    public async Task ReprocessMessagesAsync(IEnumerable<string> messageIds, CancellationToken ct)
    {
        var mailbox             = _config["Mail:MailboxAddress"]
            ?? throw new InvalidOperationException("Mail:MailboxAddress not configured");
        var maxPerMinute        = int.TryParse(_config["Mail:MaxEmailsPerMinute"],        out var r)  ? Math.Max(1, r) : 100;
        var maxConcurrency      = int.TryParse(_config["Mail:MaxConcurrency"],            out var mc) ? Math.Max(1, mc) : 8;
        var bodyContextChars    = int.TryParse(_config["Mail:BodyContextChars"],          out var bc) ? bc             : 2_000;
        var extractBodyNoJobRef = bool.TryParse(_config["Mail:ExtractBodyWithoutJobRef"], out var eb) && eb;

        var idList = messageIds.ToList();
        _reprocessTotal     = idList.Count;
        _reprocessCompleted = 0;
        _reprocessFailed    = 0;
        _reprocessActive    = true;
        _log.LogInformation("[Reprocess] Starting batch of {Total} message(s)", idList.Count);

        try
        {
            // Use the app shutdown token (not the HTTP request token) so a client
            // disconnect doesn't abort in-progress extraction and SP writes.
            var batchCt = _shutdownToken;
            await Parallel.ForEachAsync(idList,
                new ParallelOptions { MaxDegreeOfParallelism = maxConcurrency, CancellationToken = batchCt },
                async (messageId, _ct) =>
                {
                    var msg = await _mail.GetMessageByIdAsync(mailbox, messageId);
                    if (msg is null)
                    {
                        _log.LogWarning("[Reprocess] Message {Id} not found — skipping", messageId);
                        Interlocked.Increment(ref _reprocessFailed);
                        return;
                    }

                    _log.LogInformation("[Reprocess] Reprocessing \"{Subject}\" from {From}",
                        msg.Subject, msg.From?.EmailAddress?.Address);
                    try
                    {
                        await ProcessMessageAsync(mailbox, msg, maxPerMinute, bodyContextChars, extractBodyNoJobRef, _ct);
                        Interlocked.Increment(ref _reprocessCompleted);
                    }
                    catch (Exception ex) when (!_ct.IsCancellationRequested)
                    {
                        _log.LogError(ex, "[Reprocess] Failed processing {Id}", messageId);
                        Interlocked.Increment(ref _reprocessFailed);
                    }
                });
        }
        finally
        {
            _reprocessActive = false;
            _log.LogInformation("[Reprocess] Batch complete — {Completed}/{Total} succeeded, {Failed} failed",
                _reprocessCompleted, _reprocessTotal, _reprocessFailed);
        }
    }

    /// <summary>
    /// Re-runs AI extraction on each PO record
    /// LineItems JSON in SharePoint, uploads the PDF if not already stored, then re-runs the
    /// RLI matching logic.
    /// <para>
    /// <paramref name="comFetcher"/> is required for PO records that were sourced from Outlook COM
    /// (their MessageId is a MAPI EntryID, not a Graph message ID). Pass
    /// <see cref="OutlookComPollerService.FetchByEntryIdAsync"/> when available.
    /// </para>
    /// </summary>
    public async Task<(int Updated, int Skipped)> ReextractPoLineItemsAsync(
        Func<string, Task<OutlookPoMessage?>>? comFetcher,
        CancellationToken ct)
    {
        var mailbox = _config["Mail:MailboxAddress"]
            ?? throw new InvalidOperationException("Mail:MailboxAddress not configured");

        var records = await _sp.ReadPurchaseOrdersAsync();
        int updated = 0, skipped = 0;

        foreach (var po in records)
        {
            if (ct.IsCancellationRequested) break;

            if (string.IsNullOrEmpty(po.MessageId) || string.IsNullOrEmpty(po.SpItemId))
            {
                _log.LogInformation("[POReextract] Skipping PO {PoNumber} — no MessageId or SpItemId", po.PoNumber);
                skipped++;
                continue;
            }

            // MAPI EntryIDs (from Outlook COM) are long uppercase hex strings; Graph IDs are base64url.
            bool isComRecord = IsMApiEntryId(po.MessageId);

            string body;
            string subject;
            List<(string Name, byte[] Bytes)> pdfAttachments;

            if (isComRecord)
            {
                if (comFetcher is null)
                {
                    _log.LogWarning("[POReextract] COM-sourced PO {PoNumber} but Outlook COM is unavailable — skipping", po.PoNumber);
                    skipped++;
                    continue;
                }

                var comMsg = await comFetcher(po.MessageId);
                if (comMsg is null)
                {
                    _log.LogWarning("[POReextract] Could not fetch COM message for PO {PoNumber} — Outlook may not be running", po.PoNumber);
                    skipped++;
                    continue;
                }

                subject        = comMsg.Subject;
                body           = comMsg.PlainBody;
                pdfAttachments = comMsg.PdfAttachments;
            }
            else
            {
                var msg = await _mail.GetMessageByIdAsync(mailbox, po.MessageId);
                if (msg is null)
                {
                    _log.LogWarning("[POReextract] Graph message {MessageId} not found for PO {PoNumber} — skipping", po.MessageId, po.PoNumber);
                    skipped++;
                    continue;
                }

                var rawBody = msg.Body?.Content ?? string.Empty;
                if (msg.Body?.ContentType == BodyType.Html)
                {
                    rawBody = HtmlLineBreakRegex.Replace(rawBody, "\n");
                    rawBody = HtmlTagRegex     .Replace(rawBody, " ");
                    rawBody = System.Net.WebUtility.HtmlDecode(rawBody);
                }
                body    = rawBody.Trim();
                subject = msg.Subject ?? "";

                pdfAttachments = [];
                if (msg.HasAttachments == true)
                {
                    var attachments = await _mail.GetAttachmentsAsync(mailbox, po.MessageId);
                    foreach (var att in attachments)
                    {
                        if (att is not FileAttachment fa) continue;
                        if (fa.ContentBytes is null) continue;
                        var ct2 = (fa.ContentType ?? "").ToLowerInvariant();
                        var ext2 = Path.GetExtension(fa.Name ?? "").ToLowerInvariant();
                        if (!ct2.Contains("pdf") && ext2 != ".pdf") continue;
                        pdfAttachments.Add((fa.Name ?? "po.pdf", fa.ContentBytes));
                        break;
                    }
                }
            }

            var jobRefs = JobRefRegex.Matches(subject + " " + body)
                .Select(m => m.Groups[1].Value.ToUpperInvariant())
                .Distinct()
                .ToList();

            // Re-run AI extraction on the first PDF attachment
            PoExtraction? extraction = null;
            foreach (var (name, bytes) in pdfAttachments)
            {
                extraction = await _aiFactory.GetService().ExtractPurchaseOrderAsync(
                    Convert.ToBase64String(bytes),
                    name,
                    body[..Math.Min(body.Length, 1_000)],
                    subject,
                    jobRefs,
                    ct);
                if (extraction is not null) break;
            }

            if (extraction is null)
            {
                _log.LogWarning("[POReextract] No PDF or extraction failed for PO {PoNumber} — skipping", po.PoNumber);
                skipped++;
                continue;
            }

            var lineItemsJson = System.Text.Json.JsonSerializer.Serialize(
                extraction.LineItems,
                new System.Text.Json.JsonSerializerOptions { PropertyNamingPolicy = System.Text.Json.JsonNamingPolicy.CamelCase });

            await _sp.UpdatePurchaseOrderLineItemsAsync(po.SpItemId, lineItemsJson);
            _log.LogInformation("[POReextract] Updated LineItems for PO {PoNumber} ({Count} items)", po.PoNumber, extraction.LineItems.Count);

            // Upload the PDF to SharePoint if not already stored
            if (pdfAttachments.Count > 0 && string.IsNullOrEmpty(po.PdfUrl))
            {
                var (pdfName, pdfBytes) = pdfAttachments[0];
                var pdfUrl = await _sp.UploadPoAttachmentAsync(
                    po.SpItemId,
                    extraction.PoNumber ?? po.PoNumber ?? po.SpItemId,
                    pdfName,
                    pdfBytes);
                _log.LogInformation("[POReextract] Uploaded PDF for PO {PoNumber}: {Url}", po.PoNumber, pdfUrl);
            }

            // Re-run RLI matching with fresh line items
            var rfqId = po.RfqId is "UNKNOWN" or "" ? null : po.RfqId;
            if (!string.IsNullOrEmpty(rfqId))
                await _sp.UpdateRliPurchaseStatusAsync(rfqId, po.SupplierName, po.SpItemId, extraction.LineItems, extraction.PoNumber ?? po.PoNumber);
            else
                await _sp.MatchAndMarkRliByMspcAsync(po.SupplierName, po.PoNumber, extraction.LineItems);

            updated++;
        }

        return (updated, skipped);
    }

    /// <summary>
    /// Returns true if <paramref name="messageId"/> looks like a MAPI EntryID (Outlook COM source)
    /// rather than a Microsoft Graph message ID.
    /// MAPI EntryIDs are long (64+ chars) strings composed entirely of hexadecimal digits.
    /// Graph IDs are base64url strings that typically start with letters and contain hyphens/underscores.
    /// </summary>
    private static bool IsMApiEntryId(string messageId) =>
        messageId.Length >= 64 &&
        messageId.All(c => (c >= '0' && c <= '9') || (c >= 'A' && c <= 'F') || (c >= 'a' && c <= 'f'));

    private static string StripQuotedThread(string body)
    {
        var m = QuotedThreadRegex.Match(body);
        return m.Success ? body[..m.Index].TrimEnd() : body;
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
        AiServiceFactory           aiFactory,
        SharePointService          sp,
        ProductCatalogService      catalog,
        RfqNotificationService     notifications,
        RfqSummaryQueue            summaryQueue,
        SliCacheService            sliCache,
        ShrConvInRouter            shrRouter,
        CatalogAnalysisService     analysis,
        ILogger<MailPollerService> log)
    {
        _config        = config;
        _mail          = mail;
        _aiFactory     = aiFactory;
        _sp            = sp;
        _catalog       = catalog;
        _notifications = notifications;
        _summaryQueue  = summaryQueue;
        _sliCache      = sliCache;
        _shrRouter     = shrRouter;
        _analysis      = analysis;
        _log           = log;
    }

    protected override async Task ExecuteAsync(CancellationToken stoppingToken)
    {
        _shutdownToken = stoppingToken;
        var mailbox         = _config["Mail:MailboxAddress"]
            ?? throw new InvalidOperationException(
                "Mail:MailboxAddress is not configured. " +
                "Add it to User Secrets or appsettings.Development.json.");

        var intervalSeconds          = int.TryParse(_config["Mail:PollIntervalSeconds"],      out var s)  ? s          : 30;
        var lookbackHours            = double.TryParse(_config["Mail:LookbackHours"],          out var h)  ? h          : 24.0;
        var maxPerMinute             = int.TryParse(_config["Mail:MaxEmailsPerMinute"],        out var r)  ? Math.Max(1, r) : 100;
        var maxConcurrency           = int.TryParse(_config["Mail:MaxConcurrency"],            out var mc) ? Math.Max(1, mc) : 8;
        var bodyContextChars         = int.TryParse(_config["Mail:BodyContextChars"],          out var bc) ? bc         : 2_000;
        var extractBodyWithoutJobRef = bool.TryParse(_config["Mail:ExtractBodyWithoutJobRef"], out var eb) && eb;
        var interval                 = TimeSpan.FromSeconds(intervalSeconds);

        _log.LogInformation(
            "[Mail] Poller started — mailbox={Mailbox} interval={Interval}s lookback={Lookback}h rateLimit={Rate}/min concurrency={Concurrency} extractBodyNoRef={ExtractNoRef}",
            mailbox, intervalSeconds, lookbackHours, maxPerMinute, maxConcurrency, extractBodyWithoutJobRef);

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
                await PollAsync(mailbox, since, maxPerMinute, maxConcurrency, bodyContextChars, extractBodyWithoutJobRef, stoppingToken);
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
        string mailbox, DateTimeOffset since, int maxPerMinute, int maxConcurrency,
        int bodyContextChars, bool extractBodyWithoutJobRef, CancellationToken ct)
    {
        var parallelOpts = new ParallelOptions { MaxDegreeOfParallelism = maxConcurrency, CancellationToken = ct };

        _lastPollAt = DateTimeOffset.UtcNow;
        var messages = await _mail.GetMessagesAsync(mailbox, since);
        _lastPollFound = messages.Count;

        if (messages.Count == 0)
            _log.LogDebug("[Mail] No unprocessed inbox messages");
        else
        {
            _log.LogInformation("[Mail] {Count} unprocessed inbox message(s) found", messages.Count);
            await Parallel.ForEachAsync(messages, parallelOpts,
                async (msg, _ct) => await ProcessMessageAsync(mailbox, msg, maxPerMinute, bodyContextChars, extractBodyWithoutJobRef, _ct));
        }

        // Also scan Sent Items for outbound PO emails not yet processed.
        if (ct.IsCancellationRequested) return;
        var sentPo = await _mail.GetUnprocessedSentPoMessagesAsync(mailbox, since);
        if (sentPo.Count > 0)
        {
            _log.LogInformation("[Mail] {Count} unprocessed sent PO email(s) found", sentPo.Count);
            await Parallel.ForEachAsync(sentPo, parallelOpts,
                async (msg, _ct) => await ProcessMessageAsync(mailbox, msg, maxPerMinute, bodyContextChars, extractBodyWithoutJobRef, _ct));
        }
    }

    // ── Rate limiter ──────────────────────────────────────────────────────────

    /// <summary>
    /// Waits until an AI API call is allowed under the sliding-window rate limit.
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
                while (_aiCallTimestamps.Count > 0 && _aiCallTimestamps.Peek() <= cutoff)
                    _aiCallTimestamps.Dequeue();

                if (_aiCallTimestamps.Count < maxPerMinute)
                {
                    _aiCallTimestamps.Enqueue(now);
                    return;   // slot acquired
                }

                // Calculate exactly how long until the oldest call leaves the window.
                var waitUntil = _aiCallTimestamps.Peek() + window;
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
                while (_aiCallTimestamps.Count > 0 && _aiCallTimestamps.Peek() <= cutoff)
                    _aiCallTimestamps.Dequeue();
                sleepFor = _aiCallTimestamps.Count >= maxPerMinute
                    ? (_aiCallTimestamps.Peek() + window - now2 + TimeSpan.FromMilliseconds(50))
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

        // Track in-flight for /api/mail/status
        var trackId = msg.Id ?? Guid.NewGuid().ToString();
        _inFlight[trackId] = (subject, fromAddr, DateTimeOffset.UtcNow);

        // Attach email context to every log line emitted while processing this message.
        using var logScope = _log.BeginScope(new Dictionary<string, object?>
        {
            ["EmailSubject"] = subject,
            ["EmailFrom"]    = fromAddr,
            ["MessageId"]    = msg.Id,
        });

        try
        {

        // Claim this message immediately so other proxy instances skip it on their next poll.
        if (msg.Id is not null)
            await _mail.MarkClaimingAsync(mailbox, msg.Id);

        // PO documents are sourced directly from the ERP via FileWatcherService.
        // Email-based PO routing removed — see FileWatcherService.TriggerPoMatchingAsync.

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

        var searchText = subject + " " + body;
        var jobRefs = JobRefRegex.Matches(searchText)
            .Select(m => m.Groups[1].Value.ToUpperInvariant())
            .Distinct()
            .ToList();

        // Fallback: supplier replies sometimes strip the brackets. Try the bare-ID pattern
        // (e.g. "RFQ M5S4OR" or "Job Ref: M5S4OR") when the bracketed form isn't found.
        if (jobRefs.Count == 0)
        {
            jobRefs = JobRefBareRegex.Matches(searchText)
                .Select(m => (m.Groups[1].Success ? m.Groups[1] : m.Groups[2]).Value.ToUpperInvariant())
                .Distinct()
                .ToList();
            if (jobRefs.Count > 0)
                _log.LogInformation("[Mail] Job ref found via bare pattern (no brackets): [{Refs}] in \"{Subject}\"",
                    string.Join(", ", jobRefs), subject);
        }

        // ── SHR conversation tracking token ──────────────────────────────────────
        // ShrConvInRouter writes to SupplierConversations when the supplier resolves.
        // We never early-return on Routed: the token appears in quoted reply bodies
        // of initial RFQ emails so short-circuiting would silently drop pricing rows.
        var shrResult = await _shrRouter.TryRouteAsync(
            searchText:          searchText,
            fromAddr:            fromAddr,
            subject:             subject,
            body:                body,
            messageId:           msg.Id,
            hasAttachments:      msg.HasAttachments == true,
            receivedAt:          msg.ReceivedDateTime ?? DateTimeOffset.UtcNow,
            graphConversationId: msg.ConversationId);

        // Seed the rfqId so AI extraction files the row under the correct RFQ.
        if (shrResult.ShrRfqId is not null &&
            !jobRefs.Contains(shrResult.ShrRfqId, StringComparer.OrdinalIgnoreCase))
        {
            jobRefs.Insert(0, shrResult.ShrRfqId);
        }

        // Decide whether to call the AI.
        // AI is always used when a job reference is present or there is an attachment
        // (which may be a quote PDF/DOCX even without a reference in the body).
        // Body-only emails with no job reference bypass AI and get a direct placeholder row
        // unless ExtractBodyWithoutJobRef is explicitly enabled in config.
        // Exception: MSG conv replies are fully handled by ShrConvInRouter — skip the SLI path.
        bool hasJobRef    = jobRefs.Count > 0;
        bool hasAttachment = msg.HasAttachments == true;
        bool sendToAi = hasJobRef || hasAttachment || extractBodyWithoutJobRef;

        if (shrResult.Routed && shrResult.IsMsgConv)
        {
            _log.LogInformation("[SHR] MSG conv reply from {From} written to SupplierConversations — skipping SLI", fromAddr);
            await _mail.MarkProcessedAsync(mailbox, msg.Id!, "RFQ-Processed");
            return;
        }

        if (!hasJobRef && !hasAttachment)
            _log.LogInformation("[Mail] No job ref or attachment in \"{Subject}\" from {From} — {Action}",
                subject, fromAddr, sendToAi ? "sending to AI" : "writing direct row under [000000]");
        else
            _log.LogInformation("[Mail] Processing: \"{Subject}\" from {From} refs=[{Refs}] attachments={Att}",
                subject, fromAddr, string.Join(", ", jobRefs), hasAttachment);

        var bodySnippet = body[..Math.Min(body.Length, bodyContextChars)];

        bool supplierUnknown = false;

        if (!sendToAi)
        {
            // Fast path: no job ref and no attachment — write a placeholder row directly
            // without spending an AI API call. SP will assign [000000] or [WHOIS] as appropriate.
            var req = new ExtractRequest
            {
                Content      = string.Empty,
                EmailBody    = body,
                SourceType   = "body",
                JobRefs      = jobRefs,
                EmailSubject = subject,
                EmailFrom            = fromAddr,
                ReceivedAt           = received,
                GraphConversationId  = msg.ConversationId,
            };
            var extraction = new RfqExtraction();
            var placeholder = new ProductLine
            {
                SupplierProductComments = "Email recorded without extraction — no job reference or quote attachment detected."
            };
            var row = await _sp.WriteProductRowAsync(extraction, placeholder, req, "body", null, 0, msg.Id);
            supplierUnknown = row.SupplierUnknown;
            if (row.Success && !row.Updated) _notifications.NotifyRfqProcessed();
        }
        else if (!hasAttachment || msg.Id is null)
        {
            // No attachments — extract pricing from body; always write at least one row.
            var extractBody = StripQuotedThread(body);
            supplierUnknown = await RunExtractionAsync(new ExtractRequest
            {
                Content                = extractBody[..Math.Min(extractBody.Length, 12_000)],
                EmailBody              = body,
                SourceType             = "body",
                JobRefs                = jobRefs,
                EmailSubject           = subject,
                EmailFrom              = fromAddr,
                ReceivedAt             = received,
                HasAttachment          = false,
                ResolvedSupplierName   = shrResult.ResolvedSupplier,
                GraphConversationId    = msg.ConversationId,
            }, "body", null, maxPerMinute, ct, msg.Id);
        }
        else
        {
            // Has attachments — prefer attachment data for pricing; skip body-only extraction
            // so only one row is written per email.  Body is stored in EmailBody column.
            var attachments = await _mail.GetAttachmentsAsync(mailbox, msg.Id);
            bool processedAny = false;

            _log.LogInformation("[Mail] Fetched {Count} attachment(s) for message {Id}",
                attachments.Count, msg.Id);

            foreach (var att in attachments)
            {
                if (ct.IsCancellationRequested) break;
                if (att is not FileAttachment fa)
                {
                    _log.LogDebug("[Mail] Skipping non-file attachment: {Type}", att.GetType().Name);
                    continue;
                }

                var contentType = (fa.ContentType ?? "").ToLowerInvariant();
                var fileName = fa.Name ?? "(unnamed)";
                var ext = Path.GetExtension(fileName).ToLowerInvariant();

                bool isSupportedType = contentType.Contains("pdf")
                    || contentType.Contains("wordprocessingml")
                    || contentType.Contains("msword")
                    || ext is ".pdf" or ".docx" or ".doc";

                if (!isSupportedType)
                {
                    _log.LogInformation("[Mail] Skipping attachment '{Name}' — unsupported content type: {ContentType}",
                        fileName, contentType);
                    continue;
                }

                // Normalise content type from file extension when the MIME type is generic
                if (!contentType.Contains("pdf") && ext == ".pdf")
                {
                    _log.LogInformation("[Mail] Overriding content type for '{Name}': '{Original}' → 'application/pdf'",
                        fileName, contentType);
                    contentType = "application/pdf";
                }

                if (fa.ContentBytes is null)
                {
                    _log.LogWarning("[Mail] Attachment '{Name}' has null ContentBytes (type={ContentType}, size={Size})",
                        fileName, contentType, fa.Size);
                    continue;
                }

                _log.LogInformation("[Mail] Processing attachment '{Name}' ({ContentType}, {Bytes} bytes)",
                    fileName, contentType, fa.ContentBytes.Length);

                supplierUnknown = await RunExtractionAsync(new ExtractRequest
                {
                    Content              = string.Empty,
                    EmailBody            = body,
                    SourceType           = "attachment",
                    FileName             = fa.Name,
                    ContentType          = contentType,
                    Base64Data           = Convert.ToBase64String(fa.ContentBytes),
                    BodyContext          = bodySnippet,
                    JobRefs              = jobRefs,
                    EmailSubject         = subject,
                    EmailFrom            = fromAddr,
                    ReceivedAt           = received,
                    HasAttachment        = true,
                    ResolvedSupplierName = shrResult.ResolvedSupplier,
                    GraphConversationId  = msg.ConversationId,
                }, "attachment", fa.Name, maxPerMinute, ct, msg.Id);

                processedAny = true;
            }

            // No recognisable attachment format — fall back to body extraction.
            if (!processedAny && !ct.IsCancellationRequested)
            {
                var extractBody = StripQuotedThread(body);
                supplierUnknown = await RunExtractionAsync(new ExtractRequest
                {
                    Content              = extractBody[..Math.Min(extractBody.Length, 12_000)],
                    EmailBody            = body,
                    SourceType           = "body",
                    JobRefs              = jobRefs,
                    EmailSubject         = subject,
                    EmailFrom            = fromAddr,
                    ReceivedAt           = received,
                    HasAttachment        = true,
                    ResolvedSupplierName = shrResult.ResolvedSupplier,
                    GraphConversationId  = msg.ConversationId,
                }, "body", null, maxPerMinute, ct, msg.Id);
            }
        }

        if (msg.Id is not null)
        {
            var extra = supplierUnknown ? "Unknown" : null;
            if (supplierUnknown)
                _log.LogInformation("[Mail] Supplier unrecognised in \"{Subject}\" — stamping 'Unknown' category", subject);
            await _mail.MarkProcessedAsync(mailbox, msg.Id, extra);
        }

        } // end try
        finally { _inFlight.TryRemove(trackId, out _); }
    }

    // ── Purchase Order processing ─────────────────────────────────────────────

    private async Task ProcessPurchaseOrderAsync(
        string mailbox, Message msg, string subject, CancellationToken ct)
    {
        var fromAddr = msg.From?.EmailAddress?.Address ?? "unknown";
        var received = msg.ReceivedDateTime?.ToString("o") ?? DateTime.UtcNow.ToString("o");

        // Strip HTML from body
        var rawBody = msg.Body?.Content ?? string.Empty;
        if (msg.Body?.ContentType == BodyType.Html)
        {
            rawBody = HtmlLineBreakRegex.Replace(rawBody, "\n");
            rawBody = HtmlTagRegex     .Replace(rawBody, " ");
            rawBody = System.Net.WebUtility.HtmlDecode(rawBody);
            rawBody = HorizontalWhitespaceRegex.Replace(rawBody, " ");
            rawBody = ExcessiveNewlineRegex    .Replace(rawBody, "\n\n");
        }
        else rawBody = System.Net.WebUtility.HtmlDecode(rawBody);

        // Fetch PDF attachments via Graph
        var pdfs = new List<(string Name, byte[] Bytes)>();
        if (msg.HasAttachments == true && msg.Id is not null)
        {
            var attachments = await _mail.GetAttachmentsAsync(mailbox, msg.Id);
            foreach (var att in attachments)
            {
                if (att is not FileAttachment fa) continue;
                if (fa.ContentBytes is null) continue;
                var poContentType = (fa.ContentType ?? "").ToLowerInvariant();
                var poExt = Path.GetExtension(fa.Name ?? "").ToLowerInvariant();
                if (!poContentType.Contains("pdf") && poExt != ".pdf") continue;
                pdfs.Add((fa.Name ?? "po.pdf", fa.ContentBytes));
                break;
            }
        }

        var processed = await ProcessPurchaseOrderCoreAsync(
            subject, rawBody.Trim(), pdfs, msg.Id, received, fromAddr, ct);

        if (msg.Id is not null)
            await _mail.MarkProcessedAsync(mailbox, msg.Id, processed ? "PO-Processed" : "PO-NoExtract");
    }

    /// <summary>
    /// Core PO processing logic shared by the Graph poller and the Outlook COM poller.
    /// Accepts plain-data inputs so it has no dependency on Microsoft.Graph types.
    /// Returns true if extraction succeeded and the record was written.
    /// </summary>
    internal async Task<bool> ProcessPurchaseOrderCoreAsync(
        string subject,
        string plainTextBody,
        List<(string Name, byte[] Bytes)> pdfAttachments,
        string? messageId,
        string receivedAt,
        string senderHint,
        CancellationToken ct)
    {
        _log.LogInformation("[PO] Processing purchase order: \"{Subject}\" from {From}", subject, senderHint);

        var jobRefs = JobRefRegex.Matches(subject + " " + plainTextBody)
            .Select(m => m.Groups[1].Value.ToUpperInvariant())
            .Distinct()
            .ToList();

        PoExtraction? extraction = null;
        foreach (var (name, bytes) in pdfAttachments)
        {
            _log.LogInformation("[PO] Sending PDF '{File}' to AI for extraction", name);
            extraction = await _aiFactory.GetService().ExtractPurchaseOrderAsync(
                Convert.ToBase64String(bytes),
                name,
                plainTextBody[..Math.Min(plainTextBody.Length, 1_000)],
                subject,
                jobRefs,
                ct);
            break;
        }

        if (extraction is null)
        {
            _log.LogWarning("[PO] No PDF or extraction failed for \"{Subject}\" — skipping", subject);
            return false;
        }

        var rfqId = !string.IsNullOrWhiteSpace(extraction.JobReference)
            ? extraction.JobReference!.Trim().ToUpperInvariant()
            : jobRefs.FirstOrDefault();

        if (string.IsNullOrWhiteSpace(rfqId))
            _log.LogWarning("[PO] No job reference found in \"{Subject}\" — PO will be written without RFQ link", subject);

        var supplierName = extraction.SupplierName ?? senderHint;
        var lineItemsJson = System.Text.Json.JsonSerializer.Serialize(
            extraction.LineItems,
            new System.Text.Json.JsonSerializerOptions { PropertyNamingPolicy = System.Text.Json.JsonNamingPolicy.CamelCase });

        var poSpItemId = await _sp.WritePurchaseOrderAsync(
            rfqId ?? "UNKNOWN", supplierName, extraction.PoNumber,
            receivedAt, messageId, lineItemsJson);

        if (poSpItemId is not null)
        {
            // Upload the source PDF to SharePoint drive and store the URL on the PO record
            var firstPdf = pdfAttachments.FirstOrDefault();
            if (firstPdf != default)
                await _sp.UploadPoAttachmentAsync(poSpItemId, extraction.PoNumber ?? poSpItemId, firstPdf.Name, firstPdf.Bytes);

            if (!string.IsNullOrWhiteSpace(rfqId))
                await _sp.UpdateRliPurchaseStatusAsync(rfqId, supplierName, poSpItemId, extraction.LineItems, extraction.PoNumber);
            else
            {
                var matched = await _sp.MatchAndMarkRliByMspcAsync(supplierName, extraction.PoNumber, extraction.LineItems);
                // If MSPC matching found the RFQ, use it so the UI notification carries a real ID.
                if (rfqId is null && matched.Count > 0)
                    rfqId = matched.First();
            }
        }

        var notification = new Models.RfqProcessedNotification
        {
            EventType    = "PO",
            RfqId        = rfqId,
            SupplierName = supplierName,
            MessageId    = messageId,
            Products     = extraction.LineItems.Select(li => new Models.RfqNotificationProduct
            {
                Name = li.Product,
                Mspc = li.Mspc,
                Size = li.Size,
            }).ToList(),
        };
        _notifications.NotifyRfqProcessed(notification);

        _log.LogInformation("[PO] Processed: RfqId={RfqId}, Supplier={Supplier}, {Count} line item(s)",
            rfqId ?? "UNKNOWN", supplierName, extraction.LineItems.Count);

        return true;
    }

    // ── AI → SharePoint

    // Returns true if the supplier was not found in the reference list.
    private async Task<bool> RunExtractionAsync(
        ExtractRequest req, string source, string? fileName, int maxPerMinute, CancellationToken ct,
        string? messageId = null)
    {
        await AcquireRateSlotAsync(maxPerMinute, ct);

        bool dedupDryRun = bool.TryParse(_config["Dedup:DryRun"], out var dr) && dr;

        // Inject RLI requested-items context so the AI can anchor each extracted supplier
        // product to the nearest requested item and return its MSPC as productSearchKey.
        if (req.RliItems.Count == 0)
        {
            var validJobRef = req.JobRefs
                .Select(r => r.Trim('[', ']'))
                .FirstOrDefault(r => !string.IsNullOrEmpty(r) && r != "000000" && r != "WHOIS");

            if (validJobRef is not null)
            {
                try
                {
                    var rliItems = await _sp.ReadRfqLineItemsByRfqIdAsync(validJobRef);
                    if (rliItems.Count > 0)
                    {
                        req.RliItems = rliItems;
                        _log.LogInformation("[Mail] RLI anchoring: {Count} item(s) for [{RfqId}]",
                            rliItems.Count, validJobRef);
                    }
                    else
                    {
                        _log.LogDebug("[Mail] No RLI items found for [{RfqId}] — extraction without anchoring", validJobRef);
                    }
                }
                catch (Exception ex)
                {
                    _log.LogWarning(ex, "[Mail] RLI fetch failed for [{RfqId}] — continuing without anchoring", validJobRef);
                }
            }
        }

        try
        {
            var extraction = await _aiFactory.GetService().ExtractRfqAsync(req);
            _log.LogInformation("[Mail] Extracted: supplier={Supplier} quoteRef={QuoteRef}",
                extraction?.SupplierName, extraction?.QuoteReference ?? "(none)");

            // Late RLI inject: the subject had no job ref so we couldn't pre-fetch RLI,
            // but the AI extracted one from the email body. If any product lacks an MSPC,
            // fetch RLI now and re-run extraction so the AI can anchor against catalog items.
            if (req.RliItems.Count == 0
                && extraction?.JobReference is string lateJobRef
                && lateJobRef != "000000" && lateJobRef != "WHOIS"
                && extraction.Products.Any(p => string.IsNullOrEmpty(p.ProductSearchKey)))
            {
                try
                {
                    var rliItems = await _sp.ReadRfqLineItemsByRfqIdAsync(lateJobRef);
                    if (rliItems.Count > 0)
                    {
                        _log.LogInformation(
                            "[Mail] Late RLI inject: re-extracting with {Count} item(s) for [{RfqId}]",
                            rliItems.Count, lateJobRef);
                        req.RliItems = rliItems;
                        await AcquireRateSlotAsync(maxPerMinute, ct);
                        extraction = await _aiFactory.GetService().ExtractRfqAsync(req);
                        _log.LogInformation("[Mail] Late re-extract: supplier={Supplier} quoteRef={QuoteRef}",
                            extraction?.SupplierName, extraction?.QuoteReference ?? "(none)");
                    }
                }
                catch (Exception ex)
                {
                    _log.LogWarning(ex, "[Mail] Late RLI inject failed for [{RfqId}] — keeping first-pass result", lateJobRef);
                }
            }

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
                products = ProductDeduplicator.Deduplicate(products, source, dedupDryRun, _log);

                // Validate and warn when RLI context was provided.
                if (req.RliItems.Count > 0)
                {
                    var validMspcs = req.RliItems
                        .Where(r => !string.IsNullOrEmpty(r.Mspc))
                        .Select(r => r.Mspc!)
                        .ToHashSet(StringComparer.OrdinalIgnoreCase);

                    // Null out productSearchKey only when the AI hallucinated an MSPC —
                    // i.e., it's not in the catalog AND it wasn't one of the RLI MSPCs we
                    // provided as context. If the AI returned an MSPC from our RLI context,
                    // it's not hallucinating; trust it even if the catalog doesn't have it
                    // (the catalog may be incomplete — the MSPC came from the user's RFQ).
                    foreach (var p in products.Where(p => !string.IsNullOrEmpty(p.ProductSearchKey)))
                    {
                        var inCatalog = _catalog.FindBySearchKey(p.ProductSearchKey) is not null;
                        var inRli     = validMspcs.Contains(p.ProductSearchKey!);

                        if (!inCatalog && !inRli)
                        {
                            _log.LogWarning(
                                "[Mail] RLI drift: AI returned productSearchKey '{Key}' for '{Name}' " +
                                "which is not a known catalog MSPC — nulling",
                                p.ProductSearchKey, p.ProductName);
                            p.ProductSearchKey = null;
                        }
                        else if (!inCatalog && inRli)
                        {
                            _log.LogWarning(
                                "[Mail] RLI MSPC not in catalog: AI returned '{Key}' for '{Name}' — " +
                                "trusting RLI match but CatalogProductName will be blank (catalog gap)",
                                p.ProductSearchKey, p.ProductName);
                        }
                        else if (!inRli)
                        {
                            _log.LogWarning(
                                "[Mail] RLI off-list: AI returned productSearchKey '{Key}' for '{Name}' " +
                                "outside the RLI set [{Rli}] — accepting (valid catalog entry)",
                                p.ProductSearchKey, p.ProductName,
                                string.Join(", ", validMspcs));
                        }
                    }

                    // Warn for each product the AI couldn't anchor despite context being available.
                    foreach (var p in products.Where(p => string.IsNullOrEmpty(p.ProductSearchKey)))
                        _log.LogWarning(
                            "[Mail] RLI unmatched: AI returned no productSearchKey for '{Name}' " +
                            "despite {Count} RLI item(s). RLI=[{Rli}]",
                            p.ProductName,
                            req.RliItems.Count,
                            string.Join(" | ", req.RliItems.Select(r =>
                                $"{r.Mspc ?? "(no mspc)"}={r.ProductName}")));
                }
            }

            // Conv-tracked replies with no pricing skip SLI writes — SupplierConversations already has the message.
            bool hasPricing = products.Any(p =>
                p.PricePerPound.HasValue || p.PricePerFoot.HasValue ||
                p.PricePerPiece.HasValue || p.TotalPrice.HasValue);
            if (!hasPricing && !string.IsNullOrEmpty(req.ResolvedSupplierName))
            {
                _log.LogInformation(
                    "[Mail] Conv-tracked reply from {From} has no pricing — skipping SLI write to preserve existing rows",
                    req.EmailFrom);
                return false;
            }

            // Body-only emails often carry only a unit price ($/ft, $/lb, each) without
            // repeating the quantity and size from the RFQ — the supplier assumes we know.
            // Enrich those product lines with RLI qty + size before writing so that
            // EffectiveTotalPrice and ProductKey are computed correctly on the Shredder side.
            if (source == "body" && req.RliItems.Count > 0)
                foreach (var p in products)
                    EnrichFromRli(p, req.RliItems);

            // Write product rows sequentially so they all share one SupplierResponse row.
            // Concurrent writes caused a race: every call independently found no existing SR
            // and created its own, producing duplicate SR records per multi-product email.
            var rowList = new List<SpWriteResult>(products.Count);
            foreach (var (p, i) in products.Select((p, i) => (p, i)))
            {
                if (ct.IsCancellationRequested) break;
                rowList.Add(await _sp.WriteProductRowAsync(extraction!, p, req, source, fileName, i, messageId));
            }
            var rows = rowList.ToArray();

            bool anyUnknown  = false;
            bool anyInserted = false;   // true only for new SP rows, not updates/reprocessing
            for (int i = 0; i < rows.Length; i++)
            {
                var row = rows[i];
                if (row.SupplierUnknown)
                    anyUnknown = true;
                else if (row.Success)
                {
                    if (!row.Updated) anyInserted = true;
                    _log.LogInformation("[Mail] SP row {Action}: '{Product}' -> {Url}",
                        row.Updated ? "updated" : "inserted", products[i].ProductName, row.SpWebUrl);
                }
                else
                    _log.LogWarning("[Mail] SP upsert failed for '{Product}': {Error}",
                        products[i].ProductName, row.Error);
            }

            // Fire token-match diagnostics in the background — zero impact on response time.
            for (int i = 0; i < rows.Length; i++)
            {
                var row = rows[i];
                if (row.Success && !row.SupplierUnknown && row.SliSpItemId is not null)
                {
                    var p = products[i];
                    _ = FireTokenMatchDiagnosticAsync(
                        p.ProductName ?? $"Product {i + 1}",
                        row.SupplierName,
                        row.RfqId,
                        row.SliSpItemId,
                        p.ProductSearchKey);
                }
            }

            // Notify the UI whenever a supplier response was successfully written — either
            // a new insert or an update to an existing row.  Startup rescans that produce
            // no changes (anyInserted=false, no successful updates) remain silent.
            var anyUpdated = rows.Any(r => r.Success && r.Updated && !r.SupplierUnknown);
            if (anyInserted || anyUpdated)
            {
                var notification = new Models.RfqProcessedNotification
                {
                    EventType    = "SR",
                    SupplierName = rows.Zip(products)
                                       .FirstOrDefault(x => x.First.Success && !x.First.SupplierUnknown)
                                       .First?.SupplierName,
                    // Prefer the SP-resolved RFQ ID (populated even when email subject has no bracket ref).
                    // Fall back to req.JobRefs for robustness.
                    RfqId = rows.FirstOrDefault(r => r.Success && !string.IsNullOrEmpty(r.RfqId))?.RfqId
                            ?? req.JobRefs.FirstOrDefault()?.Trim('[', ']'),
                    MessageId = messageId,
                    Products = rows.Zip(products)
                                   .Where(x => x.First.Success && !x.First.SupplierUnknown)
                                   .Select(x => new Models.RfqNotificationProduct
                                   {
                                       Name       = x.First.ProductName,
                                       TotalPrice = x.Second.TotalPrice,
                                   }).ToList(),
                };

                // Read back the complete current SLI rows for this RFQ and embed them in the
                // bus message.  Bus recipients (Shredder clients + peer proxies) apply these
                // rows inline, avoiding a proxy round-trip and the SP-index-lag window.
                // MergeRfqRows keeps the rest of the SliCache valid — no full invalidation needed.
                if (!string.IsNullOrEmpty(notification.RfqId))
                {
                    try
                    {
                        notification.SliRows = await _sp.ReadSupplierItemsByRfqIdAsync(notification.RfqId);
                        _sliCache.MergeRfqRows(notification.RfqId, notification.SliRows);
                        // Race-free state-of-play: enqueue a (re)generation for this RFQ. Dup-detection on
                        // {rfqId}:{inputsHash} collapses concurrent proxies; one consumer ever generates.
                        await _summaryQueue.MaybeEnqueueAsync(notification.RfqId, notification.SliRows);
                    }
                    catch (Exception ex)
                    {
                        _log.LogWarning(ex,
                            "[Mail] Failed to read SLI rows for bus message (rfqId={RfqId}) — falling back to full cache invalidate",
                            notification.RfqId);
                        _sliCache.Invalidate();
                    }
                }
                else
                {
                    _sliCache.Invalidate();
                }

                _notifications.NotifyRfqProcessed(notification);
            }

            // Log the inbound email to SupplierConversations so the thread viewer
            // can treat SC as the single source of truth. Dedup is handled by
            // WriteConversationMessageAsync — safe if the SHR bypass already wrote a row.
            var firstGood = rows.FirstOrDefault(r => r.Success && !r.SupplierUnknown);
            if (firstGood is not null)
            {
                var receivedAt = DateTimeOffset.TryParse(req.ReceivedAt, out var rt) ? rt : DateTimeOffset.UtcNow;
                await _shrRouter.WriteConvInFromExtractionAsync(
                    rfqId:               firstGood.RfqId,
                    supplierName:        firstGood.SupplierName,
                    messageId:           messageId,
                    subject:             req.EmailSubject,
                    body:                req.EmailBody,
                    receivedAt:          receivedAt,
                    hasAttachments:      req.HasAttachment,
                    fromAddr:            req.EmailFrom,
                    graphConversationId: req.GraphConversationId,
                    attachmentName:      req.SourceType == "attachment" ? req.FileName : null);
            }

            return anyUnknown;
        }
        catch (Exception ex)
        {
            _log.LogError(ex, "[Mail] Extraction failed for {Source} ({File})", source, fileName ?? "body");
            return false;
        }
    }

    private async Task FireTokenMatchDiagnosticAsync(
        string productName, string? supplierName, string? rfqId, string sliSpItemId, string? rliMspc)
    {
        try
        {
            var tokenResult = await _analysis.MatchProductAsync(productName, supplierName);

            var resolvedSource = !string.IsNullOrEmpty(rliMspc) ? "rli_anchor"
                               : tokenResult.Source == "user_mapping" ? "user_mapping"
                               : tokenResult.Source == "token_scorer" ? "token_scorer"
                               : "none";

            var agreed = !string.IsNullOrEmpty(rliMspc) && !string.IsNullOrEmpty(tokenResult.SearchKey)
                ? string.Equals(rliMspc, tokenResult.SearchKey, StringComparison.OrdinalIgnoreCase)
                : false; // can only agree when both sides resolved something

            await _sp.WriteTokenMatchDiagnosticAsync(new TokenMatchDiagnosticEntry
            {
                RfqId          = rfqId,
                SliSpItemId    = sliSpItemId,
                SupplierName   = supplierName ?? "",
                ProductName    = productName,
                ResolvedSource = resolvedSource,
                RliMspc        = rliMspc,
                TokenMspc      = tokenResult.SearchKey,
                TokenScore     = tokenResult.Score,
                Agreed         = agreed,
                ReviewStatus   = "pending",
                CreatedAt      = DateTime.UtcNow,
            });
        }
        catch (Exception ex)
        {
            _log.LogWarning(ex, "[Diagnostic] Token match diagnostic failed for '{Name}'", productName);
        }
    }

    /// <summary>
    /// Fills in missing UnitsRequested and LengthPerUnit on a body-extracted product line
    /// when the supplier gave only a unit price ($/ft, $/lb, each) and the matching RLI
    /// item carries the quantity and size we originally requested.
    /// Only fires when these fields are absent — never overwrites values Claude extracted.
    /// </summary>
    private static void EnrichFromRli(ProductLine p, List<RliContextItem> rliItems)
    {
        // Only enrich rows that have a unit price but no explicit qty
        bool hasUnitPrice = p.PricePerFoot.HasValue || p.PricePerPound.HasValue || p.PricePerPiece.HasValue;
        bool missingQty   = !p.UnitsRequested.HasValue && !p.UnitsQuoted.HasValue;
        if (!hasUnitPrice || !missingQty) return;

        // Match by MSPC first; fall back to the only RLI item when there is exactly one
        RliContextItem? rli = null;
        if (!string.IsNullOrEmpty(p.ProductSearchKey))
            rli = rliItems.FirstOrDefault(r =>
                string.Equals(r.Mspc, p.ProductSearchKey, StringComparison.OrdinalIgnoreCase));
        if (rli is null && rliItems.Count == 1)
            rli = rliItems[0];
        if (rli is null) return;

        if (rli.Quantity.HasValue)
            p.UnitsRequested = rli.Quantity;

        // For $/ft pricing, also fill in the length when the RLI specifies a size
        if (p.PricePerFoot.HasValue && !p.LengthPerUnit.HasValue && rli.SizeOfUnits is { Length: > 0 } sizeStr)
        {
            var lenFt = ParseRliSizeFt(sizeStr);
            if (lenFt.HasValue)
            {
                p.LengthPerUnit = lenFt.Value;
                p.LengthUnit    = "ft";
            }
        }
    }

    /// <summary>
    /// Parses an RLI SizeOfUnits string to feet. Mirrors Shredder's ParseSizeLength logic.
    /// Returns null for plate/sheet cross-dimension notation ("48 x 96") or unparseable values.
    /// </summary>
    private static double? ParseRliSizeFt(string size)
    {
        var s  = size.Trim();
        var nf = System.Globalization.NumberStyles.Any;
        var ic = System.Globalization.CultureInfo.InvariantCulture;

        // Plate: "48 x 96" — not a bar length
        if (System.Text.RegularExpressions.Regex.IsMatch(s, @"\d\s*[xX]\s*\d")) return null;

        // Explicit feet: "24'", "24 ft", "24 feet"
        var mFt = System.Text.RegularExpressions.Regex.Match(s,
            @"(\d+(?:\.\d+)?)\s*(?:ft|feet|')", System.Text.RegularExpressions.RegexOptions.IgnoreCase);
        if (mFt.Success && double.TryParse(mFt.Groups[1].Value, nf, ic, out var ft)) return ft;

        // Explicit inches: "288 in", "288\"", "288 inches"
        var mIn = System.Text.RegularExpressions.Regex.Match(s,
            @"(\d+(?:\.\d+)?)\s*(?:in|inch(?:es)?|"")", System.Text.RegularExpressions.RegexOptions.IgnoreCase);
        if (mIn.Success && double.TryParse(mIn.Groups[1].Value, nf, ic, out var inches)) return inches / 12.0;

        // Bare number: > 24 treated as inches, ≤ 24 as feet (same threshold as Shredder)
        if (double.TryParse(s, nf, ic, out var bare))
            return bare > 24 ? bare / 12.0 : bare;

        return null;
    }

}




