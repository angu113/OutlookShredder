using Microsoft.AspNetCore.Mvc;
using OutlookShredder.Proxy.Models;
using OutlookShredder.Proxy.Services;

namespace OutlookShredder.Proxy.Controllers;

[ApiController]
[Route("api")]
public class ExtractController : ControllerBase
{
    private readonly AiServiceFactory            _aiFactory;
    private readonly SharePointService          _sp;
    private readonly MailService                _mail;
    private readonly MailPollerService          _poller;
    private readonly OutlookComPollerService    _comPoller;
    private readonly SupplierCacheService       _suppliers;
    private readonly ProductCatalogService      _catalog;
    private readonly RfqNotificationService     _notifications;
    private readonly ShrConvInRouter            _shrRouter;
    private readonly IConfiguration             _config;
    private readonly ILogger<ExtractController> _log;

    public ExtractController(
        AiServiceFactory            aiFactory,
        SharePointService           sp,
        MailService                 mail,
        MailPollerService           poller,
        OutlookComPollerService     comPoller,
        SupplierCacheService        suppliers,
        ProductCatalogService       catalog,
        RfqNotificationService      notifications,
        ShrConvInRouter             shrRouter,
        IConfiguration              config,
        ILogger<ExtractController>  log)
    {
        _aiFactory     = aiFactory;
        _sp            = sp;
        _mail          = mail;
        _poller        = poller;
        _comPoller     = comPoller;
        _suppliers     = suppliers;
        _catalog       = catalog;
        _notifications = notifications;
        _shrRouter     = shrRouter;
        _config        = config;
        _log           = log;
    }

    // ── POST /api/extract ────────────────────────────────────────────────────
    /// <summary>
    /// Called by the Office.js add-in task pane.
    /// Extracts RFQ data from email body or attachment, then writes
    /// one SharePoint list item per product line found.
    /// </summary>
    [HttpPost("extract")]
    public async Task<ActionResult<ExtractResponse>> Extract([FromBody] ExtractRequest req)
    {
        if (string.IsNullOrWhiteSpace(req.Content) && string.IsNullOrWhiteSpace(req.Base64Data))
            return BadRequest(new ExtractResponse { Success = false, Error = "Content or Base64Data is required." });

        // ── SHR conversation short-circuit ───────────────────────────────────────
        // Shared with MailPollerService: supplier replies carrying [SHR:{rfqId}]
        // route straight to SupplierConversations instead of producing another SLI.
        var searchText = string.Join(" ", new[] { req.EmailSubject, req.EmailBody, req.Content }
            .Where(s => !string.IsNullOrEmpty(s)));
        var receivedAt = DateTimeOffset.TryParse(req.ReceivedAt, out var rt) ? rt : DateTimeOffset.UtcNow;
        var shrResult = await _shrRouter.TryRouteAsync(
            searchText:     searchText,
            fromAddr:       req.EmailFrom   ?? string.Empty,
            subject:        req.EmailSubject ?? string.Empty,
            body:           req.EmailBody   ?? req.Content ?? string.Empty,
            messageId:      req.ItemId,
            hasAttachments: req.HasAttachment,
            receivedAt:     receivedAt);

        if (shrResult.Routed)
        {
            var mailbox = _config["Mail:MailboxAddress"];
            if (!string.IsNullOrEmpty(mailbox) && !string.IsNullOrEmpty(req.ItemId))
            {
                try { await _mail.MarkProcessedAsync(mailbox, req.ItemId, "conv-in"); }
                catch (Exception ex)
                {
                    _log.LogWarning(ex, "[Extract] Could not stamp RFQ-Processed on conv-in message {Id}", req.ItemId);
                }
            }
            return Ok(new ExtractResponse
            {
                Success   = true,
                Extracted = new RfqExtraction { JobReference = shrResult.ShrRfqId, SupplierName = shrResult.ResolvedSupplier },
                Rows      = [],
            });
        }

        // Token present but supplier unresolvable — seed the rfqId so AI extraction
        // still files the row under the correct RFQ.
        if (shrResult.ShrRfqId is not null &&
            !req.JobRefs.Contains(shrResult.ShrRfqId, StringComparer.OrdinalIgnoreCase))
        {
            req.JobRefs.Insert(0, shrResult.ShrRfqId);
        }

        RfqExtraction? extraction;
        try
        {
            extraction = await _aiFactory.GetService().ExtractRfqAsync(req, HttpContext.RequestAborted);
        }
        catch (Exception ex)
        {
            _log.LogError(ex, "AI extraction failed");
            return StatusCode(502, new ExtractResponse { Success = false, Error = ex.Message });
        }

        // Always write at least one row so every email has a visible record.
        var source     = req.SourceType == "attachment" ? "attachment" : "body";
        var sourceFile = req.SourceType == "attachment" ? req.FileName : null;
        var rows       = new List<SpWriteResult>();

        var products = extraction?.Products ?? [];
        if (products.Count == 0)
        {
            products = [new ProductLine { SupplierProductComments = "No products could be extracted from this email." }];
            extraction ??= new RfqExtraction();
        }

        for (int i = 0; i < products.Count; i++)
        {
            var row = await _sp.WriteProductRowAsync(extraction!, products[i], req, source, sourceFile, i);
            rows.Add(row);
        }

        if (rows.Any(r => r.Success))
        {
            _notifications.NotifyRfqProcessed(new RfqProcessedNotification
            {
                EventType    = "SR",
                SupplierName = rows.FirstOrDefault(r => r.Success && !r.SupplierUnknown)?.SupplierName,
                RfqId        = req.JobRefs.FirstOrDefault()?.Trim('[', ']'),
                Products     = rows.Zip(products)
                                   .Where(x => x.First.Success)
                                   .Select(x => new RfqNotificationProduct
                                   {
                                       Name       = x.First.ProductName,
                                       TotalPrice = x.Second.TotalPrice,
                                   }).ToList(),
            });

            // Log the inbound email to SupplierConversations so the thread viewer
            // has a canonical record. Dedup is on MessageId inside WriteConversationMessageAsync.
            var firstGood = rows.FirstOrDefault(r => r.Success && !r.SupplierUnknown);
            if (firstGood is not null)
            {
                var convReceivedAt = DateTimeOffset.TryParse(req.ReceivedAt, out var crt) ? crt : DateTimeOffset.UtcNow;
                await _shrRouter.WriteConvInFromExtractionAsync(
                    rfqId:          firstGood.RfqId,
                    supplierName:   firstGood.SupplierName,
                    messageId:      req.ItemId,
                    subject:        req.EmailSubject,
                    body:           req.EmailBody ?? req.Content,
                    receivedAt:     convReceivedAt,
                    hasAttachments: req.HasAttachment,
                    fromAddr:       req.EmailFrom);
            }

            // Stamp "RFQ-Processed" on the mailbox message so the background poller
            // skips it on its next cycle.  Without this stamp the poller would re-run
            // the AI on the same email and write a second (duplicate) SR row.
            var mailbox = _config["Mail:MailboxAddress"];
            if (!string.IsNullOrEmpty(mailbox) && !string.IsNullOrEmpty(req.ItemId))
            {
                try
                {
                    var extra = rows.Any(r => r.SupplierUnknown) ? "Unknown" : null;
                    await _mail.MarkProcessedAsync(mailbox, req.ItemId, extra);
                    _log.LogInformation("[Extract] Stamped RFQ-Processed on message {Id}", req.ItemId);
                }
                catch (Exception ex)
                {
                    // Non-fatal — the row is already written; poller will upsert (not duplicate)
                    // now that FindExistingSupplierResponseAsync uses client-side filtering.
                    _log.LogWarning(ex, "[Extract] Could not stamp RFQ-Processed on message {Id}", req.ItemId);
                }
            }
        }

        return Ok(new ExtractResponse
        {
            Success   = rows.Any(r => r.Success),
            Extracted = extraction!,
            Rows      = rows
        });
    }

    // ── POST /api/setup-supplier-lists ───────────────────────────────────────
    /// <summary>
    /// Creates and provisions SupplierResponses and SupplierLineItems lists.
    /// Safe to re-run — skips columns that already exist.
    /// </summary>
    [HttpPost("setup-supplier-lists")]
    public async Task<IActionResult> SetupSupplierLists()
    {
        try
        {
            var results = await _sp.EnsureSupplierListsAsync();
            return Ok(new { success = true, lists = results });
        }
        catch (Exception ex)
        {
            return StatusCode(500, new { success = false, error = ex.Message });
        }
    }

    // ── POST /api/setup-columns ──────────────────────────────────────────────
    /// <summary>Legacy — provisions the old RFQ Line Items list. Kept for recovery.</summary>
    [HttpPost("setup-columns")]
    public async Task<IActionResult> SetupColumns()
    {
        try
        {
            var results = await _sp.EnsureColumnsAsync();
            return Ok(new { success = true, columns = results });
        }
        catch (Exception ex)
        {
            return StatusCode(500, new { success = false, error = ex.Message });
        }
    }

    // ── GET /api/sp-test ─────────────────────────────────────────────────────
    /// <summary>
    /// Diagnoses SharePoint connectivity step by step.
    /// <para><b>Development use only.</b> Remove or protect this endpoint before deploying to
    /// production — it returns token audience, tenant ID, and site details.</para>
    /// </summary>
    [HttpGet("sp-test")]
    public async Task<IActionResult> SpTest()
    {
        return Ok(await _sp.DiagnoseAsync());
    }

    // ── GET /api/items ────────────────────────────────────────────────────────
    /// <summary>
    /// Returns SupplierLineItems merged with their parent SupplierResponses fields.
    /// Shape is flat field dictionaries compatible with the Shredder dashboard DTO.
    /// </summary>
    [HttpGet("items")]
    public async Task<IActionResult> GetItems(
        [FromQuery] int       top              = 5000,
        [FromQuery] string?   nextLink         = null,
        [FromQuery] bool      raw              = false,
        [FromQuery] bool      includeCompleted = false,
        [FromQuery] DateTime? since            = null)
    {
        try
        {
            bool isFirstPage = nextLink is null;

            // Fetch items and (on first page only) total count in parallel.
            // The count is a raw SLI row count — an upper bound used by Shredder
            // to drive a determinate progress bar.  Zero latency added because
            // site/list ID lookups are cached after the first request.
            // `since` is passed on the first page only; subsequent pages follow the
            // @odata.nextLink which already carries the $filter.
            var itemsTask = _sp.ReadSupplierItemsAsync(top, nextLink, skipDedup: raw,
                                                       since: isFirstPage ? since : null);
            var countTask = isFirstPage
                ? _sp.GetSupplierLineItemCountAsync()
                : Task.FromResult(0);

            await Task.WhenAll(itemsTask, countTask);

            var (items, next) = itemsTask.Result;
            int? totalCount   = isFirstPage ? countTask.Result : null;

            if (!includeCompleted)
            {
                // Service-level cache shared across all pages of a single load —
                // avoids a separate ReadRfqReferences SP call per page.
                var completed = await _sp.GetCompletedRfqIdsAsync();
                if (completed.Count > 0)
                {
                    items = items.Where(row =>
                    {
                        if (!row.TryGetValue("RFQ_ID", out var idv)) return true;
                        var rid = idv?.ToString();
                        return string.IsNullOrEmpty(rid) || !completed.Contains(rid!);
                    }).ToList();
                }
            }

            return Ok(new { items, nextLink = next, totalCount });
        }
        catch (Exception ex)
        {
            _log.LogError(ex, "Failed to read supplier items");
            return StatusCode(500, new { success = false, error = ex.Message });
        }
    }

    private static bool IsTrueBool(object? v) => v switch
    {
        bool b                                       => b,
        string s when bool.TryParse(s, out var p)    => p,
        _                                            => false,
    };

    // ── GET /api/items/by-rfq/{rfqId} ────────────────────────────────────────
    /// <summary>
    /// Returns all SupplierLineItems for a specific RFQ ID — used for targeted UI refresh.
    /// Same flat DTO shape as GET /api/items but scoped to one job.
    /// </summary>
    [HttpGet("items/by-rfq/{rfqId}")]
    public async Task<IActionResult> GetItemsByRfq(string rfqId)
    {
        try
        {
            var items = await _sp.ReadSupplierItemsByRfqIdAsync(rfqId);
            return Ok(new { items, nextLink = (string?)null });
        }
        catch (Exception ex)
        {
            _log.LogError(ex, "Failed to read supplier items for rfqId={RfqId}", rfqId);
            return StatusCode(500, new { success = false, error = ex.Message });
        }
    }

    // ── GET /api/supplier-domains ────────────────────────────────────────────
    /// <summary>
    /// Returns a map of { emailDomain → canonicalSupplierName } sourced directly
    /// from the ContactEmail column of the Suppliers SharePoint list.
    /// </summary>
    [HttpGet("supplier-domains")]
    public IActionResult GetSupplierDomains() => Ok(_suppliers.DomainMap);

    // ── GET /api/rfq-references ───────────────────────────────────────────────
    /// <summary>
    /// Returns RFQ References records (RFQ_ID + Notes) for the dashboard header display.
    /// </summary>
    [HttpGet("rfq-references")]
    public async Task<IActionResult> GetRfqReferences(
        [FromQuery] bool includeCompleted = false)
    {
        try
        {
            var refs = await _sp.ReadRfqReferencesAsync();
            if (!includeCompleted)
                refs = refs.Where(r => !IsTrueBool(
                    r.TryGetValue("Complete", out var cv) ? cv : null)).ToList();
            return Ok(refs);
        }
        catch (Exception ex)
        {
            _log.LogError(ex, "Failed to read RFQ References");
            return StatusCode(500, new { success = false, error = ex.Message });
        }
    }

    // ── PATCH /api/rfq-references/notes ──────────────────────────────────────
    /// <summary>
    /// Updates the Notes field on a single RFQ Reference.
    /// Body: { "notes": "..." }
    /// </summary>
    [HttpPatch("rfq-references/notes")]
    public async Task<IActionResult> UpdateRfqNotes(
        [FromQuery] string rfqId,
        [FromBody]  RfqNotesRequest body)
    {
        if (string.IsNullOrWhiteSpace(rfqId))
            return BadRequest(new { error = "rfqId query param is required" });
        try
        {
            await _sp.UpdateRfqNotesAsync(rfqId, body.Notes ?? "");
            return Ok(new { updated = true });
        }
        catch (Exception ex)
        {
            _log.LogError(ex, "Failed to update Notes for RFQ '{Id}'", rfqId);
            return StatusCode(500, new { success = false, error = ex.Message });
        }
    }

    // ── PATCH /api/rfq-references/requester ──────────────────────────────────
    /// <summary>
    /// Sets the Requester field on a single RFQ Reference (claim / assign owner).
    /// Body: { "requester": "..." }
    /// </summary>
    [HttpPatch("rfq-references/requester")]
    public async Task<IActionResult> UpdateRfqRequester(
        [FromQuery] string rfqId,
        [FromBody]  RfqRequesterRequest body)
    {
        if (string.IsNullOrWhiteSpace(rfqId))
            return BadRequest(new { error = "rfqId query param is required" });
        try
        {
            await _sp.UpdateRfqRequesterAsync(rfqId, body.Requester ?? "");
            return Ok(new { updated = true });
        }
        catch (Exception ex)
        {
            _log.LogError(ex, "Failed to update Requester for RFQ '{Id}'", rfqId);
            return StatusCode(500, new { success = false, error = ex.Message });
        }
    }

    // ── PATCH /api/rfq-references/complete ───────────────────────────────────
    /// <summary>
    /// Sets the Complete boolean on a single RFQ Reference.
    /// </summary>
    [HttpPatch("rfq-references/complete")]
    public async Task<IActionResult> SetRfqComplete(
        [FromQuery] string rfqId,
        [FromQuery] bool   complete)
    {
        if (string.IsNullOrWhiteSpace(rfqId))
            return BadRequest(new { error = "rfqId query param is required" });
        try
        {
            await _sp.SetRfqCompleteAsync(rfqId, complete);
            return Ok(new { updated = true });
        }
        catch (Exception ex)
        {
            _log.LogError(ex, "Failed to set Complete for RFQ '{Id}'", rfqId);
            return StatusCode(500, new { success = false, error = ex.Message });
        }
    }

    // ── PATCH /api/rfq-references/flagged ────────────────────────────────────
    /// <summary>
    /// Sets the Flagged boolean on a single RFQ Reference.
    /// </summary>
    [HttpPatch("rfq-references/flagged")]
    public async Task<IActionResult> SetRfqFlagged(
        [FromQuery] string rfqId,
        [FromQuery] bool   flagged)
    {
        if (string.IsNullOrWhiteSpace(rfqId))
            return BadRequest(new { error = "rfqId query param is required" });
        try
        {
            await _sp.SetRfqFlaggedAsync(rfqId, flagged);
            return Ok(new { updated = true });
        }
        catch (Exception ex)
        {
            _log.LogError(ex, "Failed to set Flagged for RFQ '{Id}'", rfqId);
            return StatusCode(500, new { success = false, error = ex.Message });
        }
    }

    // ── GET /api/mail/body ────────────────────────────────────────────────────
    /// <summary>
    /// Returns the plain-text body and subject of the email identified by sender address
    /// and received timestamp.  Used by the Shredder price-comparison tab.
    /// </summary>
    [HttpGet("mail/body")]
    public async Task<IActionResult> GetMailBody(
        [FromQuery] string from,
        [FromQuery] string receivedAt)
    {
        if (!DateTimeOffset.TryParse(receivedAt, out var dt))
            return BadRequest(new { error = "receivedAt must be ISO 8601" });

        try
        {
            var result = await _mail.GetBodyAsync(from, dt);
            if (result is null)
                return NotFound(new { error = "Email not found in mailbox" });

            return Ok(new { subject = result.Value.Subject, body = result.Value.BodyText });
        }
        catch (Exception ex)
        {
            _log.LogError(ex, "GetMailBody failed for from={From} receivedAt={At}", from, receivedAt);
            return StatusCode(500, new { error = ex.Message });
        }
    }

    // ── GET /api/mail/attachment ──────────────────────────────────────────────
    /// <summary>
    /// Returns the raw attachment bytes so the Shredder client can save them to a
    /// temp file and open them with the default system viewer.
    /// </summary>
    [HttpGet("mail/attachment")]
    public async Task<IActionResult> GetMailAttachment(
        [FromQuery] string from,
        [FromQuery] string receivedAt,
        [FromQuery] string filename)
    {
        if (!DateTimeOffset.TryParse(receivedAt, out var dt))
            return BadRequest(new { error = "receivedAt must be ISO 8601" });

        try
        {
            var result = await _mail.GetAttachmentAsync(from, dt, filename);
            if (result is null)
                return NotFound(new { error = $"Attachment '{filename}' not found in mailbox. " +
                    "The email may have been processed from a different mailbox that this proxy does not have access to." });

            return File(result.Value.Bytes, result.Value.ContentType, result.Value.FileName);
        }
        catch (Exception ex)
        {
            _log.LogError(ex, "GetMailAttachment failed for from={From} file={File}", from, filename);
            return StatusCode(500, new { error = ex.Message });
        }
    }

    // ── GET /api/sp-attachment ────────────────────────────────────────────────
    /// <summary>
    /// Returns a PDF/attachment stored directly on a SupplierResponses SP list item.
    /// Preferred over /api/mail/attachment — does not require mailbox search.
    /// </summary>
    [HttpGet("sp-attachment")]
    public async Task<IActionResult> GetSpAttachment(
        [FromQuery] string srId,
        [FromQuery] string filename)
    {
        if (string.IsNullOrWhiteSpace(srId) || string.IsNullOrWhiteSpace(filename))
            return BadRequest(new { error = "srId and filename are required" });

        try
        {
            var result = await _sp.GetSpItemAttachmentAsync(srId, filename);
            if (result is null)
                return NotFound(new { error = $"Attachment '{filename}' not found on SR item {srId}." });

            return File(result.Value.Bytes, result.Value.ContentType, result.Value.FileName);
        }
        catch (Exception ex)
        {
            _log.LogError(ex, "GetSpAttachment failed for srId={SrId} file={File}", srId, filename);
            return StatusCode(500, new { error = ex.Message });
        }
    }

    // ── GET /api/suppliers ────────────────────────────────────────────────────
    /// <summary>
    /// Returns the cached canonical supplier name list.
    /// Refreshes automatically on startup and every hour.
    /// </summary>
    [HttpGet("suppliers")]
    public IActionResult GetSuppliers() =>
        Ok(new { suppliers = _suppliers.CachedNames });

    // ── GET /api/events ───────────────────────────────────────────────────────
    /// <summary>
    /// Server-Sent Events stream.  The dashboard connects here and reloads its
    /// data whenever an <c>rfq-processed</c> event arrives.
    /// The connection is long-lived; the browser reconnects automatically if it drops.
    /// </summary>
    [HttpGet("events")]
    public async Task StreamEvents(CancellationToken ct)
    {
        Response.Headers.Append("Content-Type",      "text/event-stream");
        Response.Headers.Append("Cache-Control",     "no-cache");
        Response.Headers.Append("X-Accel-Buffering", "no");

        // Flush headers + opening comment immediately so the client knows
        // the connection is live before any real event arrives.
        await Response.WriteAsync(": connected\n\n", ct);
        await Response.Body.FlushAsync(ct);

        var (id, reader) = _notifications.Subscribe();
        try
        {
            while (!ct.IsCancellationRequested)
            {
                // Wait up to 30 s for an event; on timeout send a keepalive comment
                // so proxies and clients don't close the idle connection.
                using var waitCts = CancellationTokenSource.CreateLinkedTokenSource(ct);
                waitCts.CancelAfter(TimeSpan.FromSeconds(30));

                try
                {
                    await reader.WaitToReadAsync(waitCts.Token);

                    // Drain all events queued since the last flush.
                    // Message format: "{eventName}\n{dataJson}" — split on first newline.
                    while (reader.TryRead(out var msg))
                    {
                        var nl   = msg.IndexOf('\n');
                        var evt  = nl >= 0 ? msg[..nl]    : msg;
                        var data = nl >= 0 ? msg[(nl+1)..] : "{}";
                        await Response.WriteAsync($"event: {evt}\ndata: {data}\n\n", ct);
                    }
                }
                catch (OperationCanceledException) when (!ct.IsCancellationRequested)
                {
                    // 30 s keepalive — SSE comments are ignored by clients.
                    await Response.WriteAsync(": keepalive\n\n", ct);
                }

                await Response.Body.FlushAsync(ct);
            }
        }
        catch (OperationCanceledException) { /* client disconnected */ }
        finally
        {
            _notifications.Unsubscribe(id);
        }
    }

    // ── POST /api/mail/reprocess ─────────────────────────────────────────────
    /// <summary>
    /// Triggers an immediate full scan of all unprocessed inbox messages without
    /// requiring a proxy restart.  Pair with POST /api/mail/reset to reprocess everything.
    /// </summary>
    [HttpPost("mail/reprocess")]
    public IActionResult TriggerReprocess()
    {
        _poller.TriggerReprocessAll();
        return Ok(new { triggered = true });
    }

    // ── POST /api/mail/reset ─────────────────────────────────────────────────
    /// <summary>
    /// Removes "RFQ-Processed" from all inbox messages so the next poll cycle
    /// (or proxy restart) will reprocess every email from scratch.
    /// </summary>
    [HttpPost("mail/reset")]
    public async Task<IActionResult> ResetMailCategories()
    {
        var mailbox = HttpContext.RequestServices
            .GetRequiredService<IConfiguration>()["Mail:MailboxAddress"];
        if (string.IsNullOrEmpty(mailbox))
            return BadRequest(new { error = "Mail:MailboxAddress not configured" });

        try
        {
            var count = await _mail.UnmarkAllAsync(mailbox);
            _log.LogInformation("[Mail] Reset: removed RFQ-Processed from {Count} message(s)", count);
            return Ok(new { unmarked = count });
        }
        catch (Exception ex)
        {
            _log.LogError(ex, "[Mail] Reset failed");
            return StatusCode(500, new { error = ex.Message });
        }
    }

    // ── GET /api/mail/processed-emails ───────────────────────────────────────
    /// <summary>
    /// Returns the most recent inbox messages already tagged "RFQ-Processed".
    /// Used by the Shredder Reprocess panel to let the user pick emails to re-run.
    /// </summary>
    [HttpGet("mail/processed-emails")]
    public async Task<IActionResult> GetProcessedEmails([FromQuery] int top = 200)
    {
        var mailbox = HttpContext.RequestServices
            .GetRequiredService<IConfiguration>()["Mail:MailboxAddress"];
        if (string.IsNullOrEmpty(mailbox))
            return BadRequest(new { error = "Mail:MailboxAddress not configured" });

        try
        {
            var emails = await _mail.GetProcessedMessagesAsync(mailbox, top);
            return Ok(emails);
        }
        catch (Exception ex)
        {
            _log.LogError(ex, "GetProcessedEmails failed");
            return StatusCode(500, new { error = ex.Message });
        }
    }

    // ── POST /api/mail/reprocess-selected ────────────────────────────────────
    /// <summary>
    /// Fetches each listed message from Graph and re-runs the full extraction pipeline
    /// (AI extraction + SharePoint upsert + re-stamp "RFQ-Processed").
    /// Awaits completion before responding so the client knows when all rows are written.
    /// </summary>
    [HttpPost("mail/reprocess-selected")]
    public async Task<IActionResult> ReprocessSelected(
        [FromBody] ReprocessRequest req,
        CancellationToken ct)
    {
        if (req.MessageIds.Count == 0)
            return BadRequest(new { error = "messageIds is required" });

        try
        {
            await _poller.ReprocessMessagesAsync(req.MessageIds, ct);
            return Ok(new { reprocessed = req.MessageIds.Count });
        }
        catch (Exception ex)
        {
            _log.LogError(ex, "ReprocessSelected failed");
            return StatusCode(500, new { error = ex.Message });
        }
    }

    // ── POST /api/mail/reprocess-null-mspc ──────────────────────────────────
    /// <summary>
    /// Finds all SupplierLineItem rows with a valid RFQ ID but null ProductSearchKey (MSPC),
    /// then re-runs the full extraction pipeline on their source emails so the late-RLI-inject
    /// path can assign correct MSPCs.  Pass ?dryRun=true to count candidates without reprocessing.
    /// </summary>
    [HttpPost("mail/reprocess-null-mspc")]
    public async Task<IActionResult> ReprocessNullMspc(
        [FromQuery] bool dryRun = false,
        CancellationToken ct = default)
    {
        try
        {
            var messageIds = await _sp.FindNullMspcMessageIdsAsync(ct);
            if (dryRun || messageIds.Count == 0)
                return Ok(new { found = messageIds.Count, reprocessed = 0, dryRun = true });

            await _poller.ReprocessMessagesAsync(messageIds, ct);
            return Ok(new { found = messageIds.Count, reprocessed = messageIds.Count, dryRun = false });
        }
        catch (Exception ex)
        {
            _log.LogError(ex, "ReprocessNullMspc failed");
            return StatusCode(500, new { error = ex.Message });
        }
    }

    // ── GET /api/rfq-import/scan ─────────────────────────────────────────────
    /// <summary>
    /// Scans a named mail folder in the given mailbox and returns raw email data
    /// (subject, body, recipients) for all "RFQ [JobNo]" messages.
    /// Shredder parses line items locally.
    /// </summary>
    [HttpGet("rfq-import/scan")]
    public async Task<IActionResult> ScanRfqFolder(
        [FromQuery] string mailbox,
        [FromQuery] string folder = "RFQOut")
    {
        if (string.IsNullOrWhiteSpace(mailbox))
            return BadRequest(new { error = "mailbox query param is required" });
        try
        {
            var emails = await _mail.ScanRfqFolderAsync(mailbox, folder);
            return Ok(emails);
        }
        catch (Exception ex)
        {
            _log.LogError(ex, "ScanRfqFolder failed for {Mailbox}/{Folder}", mailbox, folder);
            return StatusCode(500, new { error = ex.Message });
        }
    }

    // ── GET /api/rfq-import/existing-ids ─────────────────────────────────────
    /// <summary>Returns the set of RFQ_ID values already in the RFQ References list.</summary>
    [HttpGet("rfq-import/existing-ids")]
    public async Task<IActionResult> GetExistingRfqIds()
    {
        try
        {
            var ids = await _sp.GetExistingRfqIdsAsync();
            return Ok(ids.ToArray());
        }
        catch (Exception ex)
        {
            _log.LogError(ex, "GetExistingRfqIds failed");
            return StatusCode(500, new { error = ex.Message });
        }
    }

    // ── GET /api/rfq-import/supplier-responses-raw ───────────────────────────
    /// <summary>Returns raw SR rows: RFQ_ID, EmailFrom, ReceivedAt. Used for backfill analysis.</summary>
    [HttpGet("rfq-import/supplier-responses-raw")]
    public async Task<IActionResult> GetSupplierResponsesRaw()
    {
        try
        {
            var rows = await _sp.ReadSupplierResponsesRawAsync();
            return Ok(rows);
        }
        catch (Exception ex)
        {
            _log.LogError(ex, "GetSupplierResponsesRaw failed");
            return StatusCode(500, new { error = ex.Message });
        }
    }

    // ── POST /api/rfq-import/reference ───────────────────────────────────────
    /// <summary>Creates one RFQ Reference row. No-op check is the caller's responsibility.</summary>
    [HttpPost("rfq-import/reference")]
    public async Task<IActionResult> CreateRfqReference([FromBody] RfqReferenceRequest req)
    {
        try
        {
            await _sp.CreateRfqReferenceAsync(req);
            _notifications.NotifyRfqProcessed(new RfqProcessedNotification
            {
                EventType = "RFQ",
                RfqId     = req.RfqId,
            });
            return Ok(new { created = true });
        }
        catch (Exception ex)
        {
            _log.LogError(ex, "CreateRfqReference failed for '{Id}'", req.RfqId);
            return StatusCode(500, new { error = ex.Message });
        }
    }

    // ── GET /api/rfq-import/line-items-lookup ────────────────────────────────
    /// <summary>
    /// Returns all RFQ Line Items (RFQ_ID, MSPC, Product) for the Shredder dashboard header display.
    /// </summary>
    [HttpGet("rfq-import/line-items-lookup")]
    public async Task<IActionResult> GetRfqLineItemsLookup()
    {
        try
        {
            var items  = await _sp.ReadAllRfqLineItemsAsync();
            var result = items.Select(x => new { rfqId = x.RfqId, mspc = x.Mspc, product = x.Product, units = x.Units, sizeOfUnits = x.SizeOfUnits, isPurchased = x.IsPurchased, poNumber = x.PoNumber }).ToList();
            return Ok(result);
        }
        catch (Exception ex)
        {
            _log.LogError(ex, "GetRfqLineItemsLookup failed");
            return StatusCode(500, new { error = ex.Message });
        }
    }

    // ── POST /api/rfq-import/line-items ──────────────────────────────────────
    /// <summary>
    /// Batch-creates RFQ Line Item rows.
    /// Skips any RFQ_ID that already has at least one line item in the list.
    /// </summary>
    [HttpPost("rfq-import/line-items")]
    public async Task<IActionResult> CreateRfqLineItems([FromBody] List<RfqLineItemRequest> items)
    {
        try
        {
            var existingRfqIds = await _sp.GetRfqIdsWithLineItemsAsync();
            var toCreate = items.Where(i => !existingRfqIds.Contains(i.RfqId)).ToList();
            if (toCreate.Count > 0)
            {
                await _sp.CreateRfqLineItemsAsync(toCreate);

                // Fire one "RFQ" event per newly-created RFQ so Service Bus consumers
                // (other Shredder instances, downstream systems) know the full RFQ is ready.
                foreach (var grp in toCreate.GroupBy(i => i.RfqId, StringComparer.OrdinalIgnoreCase))
                {
                    _notifications.NotifyRfqProcessed(new RfqProcessedNotification
                    {
                        EventType = "RFQ",
                        RfqId     = grp.Key,
                        Products  = grp
                            .Where(li => li.Product is not null)
                            .Select(li => new RfqNotificationProduct
                            {
                                Name = string.IsNullOrWhiteSpace(li.Mspc)
                                    ? li.Product
                                    : $"{li.Mspc} | {li.Product}",
                            })
                            .ToList(),
                    });
                    _log.LogInformation("[RfqNew] Fired RFQ event for {RfqId} ({Count} line items)",
                        grp.Key, grp.Count());
                }
            }

            return Ok(new { created = toCreate.Count, skipped = items.Count - toCreate.Count });
        }
        catch (Exception ex)
        {
            _log.LogError(ex, "CreateRfqLineItems failed");
            return StatusCode(500, new { error = ex.Message });
        }
    }

    // ── POST /api/rfq-import/dedupe-supplier-responses ──────────────────────
    /// <summary>
    /// Merges duplicate SupplierResponse rows that share the same (RFQ_ID, SupplierName).
    /// Keeps the row with the best data (attachment > priced SLI > newest), re-parents or
    /// deletes orphaned SupplierLineItems, then deletes the duplicate SR rows.
    /// Safe to call repeatedly — idempotent once all duplicates are resolved.
    /// </summary>
    [HttpPost("rfq-import/dedupe-supplier-responses")]
    public async Task<IActionResult> DedupeSupplierResponses([FromQuery] bool dryRun = false, [FromQuery] string? rfqId = null)
    {
        try
        {
            var result = await _sp.DedupeSupplierResponsesAsync(dryRun, rfqId);
            _log.LogInformation("[Dedupe-SR] Endpoint complete — dryRun={DryRun}, {G} groups, {Sr} SR deleted, {SliR} SLI re-parented, {SliD} SLI deleted",
                dryRun, result.DuplicateGroups, result.SrDeleted, result.SliReparented, result.SliDeleted);
            return Ok(new
            {
                dryRun             = result.DryRun,
                duplicateGroups    = result.DuplicateGroups,
                srDeleted          = result.SrDeleted,
                sliReparented      = result.SliReparented,
                sliDeleted         = result.SliDeleted,
                groups             = result.Groups,
                sliDuplicateGroups = result.SliDuplicateGroups,
                sliWithinSrDeleted = result.SliWithinSrDeleted,
                sliGroups          = result.SliGroups,
            });
        }
        catch (Exception ex)
        {
            _log.LogError(ex, "DedupeSupplierResponses failed");
            return StatusCode(500, new { error = ex.Message });
        }
    }

    // ── POST /api/rfq-import/dedupe-references ───────────────────────────────
    /// <summary>
    /// Removes duplicate RFQ Reference rows (same RFQ_ID), keeping the entry with Notes
    /// (or oldest if all blank).  Safe to call at any time.
    /// </summary>
    [HttpPost("rfq-import/dedupe-references")]
    public async Task<IActionResult> DedupeRfqReferences()
    {
        try
        {
            var deleted = await _sp.DedupeRfqReferencesAsync();
            _log.LogInformation("[Dedupe] Removed {Count} duplicate RFQ Reference(s)", deleted);
            return Ok(new { deleted });
        }
        catch (Exception ex)
        {
            _log.LogError(ex, "DedupeRfqReferences failed");
            return StatusCode(500, new { error = ex.Message });
        }
    }

    // ── DELETE /api/sr/{srId} ───────────────────────────────────────────────
    /// <summary>
    /// Deletes a SupplierResponse row and all of its child SupplierLineItem rows.
    /// Use this to fully remove one supplier's data for an RFQ so the email can be reprocessed cleanly.
    /// </summary>
    [HttpDelete("sr/{srId}")]
    public async Task<IActionResult> DeleteSr(string srId)
    {
        try
        {
            var (sliDeleted, srDeleted) = await _sp.DeleteSrAsync(srId);
            return Ok(new { srId, srDeleted, sliDeleted });
        }
        catch (Exception ex)
        {
            _log.LogError(ex, "DeleteSr failed for SR {Id}", srId);
            return StatusCode(500, new { error = ex.Message });
        }
    }

    // ── PATCH /api/sr/{srId}/rfq-id ─────────────────────────────────────────────

    /// <summary>Re-parents a SupplierResponse (and its child SLI rows) to a different RFQ ID.</summary>
    [HttpPatch("sr/{srId}/rfq-id")]
    public async Task<IActionResult> ReparentSr(string srId, [FromBody] ReparentSrRequest request)
    {
        if (string.IsNullOrWhiteSpace(request?.RfqId) ||
            !System.Text.RegularExpressions.Regex.IsMatch(request.RfqId.Trim(), @"^([Hh][Qq][A-Za-z0-9]{6}|[A-Za-z0-9]{6})$"))
            return BadRequest(new { error = "rfqId must be 'HQ'+6 alphanumeric (new) or 6 alphanumeric (legacy)" });

        try
        {
            var rfqId = request.RfqId.Trim().ToUpperInvariant();
            await _sp.ReparentSupplierResponseAsync(srId, rfqId);
            return Ok(new { srId, rfqId });
        }
        catch (Exception ex)
        {
            _log.LogError(ex, "ReparentSr failed for SR {Id}", srId);
            return StatusCode(500, new { error = ex.Message });
        }
    }

    // ── PATCH /api/sli/{sliItemId}/rfq-id ───────────────────────────────────────

    /// <summary>
    /// Re-parents a single SupplierLineItem to a different RFQ ID.
    /// If the source SR ends up with no remaining SLI children it is deleted automatically.
    /// </summary>
    [HttpPatch("sli/{sliItemId}/rfq-id")]
    public async Task<IActionResult> ReparentSli(string sliItemId, [FromBody] ReparentSrRequest request)
    {
        if (string.IsNullOrWhiteSpace(request?.RfqId) ||
            !System.Text.RegularExpressions.Regex.IsMatch(request.RfqId.Trim(), @"^([Hh][Qq][A-Za-z0-9]{6}|[A-Za-z0-9]{6})$"))
            return BadRequest(new { error = "rfqId must be 'HQ'+6 alphanumeric (new) or 6 alphanumeric (legacy)" });

        var rfqId = request.RfqId.Trim();
        try
        {
            var srDeleted = await _sp.ReparentSliAsync(sliItemId, rfqId);
            return Ok(new { sliItemId, rfqId, srDeleted });
        }
        catch (Exception ex)
        {
            _log.LogError(ex, "ReparentSli failed for SLI {Id}", sliItemId);
            return StatusCode(500, new { error = ex.Message });
        }
    }

    // ── DELETE /api/sli/{itemId} ─────────────────────────────────────────────
    /// <summary>
    /// Deletes a single SupplierLineItem by its SharePoint item ID.
    /// </summary>
    [HttpDelete("sli/{itemId}")]
    public async Task<IActionResult> DeleteSli(string itemId)
    {
        try
        {
            await _sp.DeleteSliAsync(itemId);
            return Ok(new { deleted = itemId });
        }
        catch (Exception ex)
        {
            _log.LogError(ex, "DeleteSli failed for item {Id}", itemId);
            return StatusCode(500, new { error = ex.Message });
        }
    }

    // ── DELETE /api/rfq-import/clean ─────────────────────────────────────────
    /// <summary>
    /// Deletes all rows from SupplierResponses and SupplierLineItems.
    /// Does not touch RFQ References (notes/dates).
    /// </summary>
    [HttpDelete("rfq-import/clean")]
    public async Task<IActionResult> CleanSupplierData()
    {
        try
        {
            var (srDeleted, sliDeleted) = await _sp.CleanSupplierDataAsync();
            _log.LogWarning("[Clean] Deleted {Sr} SupplierResponses and {Sli} SupplierLineItems", srDeleted, sliDeleted);
            return Ok(new { srDeleted, sliDeleted });
        }
        catch (Exception ex)
        {
            _log.LogError(ex, "CleanSupplierData failed");
            return StatusCode(500, new { error = ex.Message });
        }
    }

    // ── DELETE /api/rfq-import/purge-old?days=N ──────────────────────────────
    /// <summary>
    /// Deletes SupplierLineItems and SupplierResponses older than <paramref name="days"/> days.
    /// Default: 7 days.
    /// </summary>
    [HttpDelete("rfq-import/purge-old")]
    public async Task<IActionResult> PurgeOldSupplierData([FromQuery] int days = 7)
    {
        try
        {
            var (srDeleted, sliDeleted) = await _sp.PurgeOldSupplierDataAsync(days);
            _log.LogWarning("[Purge] Deleted {Sr} SupplierResponses and {Sli} SupplierLineItems older than {Days} days",
                srDeleted, sliDeleted, days);
            return Ok(new { srDeleted, sliDeleted, days });
        }
        catch (Exception ex)
        {
            _log.LogError(ex, "PurgeOldSupplierData failed");
            return StatusCode(500, new { error = ex.Message });
        }
    }

    // ── DELETE /api/rfq-import/clean-all ─────────────────────────────────────
    /// <summary>
    /// Deletes all rows from all four RFQ lists:
    /// RFQ References, RFQ Line Items, SupplierResponses, SupplierLineItems.
    /// </summary>
    [HttpDelete("rfq-import/clean-all")]
    public async Task<IActionResult> CleanAllData()
    {
        try
        {
            var (refsDeleted, rliDeleted, srDeleted, sliDeleted) = await _sp.CleanAllDataAsync();
            _log.LogWarning("[Clean] Deleted {Refs} RFQ References, {Rli} RFQ Line Items, {Sr} SupplierResponses, {Sli} SupplierLineItems",
                refsDeleted, rliDeleted, srDeleted, sliDeleted);
            return Ok(new { refsDeleted, rliDeleted, srDeleted, sliDeleted });
        }
        catch (Exception ex)
        {
            _log.LogError(ex, "CleanAllData failed");
            return StatusCode(500, new { error = ex.Message });
        }
    }

    // ── POST /api/rfq-import/com-scan-and-import ──────────────────────────────
    /// <summary>
    /// Scans hackensack@metalsupermarkets.com (or any configured mailbox) Sent Items
    /// via Outlook COM automation, parses RFQ line items from the email body, and
    /// writes RFQ References + Line Items to SharePoint.
    /// Requires Outlook to be running on this machine with the account open.
    /// </summary>
    [HttpPost("rfq-import/com-scan-and-import")]
    public async Task<IActionResult> ComScanAndImport(
        [FromQuery] string mailbox,
        [FromQuery] int    days = 90)
    {
        if (string.IsNullOrWhiteSpace(mailbox))
            return BadRequest(new { error = "mailbox query param is required" });

        List<RfqScanEmailDto> emails;
        try
        {
#pragma warning disable CA1416 // proxy runs Windows-only
            emails = await _comPoller.ScanRfqSentItemsAsync(mailbox, days);
#pragma warning restore CA1416
        }
        catch (Exception ex)
        {
            _log.LogError(ex, "COM scan failed for {Mailbox}", mailbox);
            return StatusCode(500, new { error = ex.Message });
        }

        if (emails.Count == 0)
            return Ok(new { scanned = 0, imported = 0, skipped = 0, message = "No RFQ emails found. Ensure Outlook is running with the account open." });

        var existingIds = await _sp.GetExistingRfqIdsAsync();
        var existingLiIds = await _sp.GetRfqIdsWithLineItemsAsync();

        int imported = 0, skipped = 0;

        foreach (var email in emails)
        {
            try
            {
                if (!existingIds.Contains(email.RfqId))
                {
                    var requester = ParseRequesterName(email.BodyText) ?? email.Requester;
                    await _sp.CreateRfqReferenceAsync(new RfqReferenceRequest
                    {
                        RfqId           = email.RfqId,
                        Requester       = requester,
                        DateSent        = email.SentAt,
                        EmailRecipients = email.EmailRecipients,
                    });
                    existingIds.Add(email.RfqId);
                }

                if (!existingLiIds.Contains(email.RfqId))
                {
                    var lineItems = ParseRfqLineItems(email.RfqId, email.BodyText);
                    if (lineItems.Count > 0)
                    {
                        await _sp.CreateRfqLineItemsAsync(lineItems);
                        existingLiIds.Add(email.RfqId);
                    }
                }

                imported++;
            }
            catch (Exception ex)
            {
                _log.LogError(ex, "Failed to import RFQ {Id}", email.RfqId);
                skipped++;
            }
        }

        _log.LogInformation("[COM Import] {Mailbox}: scanned={Scanned} imported={Imported} skipped={Skipped}",
            mailbox, emails.Count, imported, skipped);

        return Ok(new { scanned = emails.Count, imported, skipped });
    }

    // ── POST /api/rfq-import/graph-scan-and-import ────────────────────────────
    /// <summary>
    /// Same as com-scan-and-import but uses Microsoft Graph (for mailboxes in the
    /// mithrilmetals.com tenant, e.g. store@mithrilmetals.com).
    /// Scans the Sent Items folder for outbound "RFQ [XXXXXX]" emails.
    /// </summary>
    [HttpPost("rfq-import/graph-scan-and-import")]
    public async Task<IActionResult> GraphScanAndImport(
        [FromQuery] string mailbox,
        [FromQuery] int    days = 7)
    {
        if (string.IsNullOrWhiteSpace(mailbox))
            return BadRequest(new { error = "mailbox query param is required" });

        List<RfqScanEmailDto> emails;
        try
        {
            emails = await _mail.ScanRfqFolderAsync(mailbox, "Sent Items", days);
        }
        catch (Exception ex)
        {
            _log.LogError(ex, "Graph RFQ scan failed for {Mailbox}", mailbox);
            return StatusCode(500, new { error = ex.Message });
        }

        if (emails.Count == 0)
            return Ok(new { scanned = 0, imported = 0, skipped = 0 });

        var existingIds   = await _sp.GetExistingRfqIdsAsync();
        var existingLiIds = await _sp.GetRfqIdsWithLineItemsAsync();
        int imported = 0, skipped = 0;

        foreach (var email in emails)
        {
            try
            {
                if (!existingIds.Contains(email.RfqId))
                {
                    var requester = ParseRequesterName(email.BodyText) ?? email.Requester;
                    await _sp.CreateRfqReferenceAsync(new RfqReferenceRequest
                    {
                        RfqId           = email.RfqId,
                        Requester       = requester,
                        DateSent        = email.SentAt,
                        EmailRecipients = email.EmailRecipients,
                    });
                    existingIds.Add(email.RfqId);
                }

                if (!existingLiIds.Contains(email.RfqId))
                {
                    var lineItems = ParseRfqLineItems(email.RfqId, email.BodyText);
                    if (lineItems.Count > 0)
                    {
                        await _sp.CreateRfqLineItemsAsync(lineItems);
                        existingLiIds.Add(email.RfqId);
                    }
                }

                imported++;
            }
            catch (Exception ex)
            {
                _log.LogError(ex, "Failed to import RFQ {Id}", email.RfqId);
                skipped++;
            }
        }

        _log.LogInformation("[Graph Import] {Mailbox}: scanned={Scanned} imported={Imported} skipped={Skipped}",
            mailbox, emails.Count, imported, skipped);

        return Ok(new { scanned = emails.Count, imported, skipped });
    }

    // ── RFQ email parsing helpers ─────────────────────────────────────────────

    private static readonly System.Text.RegularExpressions.Regex _rliRegex = new(
        @"^([A-Za-z0-9/\-]+)\s*\|\s*(.+?)\s*\|\s*Qty:\s*(\d+(?:\.\d+)?)\s*\|\s*Size:\s*(.+?)\s*$",
        System.Text.RegularExpressions.RegexOptions.IgnoreCase |
        System.Text.RegularExpressions.RegexOptions.Multiline  |
        System.Text.RegularExpressions.RegexOptions.Compiled);

    private static readonly System.Text.RegularExpressions.Regex _rliLegacyRegex = new(
        @"^(.+?)\s*\|\s*Qty:\s*(\d+(?:\.\d+)?)\s*\|\s*Size:\s*(.+?)\s*$",
        System.Text.RegularExpressions.RegexOptions.IgnoreCase |
        System.Text.RegularExpressions.RegexOptions.Multiline  |
        System.Text.RegularExpressions.RegexOptions.Compiled);

    // Matches valediction line: "Thank you," / "Regards," / "Best regards," etc.
    private static readonly System.Text.RegularExpressions.Regex _valedictionRegex = new(
        @"^(thank\s+you|thanks|best\s+regards?|kind\s+regards?|regards?|sincerely|cheers)[,.]?\s*$",
        System.Text.RegularExpressions.RegexOptions.IgnoreCase |
        System.Text.RegularExpressions.RegexOptions.Multiline  |
        System.Text.RegularExpressions.RegexOptions.Compiled);

    // Matches inline valediction: "Thank you, Angus." or "Thank you, Jay M."
    private static readonly System.Text.RegularExpressions.Regex _valedictionInlineRegex = new(
        @"^(?:thank\s+you|thanks|best\s+regards?|kind\s+regards?|regards?|sincerely|cheers)[,\s]+([A-Za-z][A-Za-z.\s]{1,30}?)\.?\s*$",
        System.Text.RegularExpressions.RegexOptions.IgnoreCase |
        System.Text.RegularExpressions.RegexOptions.Multiline  |
        System.Text.RegularExpressions.RegexOptions.Compiled);

    // Lines that look like company/address info rather than a person name
    private static readonly System.Text.RegularExpressions.Regex _companyLineRegex = new(
        @"metal\s+super|mithril|franchis|phone|fax|www\.|http|@|\d{3,}|inc\b|corp\b|llc\b|ltd\b",
        System.Text.RegularExpressions.RegexOptions.IgnoreCase |
        System.Text.RegularExpressions.RegexOptions.Compiled);

    /// <summary>
    /// Extracts the sender's name from an RFQ email body by finding the valediction
    /// ("Thank you," / "Regards," etc.) and reading the name on the following line,
    /// or inline on the same line.  Returns null if no name can be found.
    /// </summary>
    private static string? ParseRequesterName(string bodyText)
    {
        // Strategy 1: valediction on its own line, name on next non-empty line
        var valedMatch = _valedictionRegex.Match(bodyText);
        if (valedMatch.Success)
        {
            var afterValediction = bodyText[(valedMatch.Index + valedMatch.Length)..];
            foreach (var line in afterValediction.Split('\n'))
            {
                var candidate = line.Trim('\r', '\n', ' ', '\t');
                if (string.IsNullOrWhiteSpace(candidate)) continue;
                if (candidate.Length > 40) break;          // too long to be a name
                if (_companyLineRegex.IsMatch(candidate)) break;  // looks like company line
                return candidate;
            }
        }

        // Strategy 2: valediction with name inline ("Thank you, Angus.")
        var inlineMatch = _valedictionInlineRegex.Match(bodyText);
        if (inlineMatch.Success)
        {
            var candidate = inlineMatch.Groups[1].Value.Trim();
            if (!_companyLineRegex.IsMatch(candidate))
                return candidate;
        }

        return null;
    }

    /// <summary>
    /// Parses RFQ line items from the email body.
    /// For items without an MSPC code (legacy format), looks up the catalog to find it.
    /// </summary>
    private List<RfqLineItemRequest> ParseRfqLineItems(string rfqId, string bodyText)
    {
        var results = new List<RfqLineItemRequest>();

        // Full format: MSPC | Product | Qty: N | Size: S
        foreach (System.Text.RegularExpressions.Match m in _rliRegex.Matches(bodyText))
        {
            var mspc    = m.Groups[1].Value.Trim();
            var product = m.Groups[2].Value.Trim();

            // Attempt catalog resolution to get canonical name + search key
            var resolved = _catalog.ResolveProduct(product);
            results.Add(new RfqLineItemRequest
            {
                RfqId            = rfqId,
                Mspc             = resolved?.SearchKey ?? mspc,
                Product          = resolved?.Name      ?? product,
                Units            = double.TryParse(m.Groups[3].Value, out var q) ? q : null,
                SizeOfUnits      = m.Groups[4].Value.Trim(),
                ProcessingSource = "import",
            });
        }

        // Legacy format: Product | Qty: N | Size: S  (no MSPC prefix)
        if (results.Count == 0)
        {
            foreach (System.Text.RegularExpressions.Match m in _rliLegacyRegex.Matches(bodyText))
            {
                var product  = m.Groups[1].Value.Trim();
                var resolved = _catalog.ResolveProduct(product);
                results.Add(new RfqLineItemRequest
                {
                    RfqId            = rfqId,
                    Mspc             = resolved?.SearchKey,
                    Product          = resolved?.Name ?? product,
                    Units            = double.TryParse(m.Groups[2].Value, out var q) ? q : null,
                    SizeOfUnits      = m.Groups[3].Value.Trim(),
                    ProcessingSource = "import",
                });
            }
        }

        return results;
    }

    // ── POST /api/rfq-import/preview-import ──────────────────────────────────
    /// <summary>
    /// Dry-run: scans both mailboxes and shows what would be written without
    /// touching SharePoint.  source=graph uses Graph API (store@); source=com uses
    /// Outlook COM (hackensack@).  Returns first N results.
    /// </summary>
    [HttpGet("rfq-import/preview-import")]
    public async Task<IActionResult> PreviewImport(
        [FromQuery] string mailbox,
        [FromQuery] string source = "graph",
        [FromQuery] int    days   = 7,
        [FromQuery] int    top    = 5)
    {
        if (string.IsNullOrWhiteSpace(mailbox))
            return BadRequest(new { error = "mailbox query param is required" });

        List<RfqScanEmailDto> emails;
        try
        {
            if (source.Equals("com", StringComparison.OrdinalIgnoreCase))
            {
#pragma warning disable CA1416
                emails = await _comPoller.ScanRfqSentItemsAsync(mailbox, days);
#pragma warning restore CA1416
            }
            else
            {
                emails = await _mail.ScanRfqFolderAsync(mailbox, "Sent Items", days);
            }
        }
        catch (Exception ex)
        {
            return StatusCode(500, new { error = ex.Message });
        }

        var preview = emails.Take(top).Select(e =>
        {
            var requester = ParseRequesterName(e.BodyText);
            var lineItems = ParseRfqLineItems(e.RfqId, e.BodyText);
            return new
            {
                rfqId            = e.RfqId,
                sentAt           = e.SentAt,
                mailbox          = e.MailboxSource,
                requester_parsed = requester ?? e.Requester,
                emailRecipients  = e.EmailRecipients,
                lineItems        = lineItems.Select(li => new
                {
                    mspc        = li.Mspc,
                    product     = li.Product,
                    units       = li.Units,
                    sizeOfUnits = li.SizeOfUnits,
                }),
            };
        });

        return Ok(new { scanned = emails.Count, showing = Math.Min(top, emails.Count), preview });
    }

    // ── GET /api/publish/version ─────────────────────────────────────────────
    /// <summary>
    /// Returns the version string from version.txt in the SharePoint publish folder.
    /// Clients use this to detect when a newer build is available.
    /// </summary>
    [HttpGet("publish/version")]
    public async Task<IActionResult> GetPublishVersion()
    {
        try
        {
            var version = await _sp.GetPublishVersionAsync();
            return Ok(new { version });
        }
        catch (Exception ex)
        {
            _log.LogError(ex, "GetPublishVersion failed");
            return StatusCode(500, new { error = ex.Message });
        }
    }

    // ── GET /api/publish/file ─────────────────────────────────────────────────
    /// <summary>
    /// Downloads a named file from the SharePoint publish folder.
    /// Only simple filenames are accepted — no path separators.
    /// </summary>
    [HttpGet("publish/file")]
    public async Task<IActionResult> GetPublishFile([FromQuery] string name)
    {
        if (string.IsNullOrWhiteSpace(name))
            return BadRequest(new { error = "name query param is required" });
        try
        {
            var (contentType, bytes, fileName) = await _sp.GetPublishFileAsync(name);
            return File(bytes, contentType, fileName);
        }
        catch (ArgumentException ex)
        {
            return BadRequest(new { error = ex.Message });
        }
        catch (Exception ex)
        {
            _log.LogError(ex, "GetPublishFile failed for '{Name}'", name);
            return StatusCode(500, new { error = ex.Message });
        }
    }

    // ── POST /api/rfq/submit ─────────────────────────────────────────────────
    /// <summary>
    /// Called by the Excel VBA macro to submit a new RFQ Reference and its line items
    /// in a single call.  The proxy writes to SharePoint; no SP credentials are needed
    /// in the macro.  The Shredder UI polls for changes and will show a toast within ~5 s.
    /// </summary>
    [HttpPost("rfq/submit")]
    public async Task<IActionResult> SubmitRfq([FromBody] RfqSubmitRequest req)
    {
        if (string.IsNullOrWhiteSpace(req.RfqId))
            return BadRequest(new { error = "rfqId is required" });

        // Parse the supplied date; fall back to today.
        if (!DateTime.TryParse(req.DateSent, null,
                System.Globalization.DateTimeStyles.RoundtripKind, out var dateSent))
            dateSent = DateTime.UtcNow;

        // Upsert the RFQ Reference (CreateRfqReferenceAsync fills blank fields on an
        // existing row and is a no-op if all fields are already populated).
        await _sp.CreateRfqReferenceAsync(new RfqReferenceRequest
        {
            RfqId           = req.RfqId,
            Requester       = req.Requester       ?? "",
            DateSent        = dateSent,
            EmailRecipients = req.EmailRecipients ?? "",
        });

        // Create line items — skipped if this RFQ_ID already has any rows.
        int liCreated = 0, liSkipped = 0;
        if (req.LineItems.Count > 0)
        {
            var existingLiIds = await _sp.GetRfqIdsWithLineItemsAsync();
            if (!existingLiIds.Contains(req.RfqId))
            {
                var items = req.LineItems.Select(li => new RfqLineItemRequest
                {
                    RfqId          = req.RfqId,
                    Mspc           = li.Mspc,
                    Product        = li.Product,
                    Units          = li.Units,
                    SizeOfUnits    = li.SizeOfUnits,
                    SupplierEmails = li.SupplierEmails,
                }).ToList();
                await _sp.CreateRfqLineItemsAsync(items);
                liCreated = items.Count;
            }
            else
            {
                liSkipped = req.LineItems.Count;
            }
        }

        _log.LogInformation("[RFQ Submit] {Id}: li={LiCreated} created, {LiSkipped} skipped",
            req.RfqId, liCreated, liSkipped);

        return Ok(new
        {
            rfqId            = req.RfqId,
            lineItemsCreated = liCreated,
            lineItemsSkipped = liSkipped,
        });
    }

    // ── GET /api/rfq/changes ─────────────────────────────────────────────────
    /// <summary>
    /// Returns SupplierLineItems created after <paramref name="since"/> (ISO 8601 UTC),
    /// grouped by supplier.  Used by the Shredder UI 5-second change-poll loop.
    /// </summary>
    [HttpGet("rfq/changes")]
    public async Task<IActionResult> GetRfqChanges([FromQuery] string since)
    {
        if (!DateTime.TryParse(since, null, System.Globalization.DateTimeStyles.RoundtripKind, out var sinceUtc))
            return BadRequest("'since' must be an ISO 8601 UTC datetime string");

        try
        {
            var supplierTask = _sp.GetNewResponsesSinceAsync(sinceUtc);
            var rfqTask      = _sp.GetNewRfqReferencesSinceAsync(sinceUtc);
            await Task.WhenAll(supplierTask, rfqTask);

            var result     = supplierTask.Result;
            result.NewRfqs = rfqTask.Result;
            return Ok(result);
        }
        catch (Exception ex)
        {
            _log.LogError(ex, "[Changes] GetRfqChanges failed for since={Since}", since);
            return StatusCode(500, new { error = ex.Message });
        }
    }

    // ── GET /api/rfq-new/product-catalog ────────────────────────────────────
    /// <summary>
    /// Returns all rows from the Product Catalog SP list.
    /// Cached by the Shredder client for the lifetime of the app session.
    /// </summary>
    [HttpGet("rfq-new/product-catalog")]
    public async Task<IActionResult> GetRfqNewProductCatalog()
    {
        try
        {
            var items = await _sp.ReadProductCatalogAsync();
            _log.LogInformation("[RfqNew] ProductCatalog: {Count} items", items.Count);
            return Ok(items);
        }
        catch (Exception ex)
        {
            _log.LogError(ex, "[RfqNew] GetProductCatalog failed");
            return StatusCode(500, new { error = ex.Message });
        }
    }

    // ── GET /api/rfq-new/product-categories ──────────────────────────────────
    /// <summary>
    /// Returns all ProductCategory values from the Metals SP list, sorted.
    /// </summary>
    [HttpGet("rfq-new/product-categories")]
    public async Task<IActionResult> GetRfqNewProductCategories()
    {
        try
        {
            var cats = await _sp.ReadMetalCategoriesAsync();
            _log.LogInformation("[RfqNew] ProductCategories: {Count} values", cats.Count);
            return Ok(cats);
        }
        catch (Exception ex)
        {
            _log.LogError(ex, "[RfqNew] GetProductCategories failed");
            return StatusCode(500, new { error = ex.Message });
        }
    }

    // ── GET /api/rfq-new/product-shapes ──────────────────────────────────────
    /// <summary>
    /// Returns all ProductShape values from the Shapes SP list, sorted.
    /// </summary>
    [HttpGet("rfq-new/product-shapes")]
    public async Task<IActionResult> GetRfqNewProductShapes()
    {
        try
        {
            var shapes = await _sp.ReadProductShapesAsync();
            _log.LogInformation("[RfqNew] ProductShapes: {Count} values", shapes.Count);
            return Ok(shapes);
        }
        catch (Exception ex)
        {
            _log.LogError(ex, "[RfqNew] GetProductShapes failed");
            return StatusCode(500, new { error = ex.Message });
        }
    }

    // ── GET /api/rfq-new/supplier-relationships ───────────────────────────────
    /// <summary>
    /// Returns all rows from the Supplier Relationships SP list.
    /// Each row has SupplierName, Email, Metal (primary), Shape (secondary; null = any shape).
    /// The Shredder client uses these to resolve BCC recipients per product metal+shape.
    /// </summary>
    [HttpGet("rfq-new/supplier-relationships")]
    public async Task<IActionResult> GetRfqNewSupplierRelationships()
    {
        try
        {
            var rels = await _sp.ReadSupplierRelationshipsAsync();
            _log.LogInformation("[RfqNew] SupplierRelationships: {Count} rows", rels.Count);
            return Ok(rels);
        }
        catch (Exception ex)
        {
            _log.LogError(ex, "[RfqNew] GetSupplierRelationships failed");
            return StatusCode(500, new { error = ex.Message });
        }
    }

    // ── GET /api/purchase-orders ─────────────────────────────────────────────
    /// <summary>
    /// Returns all purchase order records from the PurchaseOrders SharePoint list.
    /// Shredder loads these on startup to restore purchase-marker state across sessions.
    /// </summary>
    [HttpGet("purchase-orders")]
    public async Task<IActionResult> GetPurchaseOrders()
    {
        try
        {
            var records = await _sp.ReadPurchaseOrdersAsync();
            return Ok(records);
        }
        catch (Exception ex)
        {
            _log.LogError(ex, "GetPurchaseOrders failed");
            return StatusCode(500, new { error = ex.Message });
        }
    }

    // ── DELETE /api/purchase-orders/clean ────────────────────────────────────
    /// <summary>
    /// Deletes all rows from the PurchaseOrders SharePoint list.
    /// Use before reprocessing PO emails to avoid duplicate records.
    /// </summary>
    [HttpDelete("purchase-orders/clean")]
    public async Task<IActionResult> CleanPurchaseOrders()
    {
        try
        {
            var deleted = await _sp.CleanPurchaseOrdersAsync();
            _log.LogWarning("[Clean] Deleted {Count} PurchaseOrder row(s)", deleted);
            return Ok(new { deleted });
        }
        catch (Exception ex)
        {
            _log.LogError(ex, "CleanPurchaseOrders failed");
            return StatusCode(500, new { error = ex.Message });
        }
    }

    // ── POST /api/mail/reset-po ───────────────────────────────────────────────
    /// <summary>
    /// Removes "RFQ-Processed" and PO-specific categories from PO emails in
    /// the given mailbox's inbox for the last <paramref name="days"/> days.
    /// After this call the next poll cycle will re-run PO extraction on those emails.
    /// </summary>
    [HttpPost("mail/reset-po")]
    public async Task<IActionResult> ResetPoMailCategories(
        [FromQuery] string? mailbox,
        [FromQuery] int     days = 7)
    {
        var mb = mailbox
            ?? HttpContext.RequestServices.GetRequiredService<IConfiguration>()["Mail:MailboxAddress"];
        if (string.IsNullOrEmpty(mb))
            return BadRequest(new { error = "mailbox query param or Mail:MailboxAddress config required" });

        try
        {
            var unmarked = await _mail.UnmarkPoEmailsAsync(mb, days);
            _log.LogInformation("[Mail] reset-po: unmarked {Count} PO email(s) in {Mailbox} (last {Days} days)",
                unmarked, mb, days);
            return Ok(new { unmarked, mailbox = mb, days });
        }
        catch (Exception ex)
        {
            _log.LogError(ex, "ResetPoMailCategories failed");
            return StatusCode(500, new { error = ex.Message });
        }
    }

    // ── POST /api/mail/reset-po-com ──────────────────────────────────────────
    /// <summary>
    /// Removes "PO-COM-Processed" / "PO-COM-NoExtract" from PO emails in the COM mailbox
    /// (hackensack) Sent Items for the last <paramref name="days"/> days via Outlook COM,
    /// then triggers an immediate poll cycle to re-extract them.
    /// Requires Outlook to be running with the hackensack account open.
    /// </summary>
    [HttpPost("mail/reset-po-com")]
    [System.Runtime.Versioning.SupportedOSPlatform("windows")]
    public async Task<IActionResult> ResetPoComCategories([FromQuery] int days = 7)
    {
        var mailbox = _config["OutlookCom:Mailbox"];
        if (string.IsNullOrEmpty(mailbox))
            return BadRequest(new { error = "OutlookCom:Mailbox not configured" });

        try
        {
            var unstamped = await _comPoller.UnstampPoMessagesAsync(mailbox, days);
            _comPoller.TriggerReprocess();
            _log.LogInformation("[OutlookCOM] reset-po-com: unstamped {Count} PO email(s) in {Mailbox} (last {Days} days)",
                unstamped, mailbox, days);
            return Ok(new { unstamped, mailbox, days });
        }
        catch (Exception ex)
        {
            _log.LogError(ex, "ResetPoComCategories failed");
            return StatusCode(500, new { error = ex.Message });
        }
    }

    // ── POST /api/purchase-orders/backfill-rli ───────────────────────────────
    /// <summary>
    /// Reads every PurchaseOrders row and runs UpdateRliPurchaseStatus + completion check
    /// for each one with a known RFQ ID.  Idempotent — already-marked SLI rows are skipped.
    /// </summary>
    [HttpPost("purchase-orders/backfill-rli")]
    public async Task<IActionResult> BackfillRliPurchaseStatus(
        [FromQuery] int? days,
        CancellationToken ct)
    {
        try
        {
            var (processed, skipped) = await _sp.BackfillRliPurchaseStatusAsync(days, ct);
            return Ok(new { processed, skipped });
        }
        catch (OperationCanceledException)
        {
            return StatusCode(499, new { error = "Cancelled" });
        }
        catch (Exception ex)
        {
            _log.LogError(ex, "BackfillRliPurchaseStatus failed");
            return StatusCode(500, new { error = ex.Message });
        }
    }

    // ── POST /api/purchase-orders/reextract ──────────────────────────────────
    /// <summary>
    /// Re-runs AI extraction on each PO record's original email, updates the stored
    /// LineItems JSON in SharePoint, then re-runs RLI matching. Use this when the extraction
    /// prompt has been updated and existing records need to be refreshed.
    /// </summary>
    [HttpPost("purchase-orders/reextract")]
    [System.Runtime.Versioning.SupportedOSPlatform("windows")]
    public IActionResult ReextractPoLineItems()
    {
        // Fire-and-forget on a background thread so the HTTP client can disconnect
        // without cancelling the (potentially long-running) extraction loop.
        _ = Task.Run(async () =>
        {
            try
            {
                var (updated, skipped) = await _poller.ReextractPoLineItemsAsync(
                    _comPoller.FetchByEntryIdAsync, CancellationToken.None);
                _log.LogInformation("[POReextract] Background run complete — updated={Updated}, skipped={Skipped}",
                    updated, skipped);
            }
            catch (Exception ex)
            {
                _log.LogError(ex, "[POReextract] Background run failed");
            }
        });
        return Accepted(new { started = true, note = "Running in background — check proxy logs for progress" });
    }

    // ── POST /api/mail/backfill-message-ids ──────────────────────────────────
    /// <summary>
    /// Scans SupplierResponse rows from the last <paramref name="days"/> days that are
    /// missing a MessageId and attempts to match them to Graph messages by sender+time.
    /// Patches MessageId on matched SR rows and their child SLI rows.
    /// </summary>
    [HttpPost("mail/backfill-message-ids")]
    public async Task<IActionResult> BackfillMessageIds(
        [FromQuery] int days = 7,
        CancellationToken ct = default)
    {
        try
        {
            var (patched, skipped) = await _sp.BackfillMessageIdsAsync(_mail, days, ct);
            _log.LogInformation("[BackfillMessageIds] patched={Patched} skipped={Skipped}", patched, skipped);
            return Ok(new { patched, skipped });
        }
        catch (OperationCanceledException)
        {
            return StatusCode(499, new { error = "Cancelled" });
        }
        catch (Exception ex)
        {
            _log.LogError(ex, "BackfillMessageIds failed");
            return StatusCode(500, new { error = ex.Message });
        }
    }

    // ── POST /api/mail/backfill-quote-references ─────────────────────────────
    /// <summary>
    /// Scans SupplierResponse rows (last <paramref name="days"/> days) that have a
    /// MessageId but no QuoteReference.  Re-runs AI extraction on the original
    /// email/attachment and patches QuoteReference onto each SR and its child SLIs.
    /// </summary>
    [HttpPost("mail/backfill-quote-references")]
    public async Task<IActionResult> BackfillQuoteReferences(
        [FromQuery] int days = 90,
        CancellationToken ct = default)
    {
        var mailbox = _config["Mail:MailboxAddress"];
        if (string.IsNullOrWhiteSpace(mailbox))
            return BadRequest(new { error = "Mail:MailboxAddress not configured" });
        try
        {
            var (patched, skipped) = await _sp.BackfillQuoteReferencesAsync(_mail, _aiFactory.GetService(), mailbox, days, ct);
            _log.LogInformation("[BackfillQuoteRefs] patched={Patched} skipped={Skipped}", patched, skipped);
            return Ok(new { patched, skipped });
        }
        catch (OperationCanceledException)
        {
            return StatusCode(499, new { error = "Cancelled" });
        }
        catch (Exception ex)
        {
            _log.LogError(ex, "BackfillQuoteReferences failed");
            return StatusCode(500, new { error = ex.Message });
        }
    }

    // ── POST /api/mail/deduplicate ────────────────────────────────────────────
    /// <summary>
    /// Removes orphan SupplierResponse rows (no MessageId) and collapses duplicate rows
    /// that share the same MessageId.  For duplicate groups keeps the highest-scoring row
    /// (attachment source scores highest).  Deletes child SLI rows for doomed SRs.
    /// </summary>
    [HttpPost("mail/deduplicate")]
    public async Task<IActionResult> DeduplicateSupplierResponses(
        [FromQuery] int days = 7,
        CancellationToken ct = default)
    {
        try
        {
            var (srDeleted, sliDeleted) = await _sp.DeduplicateSupplierResponsesAsync(days, ct);
            _log.LogInformation("[Deduplicate] srDeleted={SrDeleted} sliDeleted={SliDeleted}", srDeleted, sliDeleted);
            return Ok(new { srDeleted, sliDeleted });
        }
        catch (OperationCanceledException)
        {
            return StatusCode(499, new { error = "Cancelled" });
        }
        catch (Exception ex)
        {
            _log.LogError(ex, "DeduplicateSupplierResponses failed");
            return StatusCode(500, new { error = ex.Message });
        }
    }

    public record ReparentSrRequest(string RfqId);
}
