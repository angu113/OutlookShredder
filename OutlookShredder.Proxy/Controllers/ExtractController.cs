using Microsoft.AspNetCore.Mvc;
using OutlookShredder.Proxy.Models;
using OutlookShredder.Proxy.Services;

namespace OutlookShredder.Proxy.Controllers;

[ApiController]
[Route("api")]
public class ExtractController : ControllerBase
{
    private readonly ClaudeService            _claude;
    private readonly SharePointService        _sp;
    private readonly MailService              _mail;
    private readonly MailPollerService        _poller;
    private readonly SupplierCacheService     _suppliers;
    private readonly RfqNotificationService   _notifications;
    private readonly ILogger<ExtractController> _log;

    public ExtractController(
        ClaudeService             claude,
        SharePointService         sp,
        MailService               mail,
        MailPollerService         poller,
        SupplierCacheService      suppliers,
        RfqNotificationService    notifications,
        ILogger<ExtractController> log)
    {
        _claude        = claude;
        _sp            = sp;
        _mail          = mail;
        _poller        = poller;
        _suppliers     = suppliers;
        _notifications = notifications;
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

        RfqExtraction? extraction;
        try
        {
            extraction = await _claude.ExtractAsync(req);
        }
        catch (Exception ex)
        {
            _log.LogError(ex, "Claude extraction failed");
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
        [FromQuery] int     top      = 5000,
        [FromQuery] string? nextLink = null)
    {
        try
        {
            var (items, next) = await _sp.ReadSupplierItemsAsync(top, nextLink);
            // Return a wrapper so clients can detect whether more pages exist.
            // nextLink is null when this is the last (or only) page.
            return Ok(new { items, nextLink = next });
        }
        catch (Exception ex)
        {
            _log.LogError(ex, "Failed to read supplier items");
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
    public async Task<IActionResult> GetRfqReferences()
    {
        try
        {
            var refs = await _sp.ReadRfqReferencesAsync();
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
    /// (Claude + SharePoint upsert + re-stamp "RFQ-Processed").
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

    // ── POST /api/rfq-import/reference ───────────────────────────────────────
    /// <summary>Creates one RFQ Reference row. No-op check is the caller's responsibility.</summary>
    [HttpPost("rfq-import/reference")]
    public async Task<IActionResult> CreateRfqReference([FromBody] RfqReferenceRequest req)
    {
        try
        {
            await _sp.CreateRfqReferenceAsync(req);
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
            var result = items.Select(x => new { rfqId = x.RfqId, mspc = x.Mspc, product = x.Product, units = x.Units, sizeOfUnits = x.SizeOfUnits }).ToList();
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
                await _sp.CreateRfqLineItemsAsync(toCreate);
            return Ok(new { created = toCreate.Count, skipped = items.Count - toCreate.Count });
        }
        catch (Exception ex)
        {
            _log.LogError(ex, "CreateRfqLineItems failed");
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

    // ── GET /api/health ──────────────────────────────────────────────────────
    [HttpGet("health")]
    public IActionResult Health() =>
        Ok(new { status = "ok", utc = DateTime.UtcNow });

    // ── GET /api/version ─────────────────────────────────────────────────────
    [HttpGet("version")]
    public IActionResult Version()
    {
        var asm = System.Reflection.Assembly.GetExecutingAssembly();
        var attr = (System.Reflection.AssemblyInformationalVersionAttribute?)
            System.Attribute.GetCustomAttribute(asm, typeof(System.Reflection.AssemblyInformationalVersionAttribute));
        return Ok(new { version = attr?.InformationalVersion ?? "unknown" });
    }
}
