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
            _notifications.NotifyRfqProcessed();

        return Ok(new ExtractResponse
        {
            Success   = rows.Any(r => r.Success),
            Extracted = extraction!,
            Rows      = rows
        });
    }

    // ── POST /api/setup-columns ──────────────────────────────────────────────
    /// <summary>
    /// Provisions all required columns on the RFQLineItems SharePoint list.
    /// Run once after creating the blank list. Safe to re-run.
    /// </summary>
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
    /// Returns all RFQLineItems for the dashboard.
    /// Uses server-side app credentials — no browser auth token required.
    /// Each item is a flat field dictionary matching the SharePoint column names.
    /// </summary>
    [HttpGet("items")]
    public async Task<IActionResult> GetItems([FromQuery] int top = 500)
    {
        try
        {
            var items = await _sp.ReadItemsAsync(top);
            return Ok(items);
        }
        catch (Exception ex)
        {
            _log.LogError(ex, "Failed to read SharePoint items");
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
                    while (reader.TryRead(out var evt))
                        await Response.WriteAsync($"event: {evt}\ndata: {{}}\n\n", ct);
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

    // ── GET /api/health ──────────────────────────────────────────────────────
    [HttpGet("health")]
    public IActionResult Health() =>
        Ok(new { status = "ok", utc = DateTime.UtcNow });
}
