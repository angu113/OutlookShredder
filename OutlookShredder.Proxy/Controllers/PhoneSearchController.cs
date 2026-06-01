using Microsoft.AspNetCore.Cors;
using Microsoft.AspNetCore.Mvc;

namespace OutlookShredder.Proxy.Controllers;

/// <summary>
/// Relay endpoint used by the Shredder desktop app and the OpenBravo Chrome extension
/// so clicking a phone chip in Shredder auto-fills the Sales Order contact filter in the browser,
/// and clicking a customer name chip fills the Customer (business partner) filter.
///
/// Phone flow:    Shredder POST /api/phone-search         → extension GET /api/phone-search/pending
///                (extension fills Contact box)            → extension POST /api/phone-search/consume
/// Customer flow: Shredder POST /api/phone-search/customer → extension GET /api/phone-search/pending-customer
///                (extension fills Customer box)           → extension POST /api/phone-search/consume-customer
/// Document flow: Shredder POST /api/phone-search/document → extension GET /api/phone-search/pending-document
///                (extension fills Doc # filter box)       → extension POST /api/phone-search/consume-document
///
/// CORS: "PhoneSearch" policy allows any origin so the extension content script
/// (running on the OpenBravo origin) can call localhost:7000 without a CORS block.
/// </summary>
[ApiController]
[Route("api/phone-search")]
[EnableCors("PhoneSearch")]
public class PhoneSearchController : ControllerBase
{
    private static string? _pending;
    private static string? _pendingCustomer;
    private static string? _pendingDocument;
    private static string? _pendingPoDocument;
    private static string? _pendingNewPo;

    [HttpOptions]
    [HttpOptions("pending")]
    [HttpOptions("consume")]
    [HttpOptions("customer")]
    [HttpOptions("pending-customer")]
    [HttpOptions("consume-customer")]
    [HttpOptions("document")]
    [HttpOptions("pending-document")]
    [HttpOptions("consume-document")]
    [HttpOptions("po-document")]
    [HttpOptions("pending-po-document")]
    [HttpOptions("consume-po-document")]
    [HttpOptions("new-po")]
    [HttpOptions("pending-new-po")]
    [HttpOptions("consume-new-po")]
    [HttpOptions("scan")]
    [HttpOptions("scan-result")]
    public IActionResult Preflight() => NoContent();

    // ── Phone / Contact ───────────────────────────────────────────────────────

    [HttpPost]
    public IActionResult Set([FromBody] PhoneSearchRequest req)
    {
        if (string.IsNullOrWhiteSpace(req.Phone))
            return BadRequest(new { error = "phone is required" });
        Volatile.Write(ref _pending, req.Phone);
        return Ok(new { ok = true });
    }

    [HttpGet("pending")]
    public IActionResult GetPending()
    {
        var phone = Volatile.Read(ref _pending);
        return Ok(new { phone });
    }

    [HttpPost("consume")]
    public IActionResult Consume()
    {
        Volatile.Write(ref _pending, null);
        return Ok(new { ok = true });
    }

    // ── Customer / Business Partner ───────────────────────────────────────────

    [HttpPost("customer")]
    public IActionResult SetCustomer([FromBody] CustomerSearchRequest req)
    {
        if (string.IsNullOrWhiteSpace(req.CustomerName))
            return BadRequest(new { error = "customerName is required" });
        Volatile.Write(ref _pendingCustomer, req.CustomerName);
        return Ok(new { ok = true });
    }

    [HttpGet("pending-customer")]
    public IActionResult GetPendingCustomer()
    {
        var customerName = Volatile.Read(ref _pendingCustomer);
        return Ok(new { customerName });
    }

    [HttpPost("consume-customer")]
    public IActionResult ConsumeCustomer()
    {
        Volatile.Write(ref _pendingCustomer, null);
        return Ok(new { ok = true });
    }

    // ── Document # (Sales Order Doc # filter) ──────────────────────────────────

    [HttpPost("document")]
    public IActionResult SetDocument([FromBody] DocumentSearchRequest req)
    {
        if (string.IsNullOrWhiteSpace(req.DocumentNo))
            return BadRequest(new { error = "documentNo is required" });
        Volatile.Write(ref _pendingDocument, req.DocumentNo);
        return Ok(new { ok = true });
    }

    [HttpGet("pending-document")]
    public IActionResult GetPendingDocument()
    {
        var documentNo = Volatile.Read(ref _pendingDocument);
        return Ok(new { documentNo });
    }

    [HttpPost("consume-document")]
    public IActionResult ConsumeDocument()
    {
        Volatile.Write(ref _pendingDocument, null);
        return Ok(new { ok = true });
    }

    // ── PO Document # (Purchase Order Doc # filter) ────────────────────────────
    // Same as the Sales Order document flow, but the extension opens the Purchase
    // Order window (windowId 181) instead of Sales Order. Used by HSK-PO clicks.

    [HttpPost("po-document")]
    public IActionResult SetPoDocument([FromBody] DocumentSearchRequest req)
    {
        if (string.IsNullOrWhiteSpace(req.DocumentNo))
            return BadRequest(new { error = "documentNo is required" });
        Volatile.Write(ref _pendingPoDocument, req.DocumentNo);
        return Ok(new { ok = true });
    }

    [HttpGet("pending-po-document")]
    public IActionResult GetPendingPoDocument()
    {
        var documentNo = Volatile.Read(ref _pendingPoDocument);
        return Ok(new { documentNo });
    }

    [HttpPost("consume-po-document")]
    public IActionResult ConsumePoDocument()
    {
        Volatile.Write(ref _pendingPoDocument, null);
        return Ok(new { ok = true });
    }

    // ── New Purchase Order (open the PO window in new-record mode) ──────────────
    // No filter fill — the extension opens a fresh blank PO form. Supplier is
    // carried for future pre-fill; presence of a value is the trigger.

    [HttpPost("new-po")]
    public IActionResult SetNewPo([FromBody] NewPoRequest? req)
    {
        // Supplier is optional; use a sentinel so an empty body still triggers a new PO.
        Volatile.Write(ref _pendingNewPo, string.IsNullOrWhiteSpace(req?.Supplier) ? "1" : req!.Supplier);
        return Ok(new { ok = true });
    }

    [HttpGet("pending-new-po")]
    public IActionResult GetPendingNewPo()
    {
        var newPo = Volatile.Read(ref _pendingNewPo);
        return Ok(new { newPo });
    }

    [HttpPost("consume-new-po")]
    public IActionResult ConsumeNewPo()
    {
        Volatile.Write(ref _pendingNewPo, null);
        return Ok(new { ok = true });
    }

    // ── Page scanner (dev tool) ───────────────────────────────────────────────
    // Shredder POSTs to /scan to request a scan; the extension GETs /scan,
    // runs scanPage(), and POSTs the JSON back to /scan-result.
    // Shredder polls /scan-result until it appears.

    private static volatile bool _pendingScan;
    private static string? _scanResult;

    [HttpPost("scan")]
    public IActionResult RequestScan()
    {
        Volatile.Write(ref _scanResult, null);
        _pendingScan = true;
        return Ok(new { ok = true });
    }

    [HttpGet("scan")]
    public IActionResult GetScanRequest()
        => Ok(new { scan = _pendingScan });

    [HttpPost("scan-result")]
    public async Task<IActionResult> PostScanResult()
    {
        using var reader = new System.IO.StreamReader(Request.Body);
        var json = await reader.ReadToEndAsync();
        Volatile.Write(ref _scanResult, json);
        _pendingScan = false;
        return Ok(new { ok = true });
    }

    [HttpGet("scan-result")]
    public IActionResult GetScanResult()
        => Ok(new { result = Volatile.Read(ref _scanResult) });

    public record PhoneSearchRequest(string Phone);
    public record CustomerSearchRequest(string CustomerName);
    public record DocumentSearchRequest(string DocumentNo);
    public record NewPoRequest(string? Supplier);
}
