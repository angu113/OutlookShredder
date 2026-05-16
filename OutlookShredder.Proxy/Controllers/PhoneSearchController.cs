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

    [HttpOptions]
    [HttpOptions("pending")]
    [HttpOptions("consume")]
    [HttpOptions("customer")]
    [HttpOptions("pending-customer")]
    [HttpOptions("consume-customer")]
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

    public record PhoneSearchRequest(string Phone);
    public record CustomerSearchRequest(string CustomerName);
}
