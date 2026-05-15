using Microsoft.AspNetCore.Cors;
using Microsoft.AspNetCore.Mvc;

namespace OutlookShredder.Proxy.Controllers;

/// <summary>
/// Relay endpoint used by the Shredder desktop app and the OpenBravo Chrome extension
/// so clicking a phone chip in Shredder auto-fills the Sales Order contact filter in the browser.
///
/// Flow: Shredder POST /api/phone-search  →  Chrome extension GET /api/phone-search/pending
///       (extension fills the box)         →  Chrome extension POST /api/phone-search/consume
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

    [HttpOptions]
    [HttpOptions("pending")]
    [HttpOptions("consume")]
    public IActionResult Preflight() => NoContent();

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

    public record PhoneSearchRequest(string Phone);
}
