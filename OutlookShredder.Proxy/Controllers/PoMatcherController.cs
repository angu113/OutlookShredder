using Microsoft.AspNetCore.Mvc;
using OutlookShredder.Proxy.Services;

namespace OutlookShredder.Proxy.Controllers;

/// <summary>On-demand trigger for the PO confirmation accelerator (Fulfillment loop).</summary>
[ApiController]
[Route("api")]
public class PoMatcherController : ControllerBase
{
    private readonly PoConfirmationMatcherService _matcher;
    private readonly ILogger<PoMatcherController>  _log;

    public PoMatcherController(PoConfirmationMatcherService matcher, ILogger<PoMatcherController> log)
    {
        _matcher = matcher; _log = log;
    }

    /// <summary>Run one PO &lt;-&gt; confirmation-email matching pass now. Returns
    /// { scanned, matched, confirmed, auto, details[] }.</summary>
    [HttpPost("purchase-orders/match-confirmations")]
    public async Task<IActionResult> MatchConfirmations(CancellationToken ct)
    {
        try { return Ok(await _matcher.RunOnceAsync(ct)); }
        catch (Exception ex)
        {
            _log.LogError(ex, "match-confirmations failed");
            return StatusCode(500, new { error = ex.Message });
        }
    }
}
