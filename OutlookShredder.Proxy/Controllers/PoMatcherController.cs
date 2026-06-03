using Microsoft.AspNetCore.Mvc;
using OutlookShredder.Proxy.Services;

namespace OutlookShredder.Proxy.Controllers;

/// <summary>On-demand trigger for the PO confirmation accelerator (Fulfillment loop).</summary>
[ApiController]
[Route("api")]
public class PoMatcherController : ControllerBase
{
    private readonly PoConfirmationMatcherService _matcher;
    private readonly BillToPoMatcherService        _billMatcher;
    private readonly ILogger<PoMatcherController>  _log;

    public PoMatcherController(PoConfirmationMatcherService matcher, BillToPoMatcherService billMatcher,
        ILogger<PoMatcherController> log)
    {
        _matcher = matcher; _billMatcher = billMatcher; _log = log;
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

    /// <summary>Run one bill &lt;-&gt; PO matching pass now. apply=false (default) suggests only;
    /// apply=true writes PaymentStatus=Required + the bill pointer for confident single matches.
    /// Returns { scanned, matched, applied, auto, matches[], ambiguous[] }.</summary>
    [HttpPost("purchase-orders/match-bills")]
    public async Task<IActionResult> MatchBills([FromQuery] bool apply = false, CancellationToken ct = default)
    {
        try { return Ok(await _billMatcher.RunOnceAsync(apply, ct)); }
        catch (Exception ex)
        {
            _log.LogError(ex, "match-bills failed");
            return StatusCode(500, new { error = ex.Message });
        }
    }
}
