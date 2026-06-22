using Microsoft.AspNetCore.Mvc;
using OutlookShredder.Proxy.Models.CutOptimizer;
using OutlookShredder.Proxy.Services.CutOptimizer;

namespace OutlookShredder.Proxy.Controllers;

/// <summary>
/// Cut-optimizer endpoint for the Pixar "Cut Optimizer" tab. Takes the parts to cut + the available
/// stock and returns a cut plan (text summary + structured layout; PDF report added in Phase 3).
/// Parallel to <see cref="DrawingController"/> — all geometry/optimization is proxy-side; the client is thin.
/// </summary>
[ApiController]
public class CutOptimizerController : ControllerBase
{
    private readonly ILogger<CutOptimizerController> _log;

    public CutOptimizerController(ILogger<CutOptimizerController> log) => _log = log;

    [HttpPost("/api/cut-optimizer/optimize")]
    public IActionResult Optimize([FromBody] OptimizeRequest req)
    {
        if (req is null) return BadRequest(new { error = "request body is required" });
        try
        {
            var result = CutOptimizerService.Optimize(req);
            return Ok(result);
        }
        catch (Exception ex)
        {
            _log.LogWarning(ex, "[CutOptimizer] optimize failed");
            return StatusCode(500, new { error = ex.Message });
        }
    }
}
