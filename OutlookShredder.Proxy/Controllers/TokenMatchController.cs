using Microsoft.AspNetCore.Mvc;
using OutlookShredder.Proxy.Models;
using OutlookShredder.Proxy.Services;

namespace OutlookShredder.Proxy.Controllers;

/// <summary>
/// Token-match diagnostic review endpoints.
/// All write endpoints automatically trigger a cache refresh on SupplierProductMappingsCacheService
/// so promoted mappings are visible immediately without waiting for the 5-minute refresh.
/// </summary>
[ApiController]
[Route("api/token-match")]
public class TokenMatchController : ControllerBase
{
    private readonly SharePointService                   _sp;
    private readonly CatalogAnalysisService              _analysis;
    private readonly SupplierProductMappingsCacheService _mappingsCache;
    private readonly ILogger<TokenMatchController>       _log;

    public TokenMatchController(
        SharePointService                   sp,
        CatalogAnalysisService              analysis,
        SupplierProductMappingsCacheService mappingsCache,
        ILogger<TokenMatchController>       log)
    {
        _sp            = sp;
        _analysis      = analysis;
        _mappingsCache = mappingsCache;
        _log           = log;
    }

    /// <summary>GET /api/token-match/diagnostics — list diagnostic rows (filtered)</summary>
    [HttpGet("diagnostics")]
    public async Task<IActionResult> GetDiagnostics(
        [FromQuery] string? rfqId        = null,
        [FromQuery] string? reviewStatus = null,
        [FromQuery] int     top          = 200)
    {
        try
        {
            var rows = await _sp.GetTokenMatchDiagnosticsAsync(rfqId, reviewStatus, top);
            return Ok(rows);
        }
        catch (Exception ex)
        {
            _log.LogError(ex, "GetDiagnostics failed");
            return StatusCode(500, new { error = ex.Message });
        }
    }

    /// <summary>GET /api/token-match/diagnostics/stats — aggregate stats</summary>
    [HttpGet("diagnostics/stats")]
    public async Task<IActionResult> GetStats()
    {
        try
        {
            // Fetch all rows (no filter) to compute stats. Reasonable for now — thousands of rows max.
            var rows = await _sp.GetTokenMatchDiagnosticsAsync(top: 5000);
            var stats = new TokenMatchStats
            {
                Total      = rows.Count,
                Agreed     = rows.Count(r => r.Agreed),
                Disagreed  = rows.Count(r => !r.Agreed),
                Pending    = rows.Count(r => r.ReviewStatus == "pending"),
                Confirmed  = rows.Count(r => r.ReviewStatus == "confirmed"),
                Rejected   = rows.Count(r => r.ReviewStatus == "rejected"),
                Overridden = rows.Count(r => r.ReviewStatus == "overridden"),
            };
            return Ok(stats);
        }
        catch (Exception ex)
        {
            _log.LogError(ex, "GetStats failed");
            return StatusCode(500, new { error = ex.Message });
        }
    }

    /// <summary>POST /api/token-match/diagnostics/{id}/review — confirm, reject, or override</summary>
    [HttpPost("diagnostics/{id}/review")]
    public async Task<IActionResult> Review(string id, [FromBody] ReviewRequest req)
    {
        if (string.IsNullOrEmpty(req.ReviewStatus))
            return BadRequest(new { error = "reviewStatus is required" });
        if (!new[] { "confirmed", "rejected", "overridden" }.Contains(req.ReviewStatus))
            return BadRequest(new { error = "reviewStatus must be confirmed | rejected | overridden" });

        try
        {
            await _sp.ReviewTokenMatchDiagnosticAsync(id, req.ReviewStatus, req.OverriddenMspc);
            return Ok(new { success = true });
        }
        catch (Exception ex)
        {
            _log.LogError(ex, "Review failed for {Id}", id);
            return StatusCode(500, new { error = ex.Message });
        }
    }

    /// <summary>POST /api/token-match/diagnostics/{id}/promote-mapping — write to SupplierProductMappings</summary>
    [HttpPost("diagnostics/{id}/promote-mapping")]
    public async Task<IActionResult> PromoteMapping(string id)
    {
        try
        {
            await _sp.PromoteTokenMatchMappingAsync(id);
            _ = _mappingsCache.ForceRefreshAsync();
            return Ok(new { success = true });
        }
        catch (Exception ex)
        {
            _log.LogError(ex, "PromoteMapping failed for {Id}", id);
            return StatusCode(500, new { error = ex.Message });
        }
    }

    /// <summary>GET /api/token-match/resolve?name=&supplier= — live token resolve, returns top-3 candidates</summary>
    [HttpGet("resolve")]
    public async Task<IActionResult> Resolve(
        [FromQuery] string name, [FromQuery] string? supplier = null)
    {
        if (string.IsNullOrWhiteSpace(name))
            return BadRequest(new { error = "name is required" });

        try
        {
            var result = await _analysis.MatchProductAsync(name, supplier, HttpContext.RequestAborted);
            return Ok(new
            {
                searchKey      = result.SearchKey,
                catalogName    = result.CatalogName,
                score          = result.Score,
                source         = result.Source,
                topCandidates  = result.TopCandidates,
            });
        }
        catch (Exception ex)
        {
            _log.LogError(ex, "Resolve failed for '{Name}'", name);
            return StatusCode(500, new { error = ex.Message });
        }
    }

    /// <summary>DELETE /api/token-match/diagnostics — wipe all rows (dev/cleanup use only)</summary>
    [HttpDelete("diagnostics")]
    public async Task<IActionResult> ClearDiagnostics()
    {
        try
        {
            var count = await _sp.ClearTokenMatchDiagnosticsAsync();
            return Ok(new { deleted = count });
        }
        catch (Exception ex)
        {
            _log.LogError(ex, "ClearDiagnostics failed");
            return StatusCode(500, new { error = ex.Message });
        }
    }

    public class ReviewRequest
    {
        public string  ReviewStatus   { get; set; } = "";
        public string? OverriddenMspc { get; set; }
    }
}
