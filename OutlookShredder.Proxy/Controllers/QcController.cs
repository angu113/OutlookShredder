using Microsoft.AspNetCore.Mvc;
using OutlookShredder.Proxy.Services;

namespace OutlookShredder.Proxy.Controllers;

[ApiController]
[Route("api/qc")]
public class QcController : ControllerBase
{
    private readonly SharePointService      _sp;
    private readonly PricingAnalysisService _pricing;

    public QcController(SharePointService sp, PricingAnalysisService pricing)
    {
        _sp      = sp;
        _pricing = pricing;
    }

    /// <summary>
    /// Returns the QC SharePoint list as { columns: [...], rows: [[...], ...], lastModified: "..." }.
    /// Uses the same app-only credentials as the rest of the proxy.
    /// </summary>
    [HttpGet]
    public async Task<IActionResult> GetAsync()
    {
        var result = await _sp.ReadQcListAsync();
        return Ok(result);
    }

    /// <summary>
    /// Returns the last-modified UTC timestamp of the QC SharePoint list.
    /// Response: { lastModified: "2026-04-01T12:34:56Z" } or { lastModified: null }.
    /// </summary>
    [HttpGet("last-modified")]
    public async Task<IActionResult> GetLastModifiedAsync()
    {
        var lastModified = await _sp.GetQcLastModifiedAsync();
        return Ok(new { lastModified = lastModified?.ToString("o") });
    }

    /// <summary>
    /// Reads recent supplier quotes, derives $/lb for each, matches against QC list
    /// Metal+Shape rows, and patches the 'LQ' column.
    /// Returns { updated: [...], misses: [...] }.
    /// </summary>
    [HttpPost("update-lq")]
    public async Task<IActionResult> UpdateLqAsync()
    {
        var result = await _sp.UpdateQcLqAsync();
        return Ok(result);
    }

    /// <summary>
    /// Returns a pricing analysis for a single day of supplier quote SLI data.
    /// Raw SLI rows are cached per date so repeat calls are fast.
    /// Query param: date=2026-05-08 (ISO date; defaults to yesterday).
    /// Response: PricingReport — grouped by metal/shape/conditions with avg/min/max $/lb.
    /// Only high-confidence quotes contribute to the average; low/medium are logged and returned separately.
    /// </summary>
    [HttpGet("pricing-report")]
    public async Task<IActionResult> GetPricingReportAsync(
        [FromQuery] string? date, CancellationToken ct)
    {
        if (!DateOnly.TryParse(date, out var parsedDate))
            parsedDate = DateOnly.FromDateTime(DateTime.UtcNow.AddDays(-1));

        var report = await _pricing.GetReportAsync(parsedDate, ct);
        return Ok(report);
    }

    /// <summary>
    /// Patches the QC and QC Cut fields of a single QC list item.
    /// Body: { itemId, qc, qcCut }
    /// </summary>
    [HttpPatch("update-row")]
    public async Task<IActionResult> UpdateRowAsync([FromBody] QcRowUpdateRequest req)
    {
        await _sp.UpdateQcRowAsync(req.ItemId, req.Qc, req.QcCut);
        return Ok();
    }
}

public record QcRowUpdateRequest(string ItemId, string? Qc, string? QcCut);
