using Microsoft.AspNetCore.Mvc;
using OutlookShredder.Proxy.Services;

namespace OutlookShredder.Proxy.Controllers;

[ApiController]
[Route("api/qc")]
public class QcController : ControllerBase
{
    private readonly SharePointService      _sp;
    private readonly PricingAnalysisService _pricing;
    private readonly ProductCatalogService  _catalog;

    public QcController(SharePointService sp, PricingAnalysisService pricing, ProductCatalogService catalog)
    {
        _sp      = sp;
        _pricing = pricing;
        _catalog = catalog;
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
    /// Returns a pricing analysis over a date range of supplier quote SLI data.
    /// Raw SLI rows are cached per date so repeat calls are fast.
    /// Query params (pick one):
    ///   date=2026-05-08          — single day (defaults to yesterday)
    ///   startDate=&amp;endDate=  — explicit range
    ///   days=N                   — last N calendar days ending yesterday (default 1)
    /// Response: PricingReport — grouped by metal/shape/conditions with avg/min/max $/lb.
    /// Only high-confidence quotes contribute to the average; low/medium are logged separately.
    /// </summary>
    [HttpGet("pricing-report")]
    public async Task<IActionResult> GetPricingReportAsync(
        [FromQuery] string? date,
        [FromQuery] string? startDate,
        [FromQuery] string? endDate,
        [FromQuery] int     days = 7,
        CancellationToken   ct   = default)
    {
        var yesterday = DateOnly.FromDateTime(DateTime.UtcNow.AddDays(-1));
        DateOnly start, end;

        if (DateOnly.TryParse(startDate, out var s) && DateOnly.TryParse(endDate, out var e))
        {
            start = s;
            end   = e;
        }
        else if (DateOnly.TryParse(date, out var d))
        {
            start = end = d;
        }
        else
        {
            end   = yesterday;
            start = end.AddDays(-(days - 1));
        }

        var report = await _pricing.GetRangeReportAsync(start, end, ct);
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

    /// <summary>
    /// Analyses SLI pricing coverage for the given rolling window.
    /// Reports currently-priced rows, rows unlocked by existing catalog weights, rows that
    /// would unlock if catalog WeightPerFoot was populated, and rows with no usable price.
    /// Query param: days (default 14).
    /// </summary>
    [HttpGet("lq-analysis")]
    public async Task<IActionResult> LqAnalysisAsync([FromQuery] int days = 14)
    {
        var result = await _sp.AnalyzeLqAsync(days);
        return Ok(result);
    }

    /// <summary>
    /// Clears all LQ columns (LQ, LQ Count, LQ Min, LQ Max, LQ Long, LQ Long Count, LQ Long Min,
    /// LQ Long Max) on every QC list item.  Returns { cleared: N }.
    /// </summary>
    [HttpPost("clear-lq")]
    public async Task<IActionResult> ClearLqAsync()
    {
        var cleared = await _sp.ClearQcLqAsync();
        return Ok(new { cleared });
    }
}

public record QcRowUpdateRequest(string ItemId, string? Qc, string? QcCut);
