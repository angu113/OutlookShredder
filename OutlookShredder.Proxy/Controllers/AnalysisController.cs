using Microsoft.AspNetCore.Mvc;
using OutlookShredder.Proxy.Services;

namespace OutlookShredder.Proxy.Controllers;

[ApiController]
[Route("api/analysis")]
public class AnalysisController(CatalogAnalysisService analysis) : ControllerBase
{
    /// <summary>Snapshot in-memory product catalog → analysis-cache/catalog.json</summary>
    [HttpPost("catalog/fetch")]
    public IActionResult FetchCatalog()
        => Ok(analysis.FetchCatalog());

    /// <summary>Snapshot SLI rows with a known ProductSearchKey → analysis-cache/sli-sample.json</summary>
    [HttpPost("sli/fetch")]
    public async Task<IActionResult> FetchSli(CancellationToken ct)
        => Ok(await analysis.FetchSliAsync(ct));

    /// <summary>
    /// AI-tokenise the catalog or SLI source file (resumable).
    /// clearShapes: comma-separated shape names to force re-tokenise (e.g. "sheet,plate,pipe")
    /// </summary>
    [HttpPost("catalog/tokenize")]
    public async Task<IActionResult> TokenizeCatalog(
        [FromQuery] string? clearShapes, CancellationToken ct)
        => Ok(await analysis.TokenizeAsync("catalog",
            clearShapes?.Split(',', StringSplitOptions.RemoveEmptyEntries) ?? [], ct));

    [HttpPost("sli/tokenize")]
    public async Task<IActionResult> TokenizeSli(
        [FromQuery] string? clearShapes, CancellationToken ct)
        => Ok(await analysis.TokenizeAsync("sli",
            clearShapes?.Split(',', StringSplitOptions.RemoveEmptyEntries) ?? [], ct));

    /// <summary>
    /// Score sli-tokens against catalog-tokens in-memory. No network calls.
    /// limit: max SLI rows to test (0 = all, default 500)
    /// </summary>
    [HttpGet("match-test")]
    public async Task<IActionResult> MatchTest([FromQuery] int limit = 500, CancellationToken ct = default)
        => Ok(await analysis.RunMatchTestAsync(limit, ct));

    /// <summary>Cache file sizes and modification times</summary>
    [HttpGet("status")]
    public IActionResult Status()
        => Ok(analysis.GetStatus());

    /// <summary>
    /// Mine catalog-tokens + sli-tokens for unique alloy/condition/temper values and upsert
    /// them into the SP IndustryDictionary list. Safe to re-run — existing entries are updated
    /// with fresh occurrence counts; Definition/Examples are never overwritten.
    /// </summary>
    [HttpPost("dictionary/build")]
    public async Task<IActionResult> BuildDictionary(CancellationToken ct)
        => Ok(await analysis.BuildDictionaryAsync(ct));

    /// <summary>Read all IndustryDictionary entries from SharePoint</summary>
    [HttpGet("dictionary")]
    public async Task<IActionResult> GetDictionary(CancellationToken ct)
        => Ok(await analysis.ReadDictionaryAsync(ct));

    /// <summary>
    /// Preview or apply GT audit patches to SharePoint SLI ProductSearchKey values.
    /// Reads match-test-results.json, classifies each miss (clear/update/review),
    /// resolves SpItemId from sli-sample.json or live SLI cache, and optionally patches SP.
    ///
    /// dryRun=true (default): returns planned actions without touching SP.
    /// dryRun=false: applies clear+update actions to SP.
    /// action: "both" (default) | "clear" | "update" | "review" | "all"
    ///   "both"   = show/apply clear + update (skip review)
    ///   "clear"  = only clear actions
    ///   "update" = only update actions
    ///   "review" = only review bucket (always dry-run)
    ///   "all"    = all three buckets (review bucket is always dry-run)
    /// </summary>
    [HttpPost("gt-audit/apply")]
    public async Task<IActionResult> ApplyGtAudit(
        [FromQuery] bool dryRun = true,
        [FromQuery] string action = "both",
        CancellationToken ct = default)
        => Ok(await analysis.ApplyGtAuditAsync(dryRun, action, ct));
}
