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
}
