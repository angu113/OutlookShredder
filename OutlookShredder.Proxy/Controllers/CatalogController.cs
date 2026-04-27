using Microsoft.AspNetCore.Mvc;
using OutlookShredder.Proxy.Services;

namespace OutlookShredder.Proxy.Controllers;

[ApiController]
[Route("api/catalog")]
public class CatalogController(ProductCatalogService catalog, SharePointService sp) : ControllerBase
{
    /// <summary>
    /// GET /api/catalog — lists the products currently loaded in the in-memory cache,
    /// plus cache status info.
    /// </summary>
    [HttpGet]
    public IActionResult GetAll() => Ok(new
    {
        count       = catalog.CachedNames.Count,
        lastRefresh = catalog.LastRefreshAt,
        lastError   = catalog.LastError,
        products    = catalog.CachedNames,
    });

    /// <summary>
    /// POST /api/catalog/refresh — forces an immediate cache refresh.
    /// </summary>
    [HttpPost("refresh")]
    public async Task<IActionResult> Refresh()
    {
        await catalog.RefreshAsync();
        return Ok(new
        {
            count     = catalog.CachedNames.Count,
            lastError = catalog.LastError,
            diag      = catalog.LastDiag,
        });
    }

    /// <summary>
    /// POST /api/catalog/backfill — patches CatalogProductName + ProductSearchKey on every
    /// existing SupplierLineItem using the current in-memory catalog.
    /// Safe to run repeatedly (idempotent). Runs synchronously; expect ~1 s per 10 rows.
    /// </summary>
    [HttpPost("backfill")]
    public async Task<IActionResult> Backfill(CancellationToken ct)
    {
        var (total, updated, matched) = await sp.BackfillCatalogMatchesAsync(ct);
        return Ok(new { total, updated, matched });
    }

    /// <summary>
    /// POST /api/catalog/load — loads new products from a CSV file sent as the request body.
    /// Deduplicates by Line No. (MSPC); only rows absent from SharePoint are written.
    /// Send the CSV as the raw request body (Content-Type: text/csv or application/octet-stream).
    /// </summary>
    [HttpPost("load")]
    [RequestSizeLimit(50_000_000)]
    public async Task<IActionResult> Load(CancellationToken ct)
    {
        if (Request.ContentLength is 0)
            return BadRequest("Send the CSV file as the request body.");

        // Buffer into MemoryStream — Kestrel disallows synchronous reads on the request stream.
        using var ms = new System.IO.MemoryStream();
        await Request.Body.CopyToAsync(ms, ct);
        ms.Position = 0;

        var (added, skipped) = await sp.LoadCatalogFromCsvAsync(ms, ct);

        await catalog.RefreshAsync();

        return Ok(new { added, skipped, total = added + skipped });
    }

    /// <summary>
    /// GET /api/catalog/resolve?name=foo — resolves a vendor description against the cache.
    /// Returns the matched catalog entry, or 404 with the raw name if no match.
    /// </summary>
    [HttpGet("resolve")]
    public IActionResult Resolve([FromQuery] string? name)
    {
        if (string.IsNullOrWhiteSpace(name))
            return BadRequest("Provide a ?name= query parameter.");

        var match = catalog.ResolveProduct(name);
        if (match is null)
            return NotFound(new { raw = name, match = (string?)null });

        return Ok(new { raw = name, match = match.Value.Name, searchKey = match.Value.SearchKey });
    }
}
