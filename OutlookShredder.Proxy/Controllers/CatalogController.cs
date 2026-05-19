using Microsoft.AspNetCore.Mvc;
using OutlookShredder.Proxy.Services;

namespace OutlookShredder.Proxy.Controllers;

[ApiController]
[Route("api/catalog")]
public class CatalogController(ProductCatalogService catalog, SharePointService sp) : ControllerBase
{
    /// <summary>
    /// GET /api/catalog — lists the products currently loaded in the in-memory cache.
    /// Add ?detail=true to include searchKeys alongside product names.
    /// Add ?search=foo to filter by substring (case-insensitive).
    /// </summary>
    [HttpGet]
    public IActionResult GetAll([FromQuery] bool detail = false, [FromQuery] string? search = null)
    {
        var entries = catalog.CachedEntries;
        if (!string.IsNullOrWhiteSpace(search))
            entries = entries.Where(e => e.Name.Contains(search, StringComparison.OrdinalIgnoreCase)).ToList();

        return Ok(new
        {
            count       = catalog.CachedNames.Count,
            lastRefresh = catalog.LastRefreshAt,
            lastError   = catalog.LastError,
            products    = detail
                ? (object)entries.Select(e => new { e.Name, e.SearchKey })
                : entries.Select(e => e.Name),
        });
    }

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

        // Auto-compute weights for newly added items (non-fatal)
        int weightsComputed = 0;
        try
        {
            var wr = await catalog.ComputeWeightsAsync(dryRun: false, overwrite: false, ct: ct);
            weightsComputed = wr.Computed;
        }
        catch (Exception ex)
        {
            // Weight computation is best-effort — column may not exist yet
            _ = ex;
        }

        return Ok(new { added, skipped, total = added + skipped, weightsComputed });
    }

    /// <summary>
    /// POST /api/catalog/dedup?dryRun=true — removes duplicate SP rows sharing the same
    /// SearchKey (MSPC). Keeps the row with the lowest SP item ID (first written).
    /// Pass ?dryRun=false (or omit) to actually delete; ?dryRun=true to preview.
    /// </summary>
    [HttpPost("dedup")]
    public async Task<IActionResult> Dedup([FromQuery] bool dryRun = false, CancellationToken ct = default)
    {
        var (groups, deleted, report) = await catalog.DedupAsync(dryRun, ct);
        return Ok(new
        {
            dryRun,
            duplicateGroups = groups,
            deleted,
            groups = report.Select(g => new
            {
                searchKey  = g.SearchKey,
                keepSpId   = g.KeeperSpId,
                keepName   = g.KeeperName,
                deleteCount = g.Extras.Count,
                deleting   = g.Extras.Select(x => new { x.SpId, x.Name }),
            }),
        });
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

        return Ok(new { raw = name, match = match.Value.Name, searchKey = match.Value.SearchKey, confidence = match.Value.Confidence });
    }

    /// <summary>
    /// GET /api/catalog/diagnose?name=foo&amp;top=10 — shows tokenisation and top-N match scores.
    /// Use to debug why a vendor string matched (or didn't match) the expected catalog entry.
    /// </summary>
    [HttpGet("diagnose")]
    public IActionResult Diagnose([FromQuery] string? name, [FromQuery] int top = 10)
    {
        if (string.IsNullOrWhiteSpace(name))
            return BadRequest("Provide a ?name= query parameter.");
        return Ok(catalog.Diagnose(name, top));
    }

    /// <summary>Idempotently creates the 'WeightPerFoot' number column on the Product Catalog list.</summary>
    [HttpPost("ensure-weight-column")]
    public async Task<IActionResult> EnsureWeightColumnAsync()
    {
        await catalog.EnsureCatalogWeightColumnAsync();
        return Ok(new { ok = true });
    }

    /// <summary>
    /// POST /api/catalog/backfill-rli-custom-ids?rfqId=HQXXXXXX
    /// For every RFQ Line Item row under the given rfqId that has no MSPC, creates a
    /// deterministic CUSTOM_ID and patches it back to SharePoint.
    /// Call this once for existing RFQs created before the CUSTOM_ID system was added.
    /// </summary>
    [HttpPost("backfill-rli-custom-ids")]
    public async Task<IActionResult> BackfillRliCustomIds(
        [FromQuery] string rfqId, CancellationToken ct)
    {
        if (string.IsNullOrWhiteSpace(rfqId))
            return BadRequest(new { error = "rfqId is required" });
        var patched = await sp.BackfillRliCustomIdsAsync(rfqId, ct);
        return Ok(new
        {
            rfqId,
            patched = patched.Count,
            items   = patched.Select(p => new { product = p.Product, customId = p.CustomId }),
        });
    }

    /// <summary>
    /// POST /api/catalog/compute-weights?dryRun=true&amp;overwrite=false
    /// Computes theoretical weight (lb/ft for linear products, lb/sqft for sheet/plate)
    /// from each product's name and PATCHes WeightPerFoot + WeightUnit on the SP list.
    /// Run POST /api/catalog/ensure-weight-column first to create the SP columns.
    /// </summary>
    [HttpPost("compute-weights")]
    public async Task<IActionResult> ComputeWeights(
        [FromQuery] bool dryRun   = true,
        [FromQuery] bool overwrite = false,
        CancellationToken ct = default)
    {
        var report = await catalog.ComputeWeightsAsync(dryRun, overwrite, ct);
        return Ok(new
        {
            dryRun,
            overwrite,
            report.Total,
            report.Computed,
            report.AlreadySet,
            report.Skipped,
            items = report.Items.Select(i => new
            {
                i.SearchKey,
                i.ProductName,
                weight    = i.WeightValue.HasValue ? $"{i.WeightValue:F4} {i.WeightUnit}" : null,
                i.Formula,
                i.Note,
                i.Updated,
            }),
        });
    }
}
