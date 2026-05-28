using Microsoft.AspNetCore.Mvc;
using OutlookShredder.Proxy.Models;
using OutlookShredder.Proxy.Services;

namespace OutlookShredder.Proxy.Controllers;

[ApiController]
[Route("api/archive")]
public class ArchiveController : ControllerBase
{
    private readonly ArchiveCacheService _archive;

    public ArchiveController(ArchiveCacheService archive) => _archive = archive;

    /// <summary>
    /// POST /api/archive/search
    /// At least one content filter (rfqId, customerName, requester, supplierName,
    /// product, hskNumber, notesContains) must be non-empty. Date range alone is rejected.
    /// </summary>
    [HttpPost("search")]
    public IActionResult Search([FromBody] ArchiveSearchRequest req)
    {
        var hasContent = !string.IsNullOrWhiteSpace(req.RfqId)
                      || !string.IsNullOrWhiteSpace(req.CustomerName)
                      || !string.IsNullOrWhiteSpace(req.Requester)
                      || !string.IsNullOrWhiteSpace(req.SupplierName)
                      || !string.IsNullOrWhiteSpace(req.Product)
                      || !string.IsNullOrWhiteSpace(req.HskNumber)
                      || !string.IsNullOrWhiteSpace(req.NotesContains);

        if (!hasContent)
            return BadRequest(new { error = "At least one content filter is required (date range alone is not enough)." });

        var result = _archive.Search(req);
        return Ok(result);
    }

    /// <summary>GET /api/archive/status — quick cold/warm check without full cache/status overhead.</summary>
    [HttpGet("status")]
    public IActionResult Status() => Ok(new
    {
        refCount      = _archive.Refs.Count,
        isLoading     = _archive.IsLoading,
        cacheBuiltUtc = _archive.CacheBuiltUtc,
        lastDeltaUtc  = _archive.LastDeltaUtc,
    });
}
