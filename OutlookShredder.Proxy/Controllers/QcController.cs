using Microsoft.AspNetCore.Mvc;
using OutlookShredder.Proxy.Services;

namespace OutlookShredder.Proxy.Controllers;

[ApiController]
[Route("api/qc")]
public class QcController : ControllerBase
{
    private readonly SharePointService _sp;

    public QcController(SharePointService sp) => _sp = sp;

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
}
