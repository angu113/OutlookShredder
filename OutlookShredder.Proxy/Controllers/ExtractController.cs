using Microsoft.AspNetCore.Mvc;
using OutlookShredder.Proxy.Models;
using OutlookShredder.Proxy.Services;

namespace OutlookShredder.Proxy.Controllers;

[ApiController]
[Route("api")]
public class ExtractController : ControllerBase
{
    private readonly ClaudeService      _claude;
    private readonly SharePointService  _sp;
    private readonly ILogger<ExtractController> _log;

    public ExtractController(
        ClaudeService     claude,
        SharePointService sp,
        ILogger<ExtractController> log)
    {
        _claude = claude;
        _sp     = sp;
        _log    = log;
    }

    // ── POST /api/extract ────────────────────────────────────────────────────
    /// <summary>
    /// Called by the Office.js add-in task pane.
    /// Extracts RFQ data from email body or attachment, then writes
    /// one SharePoint list item per product line found.
    /// </summary>
    [HttpPost("extract")]
    public async Task<ActionResult<ExtractResponse>> Extract([FromBody] ExtractRequest req)
    {
        if (string.IsNullOrWhiteSpace(req.Content) && string.IsNullOrWhiteSpace(req.Base64Data))
            return BadRequest(new ExtractResponse { Success = false, Error = "Content or Base64Data is required." });

        RfqExtraction? extraction;
        try
        {
            extraction = await _claude.ExtractAsync(req);
        }
        catch (Exception ex)
        {
            _log.LogError(ex, "Claude extraction failed");
            return StatusCode(502, new ExtractResponse { Success = false, Error = ex.Message });
        }

        if (extraction is null || extraction.Products.Count == 0)
        {
            return Ok(new ExtractResponse
            {
                Success   = false,
                Extracted = extraction,
                Error     = "No products array returned by extraction."
            });
        }

        // Write one SharePoint row per product
        var source     = req.SourceType == "attachment" ? "attachment" : "body";
        var sourceFile = req.SourceType == "attachment" ? req.FileName : null;
        var rows       = new List<SpWriteResult>();

        for (int i = 0; i < extraction.Products.Count; i++)
        {
            var row = await _sp.WriteProductRowAsync(extraction, extraction.Products[i], req, source, sourceFile, i);
            rows.Add(row);
        }

        return Ok(new ExtractResponse
        {
            Success   = rows.Any(r => r.Success),
            Extracted = extraction,
            Rows      = rows
        });
    }

    // ── POST /api/setup-columns ──────────────────────────────────────────────
    /// <summary>
    /// Provisions all required columns on the RFQLineItems SharePoint list.
    /// Run once after creating the blank list. Safe to re-run.
    /// </summary>
    [HttpPost("setup-columns")]
    public async Task<IActionResult> SetupColumns()
    {
        try
        {
            var results = await _sp.EnsureColumnsAsync();
            return Ok(new { success = true, columns = results });
        }
        catch (Exception ex)
        {
            return StatusCode(500, new { success = false, error = ex.Message });
        }
    }

    // ── GET /api/health ──────────────────────────────────────────────────────
    [HttpGet("health")]
    public IActionResult Health() =>
        Ok(new { status = "ok", utc = DateTime.UtcNow });
}
