using Microsoft.AspNetCore.Mvc;
using OutlookShredder.Proxy.Services;

namespace OutlookShredder.Proxy.Controllers;

[ApiController]
public class ErpController : ControllerBase
{
    private readonly FileWatcherService _fw;
    private readonly SharePointService _sp;
    private readonly IConfiguration _config;
    private readonly ILogger<ErpController> _log;

    public ErpController(
        FileWatcherService fw,
        SharePointService sp,
        IConfiguration config,
        ILogger<ErpController> log)
    {
        _fw     = fw;
        _sp     = sp;
        _config = config;
        _log    = log;
    }

    /// <summary>
    /// Scans a folder for PDFs and runs them through ERP detection.
    /// Defaults to the configured FileWatcher:WatchPath (or Downloads if unset).
    /// </summary>
    [HttpPost("/api/erp/scan")]
    public async Task<IActionResult> Scan([FromQuery] string? folder, CancellationToken ct)
    {
        var path = folder
            ?? _config["FileWatcher:WatchPath"]
            ?? Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.UserProfile), "Downloads");

        _log.LogInformation("[ERP] Manual scan triggered for {Path}", path);
        var result = await _fw.ScanFolderAsync(path, ct);
        return Ok(result);
    }

    /// <summary>
    /// Returns recent ERP document records from SharePoint.
    /// </summary>
    [HttpGet("/api/erp/documents")]
    public async Task<IActionResult> GetDocuments(
        [FromQuery] int top = 50,
        [FromQuery] bool includeArchived = false,
        CancellationToken ct = default)
    {
        var docs = await _sp.ReadErpDocumentsAsync(top, includeArchived, ct);
        return Ok(docs);
    }

    /// <summary>
    /// Idempotent: ensures the ErpDocuments SharePoint list and its columns exist.
    /// </summary>
    [HttpPost("/api/erp/setup")]
    public async Task<IActionResult> Setup(CancellationToken ct)
    {
        await _sp.EnsureErpDocumentsListAsync(ct);
        return Ok(new { success = true, message = "ErpDocuments list ensured" });
    }
}
