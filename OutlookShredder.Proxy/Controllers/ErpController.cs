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
    /// Pass maxAgeDays to skip files older than N days.
    /// Pass reset=true to clear the processed-file cache before scanning (re-processes everything in window).
    /// </summary>
    [HttpPost("/api/erp/scan")]
    public async Task<IActionResult> Scan(
        [FromQuery] string? folder,
        [FromQuery] int? maxAgeDays,
        [FromQuery] bool reset,
        CancellationToken ct)
    {
        var cfgPath = _config["FileWatcher:WatchPath"];
        var path = !string.IsNullOrWhiteSpace(folder) ? folder
            : !string.IsNullOrWhiteSpace(cfgPath)    ? cfgPath
            : Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.UserProfile), "Downloads");

        if (reset)
        {
            _fw.ClearProcessedCache();
            _log.LogInformation("[ERP] Processed-file cache cleared by manual scan request");
        }

        _log.LogInformation("[ERP] Manual scan triggered for {Path} (maxAgeDays={Days} reset={Reset})", path, maxAgeDays, reset);
        var result = await _fw.ScanFolderAsync(path, ct, maxAgeDays);
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
