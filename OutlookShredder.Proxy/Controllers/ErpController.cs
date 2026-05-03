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
    /// Proxies a SharePoint PDF download using app-only credentials.
    /// Clients cannot fetch SharePoint WebUrls directly (no auth); this endpoint adds the Bearer token.
    /// </summary>
    [HttpGet("/api/erp/pdf")]
    public async Task<IActionResult> GetPdf([FromQuery] string url, CancellationToken ct)
    {
        if (string.IsNullOrWhiteSpace(url))
            return BadRequest(new { error = "url query parameter is required" });

        try
        {
            var bytes = await _sp.DownloadSpFileAsync(url, ct);
            return File(bytes, "application/pdf");
        }
        catch (Exception ex)
        {
            _log.LogWarning(ex, "[ERP] PDF proxy download failed for {Url}", url);
            return StatusCode(502, new { error = ex.Message });
        }
    }

    /// <summary>
    /// Returns recent ERP document records from SharePoint.
    /// </summary>
    [HttpGet("/api/erp/documents")]
    public async Task<IActionResult> GetDocuments(
        [FromQuery] int top = 200,
        [FromQuery] bool includeArchived = false,
        [FromQuery] int? daysBack = null,
        CancellationToken ct = default)
    {
        var docs = await _sp.ReadErpDocumentsAsync(top, includeArchived, daysBack, ct);
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

    /// <summary>
    /// Deletes ErpDocuments records matching the given document types (comma-separated).
    /// Example: DELETE /api/erp/clean-by-type?types=Payment,Quotation
    /// Does NOT clear the processed-file cache — ignored files remain ignored.
    /// </summary>
    [HttpDelete("/api/erp/clean-by-type")]
    public async Task<IActionResult> CleanByType([FromQuery] string types, CancellationToken ct)
    {
        var typeList = (types ?? "").Split(',', StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries);
        if (typeList.Length == 0)
            return BadRequest(new { error = "Provide at least one type via ?types=Payment,Quotation" });

        _log.LogInformation("[ERP] CleanByType: deleting types [{Types}]", string.Join(", ", typeList));
        var deleted = await _sp.DeleteErpDocumentsByTypeAsync(typeList, ct);
        return Ok(new { deleted, types = typeList });
    }

    /// <summary>
    /// Retroactively archives all duplicate ErpDocuments records in SharePoint.
    /// For each DocumentNumber that has more than one non-archived record, keeps the most
    /// recently received and marks the rest IsArchived=true.
    /// Safe to call multiple times; idempotent.
    /// </summary>
    [HttpPost("/api/erp/archive-duplicates")]
    public async Task<IActionResult> ArchiveDuplicates(CancellationToken ct)
    {
        var all = await _sp.ReadErpDocumentsAsync(top: 1000, includeArchived: false, ct: ct);

        var dupeGroups = all
            .Where(d => !string.IsNullOrEmpty(d.DocumentNumber))
            .GroupBy(d => d.DocumentNumber!)
            .Where(g => g.Count() > 1)
            .ToList();

        int totalArchived = 0;
        foreach (var group in dupeGroups)
        {
            var winner = group
                .OrderByDescending(d => DateTimeOffset.TryParse(d.ReceivedAt, out var t) ? t : DateTimeOffset.MinValue)
                .First();

            if (winner.SpItemId is null) continue;

            await _sp.ArchiveOlderErpDocumentsAsync(group.Key, winner.SpItemId, ct);
            totalArchived += group.Count() - 1;

            _log.LogInformation("[ERP] archive-duplicates: kept {Id} for {Number}, archived {Count} older record(s)",
                winner.SpItemId, group.Key, group.Count() - 1);
        }

        return Ok(new { duplicateGroups = dupeGroups.Count, archived = totalArchived });
    }

    /// <summary>
    /// Saves user-applied stamp annotations for one ERP document.
    /// Replaces the full annotation list — pass an empty array to clear all stamps.
    /// </summary>
    [HttpPatch("/api/erp/documents/{spItemId}/annotations")]
    public async Task<IActionResult> PatchAnnotations(
        string spItemId,
        [FromBody] List<OutlookShredder.Proxy.Models.ErpAnnotation> annotations,
        CancellationToken ct)
    {
        var json = System.Text.Json.JsonSerializer.Serialize(annotations);
        await _sp.PatchErpAnnotationsAsync(spItemId, json, ct);
        return Ok(new { spItemId, count = annotations.Count });
    }

    /// <summary>
    /// Deletes all records from the ErpDocuments SharePoint list.
    /// Also clears the local processed-file cache so a subsequent scan re-processes everything.
    /// </summary>
    [HttpDelete("/api/erp/clean")]
    public async Task<IActionResult> Clean(CancellationToken ct)
    {
        _fw.ClearProcessedCache();
        var deleted = await _sp.DeleteAllErpDocumentsAsync(ct);
        _log.LogInformation("[ERP] Clean: deleted {Count} SP records and cleared processed-file cache", deleted);
        return Ok(new { deleted });
    }
}
