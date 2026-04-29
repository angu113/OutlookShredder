using System.Collections.Concurrent;
using System.Text.Json;
using OutlookShredder.Proxy.Models;

namespace OutlookShredder.Proxy.Services;

/// <summary>
/// Watches the configured downloads folder for new PDF files and processes them through
/// ERP document detection and SharePoint ingestion.
///
/// Two modes:
///   Real-time: FileSystemWatcher fires for each new or renamed-to-.pdf file.
///   Batch:     ScanFolderAsync processes all PDFs in a folder (called on startup and via API).
///
/// Processed files are tracked by a key (name|size|lastWriteTicks) persisted to
/// erp-processed.json so re-scans across restarts skip already-handled files.
/// AI is only called when the key is not in the processed set.
///
/// Config keys:
///   FileWatcher:Enabled    — "false" disables the service (default enabled)
///   FileWatcher:WatchPath  — folder to watch (default: %USERPROFILE%\Downloads)
/// </summary>
public class FileWatcherService : BackgroundService
{
    private readonly IConfiguration _config;
    private readonly ErpAiService _ai;
    private readonly SharePointService _sp;
    private readonly RfqNotificationService _notify;
    private readonly ILogger<FileWatcherService> _log;

    // Files waiting to be processed (real-time FSW events queue to this)
    private readonly ConcurrentQueue<string> _pending = new();

    // Paths currently in the queue or being processed — prevents double-queuing
    private readonly HashSet<string> _inQueue = new(StringComparer.OrdinalIgnoreCase);
    private readonly object _inQueueLock = new();

    // Health state — set during ExecuteAsync, read by HealthController
    private bool    _enabled;
    private string? _watchPath;
    private bool    _fswActive;
    private bool    _watchPathExists;

    // Persistent processed-file tracking: key = "{name}|{size}|{lastWriteTicks}"
    private readonly HashSet<string> _processedKeys = new(StringComparer.Ordinal);
    private readonly object _processedLock = new();
    private string? _processedFilePath;

    public FileWatcherService(
        IConfiguration config,
        ErpAiService ai,
        SharePointService sp,
        RfqNotificationService notify,
        ILogger<FileWatcherService> log)
    {
        _config  = config;
        _ai      = ai;
        _sp      = sp;
        _notify  = notify;
        _log     = log;
    }

    protected override async Task ExecuteAsync(CancellationToken ct)
    {
        // PDF attachment cleanup runs regardless of whether file watching is enabled
        _ = RunPdfCleanupLoopAsync(ct);

        if ("false".Equals(_config["FileWatcher:Enabled"], StringComparison.OrdinalIgnoreCase))
        {
            _log.LogInformation("[FW] File watcher disabled via FileWatcher:Enabled=false");
            _enabled = false;
            return;
        }

        _enabled = true;

        var cfgPath   = _config["FileWatcher:WatchPath"];
        var watchPath = !string.IsNullOrWhiteSpace(cfgPath)
            ? cfgPath
            : Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.UserProfile), "Downloads");

        _watchPath = watchPath;

        _processedFilePath = Path.Combine(AppContext.BaseDirectory, "erp-processed.json");
        LoadProcessedLog();

        _watchPathExists = Directory.Exists(watchPath);
        if (!_watchPathExists)
        {
            _log.LogWarning("[FW] Watch path does not exist: {Path} — file watcher inactive", watchPath);
            return;
        }

        _log.LogInformation("[FW] Watching {Path} for ERP PDFs", watchPath);

        // Batch scan on startup — only files from the last 30 days to avoid re-sending
        // the entire downloads folder through Claude every time the proxy restarts.
        await ScanFolderAsync(watchPath, ct, maxAgeDays: 1);

        // Real-time FileSystemWatcher
        _fswActive = true;
        using var fsw = new FileSystemWatcher(watchPath, "*.pdf")
        {
            NotifyFilter        = NotifyFilters.FileName | NotifyFilters.LastWrite,
            EnableRaisingEvents = true
        };
        fsw.Created += (_, e) => EnqueueFile(e.FullPath);
        fsw.Renamed += (_, e) =>
        {
            if (e.FullPath.EndsWith(".pdf", StringComparison.OrdinalIgnoreCase))
                EnqueueFile(e.FullPath);
        };

        // Drain queue loop
        while (!ct.IsCancellationRequested)
        {
            while (_pending.TryDequeue(out var path))
            {
                lock (_inQueueLock) _inQueue.Remove(path);
                try { await ProcessFileAsync(path, ct); }
                catch (Exception ex) { _log.LogError(ex, "[FW] Unhandled error processing {File}", path); }
            }
            try { await Task.Delay(1_000, ct); }
            catch (OperationCanceledException) { break; }
        }
    }

    // ── Public API ───────────────────────────────────────────────────────────

    public FileWatcherHealthStatus GetHealthStatus()
    {
        int processed;
        lock (_processedLock) processed = _processedKeys.Count;
        return new FileWatcherHealthStatus(_enabled, _watchPath, _watchPathExists, _fswActive, processed);
    }

    public void ClearProcessedCache()
    {
        lock (_processedLock)
        {
            _processedKeys.Clear();
            SaveProcessedLog();
        }
    }

    // ── Public batch-scan API ────────────────────────────────────────────────

    public async Task<ErpScanResult> ScanFolderAsync(
        string folder, CancellationToken ct = default, int? maxAgeDays = null)
    {
        var result = new ErpScanResult { Folder = folder };

        if (!Directory.Exists(folder))
        {
            _log.LogWarning("[FW] Scan folder not found: {Folder}", folder);
            return result;
        }

        var files = Directory.GetFiles(folder, "*.pdf", SearchOption.TopDirectoryOnly);
        result.FilesFound = files.Length;

        var cutoff = maxAgeDays.HasValue
            ? DateTime.UtcNow.AddDays(-maxAgeDays.Value)
            : (DateTime?)null;

        _log.LogInformation("[FW] Batch scan: {Count} PDF(s) in {Folder}{Age}",
            files.Length, folder, cutoff.HasValue ? $" (last {maxAgeDays}d only)" : "");

        foreach (var file in files)
        {
            if (ct.IsCancellationRequested) break;

            // Age filter — skip files older than cutoff on startup scans
            if (cutoff.HasValue)
            {
                try
                {
                    if (File.GetCreationTimeUtc(file) < cutoff.Value) continue;
                }
                catch { /* ignore stat errors */ }
            }

            var key = GetFileKey(file);
            if (key is not null)
            {
                lock (_processedLock)
                {
                    if (_processedKeys.Contains(key))
                    {
                        result.AlreadyProcessed++;
                        continue;
                    }
                }
            }

            try
            {
                var wasErp = await ProcessFileAsync(file, ct);
                if (wasErp)
                {
                    result.ErpDocuments++;
                    result.ProcessedFiles.Add(Path.GetFileName(file));
                }
                else
                {
                    result.NonErpFiles++;
                }
            }
            catch (Exception ex)
            {
                _log.LogError(ex, "[FW] Error processing {File}", file);
                result.Errors++;
                result.ErrorFiles.Add(Path.GetFileName(file));
            }
        }

        _log.LogInformation("[FW] Scan complete — erp={Erp} nonErp={NonErp} skipped={Skip} errors={Err}",
            result.ErpDocuments, result.NonErpFiles, result.AlreadyProcessed, result.Errors);
        return result;
    }

    // ── Internal ─────────────────────────────────────────────────────────────

    private void EnqueueFile(string path)
    {
        lock (_inQueueLock)
        {
            if (!_inQueue.Add(path)) return;
        }
        _pending.Enqueue(path);
        _log.LogInformation("[FW] New PDF detected: {File}", Path.GetFileName(path));
    }

    /// <summary>
    /// Processes one PDF file through the ERP pipeline.
    /// Returns true if an ERP document was written to SharePoint; false if skipped or not ERP.
    /// </summary>
    private async Task<bool> ProcessFileAsync(string path, CancellationToken ct)
    {
        var fileName = Path.GetFileName(path);

        // Read bytes — retry up to 6 times to handle files still being written
        byte[]? bytes = null;
        for (int attempt = 0; attempt < 6 && bytes is null; attempt++)
        {
            if (attempt > 0) await Task.Delay(500, ct);
            try
            {
                await using var fs = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.Read);
                bytes = new byte[fs.Length];
                var read = await fs.ReadAsync(bytes, ct);
                if (read < bytes.Length) Array.Resize(ref bytes, read);
            }
            catch (IOException) when (attempt < 5)
            {
                _log.LogDebug("[FW] {File} locked — retry {A}", fileName, attempt + 1);
                bytes = null;
            }
        }

        if (bytes is null || bytes.Length == 0)
        {
            _log.LogWarning("[FW] Could not read {File} after retries — skipping", fileName);
            return false;
        }

        // Check processed key AFTER reading (size/time is stable now)
        var key = GetFileKey(path);
        if (key is not null)
        {
            lock (_processedLock)
            {
                if (_processedKeys.Contains(key))
                {
                    _log.LogDebug("[FW] {File} already processed — skipping", fileName);
                    return false;
                }
            }
        }

        // Gate on filename pattern — avoids calling AI on non-ERP files entirely.
        var erpInfo = ErpFilenameParser.Parse(fileName);
        if (erpInfo is null)
        {
            _log.LogDebug("[FW] {File} — filename does not match ERP pattern, skipping", fileName);
            if (key is not null) MarkProcessed(key);
            return false;
        }

        _log.LogInformation("[FW] ERP filename matched: {Type} {DocNum} in {File}",
            erpInfo.DocumentType, erpInfo.DocumentNumber ?? "(no record id)", fileName);

        var base64 = Convert.ToBase64String(bytes);

        // Call AI to extract detail fields only (CustomerName, TotalAmount, DocumentDate, LineItems).
        // The filename already provides DocumentType and DocumentNumber — AI output for those is ignored.
        var extraction = await _ai.ExtractAsync(base64, fileName, ct);

        // Mark processed after the AI call so re-scans don't repeat the work.
        if (key is not null) MarkProcessed(key);

        if (extraction is null)
        {
            _log.LogWarning("[FW] {File} — AI returned null; recording with filename data only", fileName);
            extraction = new OutlookShredder.Proxy.Models.ErpExtraction { IsErpDocument = true };
        }

        // Always override with filename-derived identity (more reliable than AI for these fields).
        extraction.IsErpDocument = true;
        extraction.DocumentType  = erpInfo.DocumentType;
        if (erpInfo.HasDocNumber)
        {
            extraction.DocumentNumber = erpInfo.DocumentNumber;
        }
        else if (!string.IsNullOrEmpty(extraction.DocumentNumber) &&
                 !extraction.DocumentNumber.StartsWith("HSK-", StringComparison.OrdinalIgnoreCase) &&
                 !extraction.DocumentNumber.StartsWith("020803-", StringComparison.OrdinalIgnoreCase))
        {
            // PurchaseOrder without filename record info: AI gave a number that looks like a
            // customer PO, not ours — discard it and let the SP title fall back to the filename.
            _log.LogWarning("[FW] {File} — AI returned DocumentNumber '{Num}' without HSK-/020803- prefix; discarding",
                fileName, extraction.DocumentNumber);
            extraction.DocumentNumber = null;
        }

        _log.LogInformation("[FW] ERP document: {Type} {Number} in {File}",
            extraction.DocumentType, extraction.DocumentNumber, fileName);

        // Upload PDF
        string? pdfUrl = null;
        try
        {
            pdfUrl = await _sp.UploadErpPdfAsync(
                extraction.DocumentNumber ?? Path.GetFileNameWithoutExtension(fileName),
                fileName, bytes, ct);
        }
        catch (Exception ex)
        {
            _log.LogWarning(ex, "[FW] PDF upload failed for {File} — continuing without URL", fileName);
        }

        // Write SP record
        string? spItemId;
        try
        {
            spItemId = await _sp.WriteErpDocumentAsync(
                extraction, fileName, DateTimeOffset.UtcNow, Environment.MachineName, Environment.UserName, pdfUrl, ct);
        }
        catch (Exception ex)
        {
            _log.LogError(ex, "[FW] SP write failed for {File}", fileName);
            return false;
        }

        // Notify all Shredder clients via Service Bus
        _notify.NotifyErpDocument(new OutlookShredder.Proxy.Models.ErpBusRecord
        {
            SpItemId          = spItemId,
            DocumentNumber    = extraction.DocumentNumber,
            DocumentType      = extraction.DocumentType,
            DocumentDate      = extraction.DocumentDate,
            CustomerName      = extraction.CustomerName,
            CustomerReference = extraction.CustomerReference,
            TotalAmount       = extraction.TotalAmount,
            Currency          = extraction.Currency ?? "USD",
            FileName          = fileName,
            PdfUrl            = pdfUrl,
            ReceivedAt        = DateTimeOffset.UtcNow.ToString("o"),
            IsArchived        = false,
            IsNew             = true,
            SourceMachine     = Environment.MachineName,
            SourceUser        = Environment.UserName,
        });

        // Archive older duplicates for the same document number
        if (spItemId is not null && !string.IsNullOrEmpty(extraction.DocumentNumber))
        {
            try
            {
                await _sp.ArchiveOlderErpDocumentsAsync(extraction.DocumentNumber, spItemId, ct);
            }
            catch (Exception ex)
            {
                _log.LogWarning(ex, "[FW] Archive-older step failed for {Number}", extraction.DocumentNumber);
            }
        }

        _log.LogInformation("[FW] Recorded {Type} {Number} ({File}) → SP {Id}",
            extraction.DocumentType, extraction.DocumentNumber, fileName, spItemId);
        return true;
    }

    // ── Processed-file tracking ───────────────────────────────────────────────

    private static string? GetFileKey(string path)
    {
        try
        {
            var info = new FileInfo(path);
            return $"{info.Name}|{info.Length}|{info.LastWriteTimeUtc.Ticks}";
        }
        catch { return null; }
    }

    private void MarkProcessed(string key)
    {
        lock (_processedLock)
        {
            _processedKeys.Add(key);
            SaveProcessedLog();
        }
    }

    private void LoadProcessedLog()
    {
        if (_processedFilePath is null || !File.Exists(_processedFilePath)) return;
        try
        {
            using var doc = JsonDocument.Parse(File.ReadAllText(_processedFilePath));
            if (doc.RootElement.TryGetProperty("keys", out var arr))
                foreach (var el in arr.EnumerateArray())
                    if (el.GetString() is string k) _processedKeys.Add(k);
            _log.LogInformation("[FW] Loaded {Count} processed file keys from erp-processed.json",
                _processedKeys.Count);
        }
        catch (Exception ex)
        {
            _log.LogWarning(ex, "[FW] Could not load erp-processed.json — starting fresh");
        }
    }

    private void SaveProcessedLog()
    {
        if (_processedFilePath is null) return;
        try
        {
            File.WriteAllText(_processedFilePath,
                JsonSerializer.Serialize(new { keys = _processedKeys.ToArray() }));
        }
        catch (Exception ex)
        {
            _log.LogWarning(ex, "[FW] Could not save erp-processed.json");
        }
    }

    // ── PDF cleanup ───────────────────────────────────────────────────────────

    private async Task RunPdfCleanupLoopAsync(CancellationToken ct)
    {
        // Wait a few minutes after startup before the first pass
        try { await Task.Delay(TimeSpan.FromMinutes(3), ct); }
        catch (OperationCanceledException) { return; }

        while (!ct.IsCancellationRequested)
        {
            try
            {
                var count = await _sp.RemoveOldErpPdfsAsync(TimeSpan.FromHours(48), ct);
                if (count > 0)
                    _log.LogInformation("[FW] PDF cleanup: cleared attachments from {Count} old ERP document(s)", count);
            }
            catch (OperationCanceledException) { return; }
            catch (Exception ex)
            {
                _log.LogWarning(ex, "[FW] PDF cleanup error");
            }

            try { await Task.Delay(TimeSpan.FromHours(1), ct); }
            catch (OperationCanceledException) { return; }
        }
    }
}

public record FileWatcherHealthStatus(
    bool    Enabled,
    string? WatchPath,
    bool    WatchPathExists,
    bool    FswActive,
    int     ProcessedCount);
