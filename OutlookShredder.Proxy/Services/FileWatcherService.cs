using System.Collections.Concurrent;
using System.Text.Json;
using System.Threading.Channels;
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
    private readonly WorkflowCardService _workflow;
    private readonly ILogger<FileWatcherService> _log;

    // Configured shop-operation keywords scanned in PickingSlip B: comment lines.
    // Drives auto-add into the Trigger Prioritize column (WorkflowCardService.AutoCreateFromPickingSlipAsync).
    private readonly IReadOnlyList<string> _processingKeywords;

    // ── ErpFooter: stamps a configurable T&C footer on Sales Orders / Quotations ──
    private readonly bool _footerEnabled;
    private readonly HashSet<string> _footerDocTypes;
    private readonly ErpDocumentFooterService.FooterOptions? _footerOptions;

    // Files waiting to be processed — Channel gives immediate wake-up vs. poll loop
    private readonly Channel<string> _fileChannel = Channel.CreateUnbounded<string>(
        new UnboundedChannelOptions { SingleWriter = false, SingleReader = true });

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
        WorkflowCardService workflow,
        ILogger<FileWatcherService> log)
    {
        _config   = config;
        _ai       = ai;
        _sp       = sp;
        _notify   = notify;
        _workflow = workflow;
        _log      = log;

        _processingKeywords = config.GetSection("Workflow:ProcessingKeywords").Get<string[]>()
                              ?? ["Laser Cutting", "Bending", "Welding", "Drilling", "Fabricating"];

        // ErpFooter — build the footer options once from config. Disabled (null options) when
        // the section is off or no text is configured, so stamping is skipped entirely.
        _footerEnabled  = config.GetValue("ErpFooter:Enabled", false);
        _footerDocTypes = new(
            config.GetSection("ErpFooter:DocumentTypes").Get<string[]>() ?? ["SalesOrder", "Quotation"],
            StringComparer.OrdinalIgnoreCase);
        var footerText = config["ErpFooter:Text"];
        _footerOptions = _footerEnabled && !string.IsNullOrWhiteSpace(footerText)
            ? new ErpDocumentFooterService.FooterOptions
            {
                Text           = footerText,
                FontSizePt     = config.GetValue("ErpFooter:FontSizePt",     7.0),
                SideMarginPt   = config.GetValue("ErpFooter:SideMarginPt",   36.0),
                BottomMarginPt = config.GetValue("ErpFooter:BottomMarginPt", 14.0),
                BoxHeightPt    = config.GetValue("ErpFooter:BoxHeightPt",    30.0),
                EveryPage      = config.GetValue("ErpFooter:EveryPage",      true),
                TopRule        = config.GetValue("ErpFooter:TopRule",        true),
                Center         = config.GetValue("ErpFooter:Center",         true),
            }
            : null;
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

        // Store in %APPDATA%\Shredder\ so reinstalls don't wipe the processed-file list
        var dataDir = Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "Shredder");
        Directory.CreateDirectory(dataDir);
        _processedFilePath = Path.Combine(dataDir, "erp-processed.json");
        LoadProcessedLog();

        _watchPathExists = Directory.Exists(watchPath);
        if (!_watchPathExists)
        {
            _log.LogWarning("[FW] Watch path does not exist: {Path} — file watcher inactive", watchPath);
            return;
        }

        _log.LogInformation("[FW] Watching {Path} for ERP PDFs", watchPath);

        // On startup, mark every existing PDF as already processed so they are silently
        // ignored. Only PDFs that arrive after this point will be sent through the AI pipeline.
        SeedExistingFilesAsProcessed(watchPath);
        ArchiveOldCaptures(watchPath);   // tidy export cruft >5 business days old into an Archive folder

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

        // Steve CSV watcher — detects ExportedData*.csv (OB grid export) dropped by the extension.
        using var csvFsw = new FileSystemWatcher(watchPath, "*.csv")
        {
            NotifyFilter        = NotifyFilters.FileName | NotifyFilters.LastWrite,
            EnableRaisingEvents = true
        };
        csvFsw.Created += (_, e) => OnExportCsvDetected(e.FullPath);
        csvFsw.Renamed += (_, e) =>
        {
            if (e.FullPath.EndsWith(".csv", StringComparison.OrdinalIgnoreCase))
                OnExportCsvDetected(e.FullPath);
        };

        // Drain channel — process up to 3 files concurrently so two simultaneous ERP
        // documents don't stall behind each other's AI calls.
        using var sem = new SemaphoreSlim(3);
        try
        {
            await foreach (var path in _fileChannel.Reader.ReadAllAsync(ct))
            {
                lock (_inQueueLock) _inQueue.Remove(path);
                await sem.WaitAsync(ct);
                _ = Task.Run(async () =>
                {
                    try
                    {
                        var wasErp = await ProcessFileAsync(path, ct);
                        if (wasErp)
                            _notify.NotifyRfqProcessed(new RfqProcessedNotification { EventType = "ErpDocument" });
                    }
                    catch (OperationCanceledException) when (ct.IsCancellationRequested) { /* shutdown — drop in-flight work */ }
                    catch (Exception ex) { _log.LogError(ex, "[FW] Unhandled error processing {File}", path); }
                    finally { sem.Release(); }
                }, ct);
            }
        }
        catch (OperationCanceledException) when (ct.IsCancellationRequested)
        {
            // Normal shutdown: the host cancelled ct, so ReadAllAsync/WaitAsync throw.
            // Swallow it so ExecuteAsync returns cleanly — an unhandled OCE here trips
            // BackgroundServiceExceptionBehavior.StopHost and disrupts graceful shutdown,
            // leaving the process orphaned on port 7000 after `schtasks /End`.
            _log.LogInformation("[FW] Shutdown requested — file watcher stopped cleanly");
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

    // ── Steve: invoice CSV detection ──────────────────────────────────────────

    // Captured Steve/GP export filenames: OB grid exports (ExportedData*.csv) and the Heartland
    // "Merchant Batch Download*.csv". The recon consumer classifies by content, so matching either
    // filename here is safe.
    private static readonly System.Text.RegularExpressions.Regex _exportCsvRx =
        new(@"^(ExportedData(\s*\(\d+\))?|Merchant Batch Download.*)\.csv$",
            System.Text.RegularExpressions.RegexOptions.IgnoreCase |
            System.Text.RegularExpressions.RegexOptions.Compiled);

    private void OnExportCsvDetected(string path)
    {
        var name = Path.GetFileName(path);
        if (!_exportCsvRx.IsMatch(name)) return;

        try
        {
            var steveDir = Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
                "Shredder", "steve-exports");
            Directory.CreateDirectory(steveDir);
            var dest = Path.Combine(steveDir,
                $"export-{DateTime.Now:yyyyMMdd-HHmmss}.csv");
            File.Copy(path, dest, overwrite: true);
            SteveState.SetExportResult(dest);
            _log.LogInformation("[Steve] Export detected ({Name}) and copied to {Dest}", name, dest);
            ArchiveOldCaptures(Path.GetDirectoryName(path));
        }
        catch (Exception ex)
        {
            _log.LogWarning(ex, "[Steve] Failed to copy export from {Path}", path);
        }
    }

    // ── Archive: tidy captured export cruft (>N business days) ────────────────

    private const int ArchiveAfterBusinessDays = 5;

    /// <summary>
    /// Moves captured export files older than <see cref="ArchiveAfterBusinessDays"/> business days into a
    /// "Shredder Archive" folder (created if absent) so cruft doesn't accumulate. Covers the export
    /// originals our watcher captures in the watched folder (our filename patterns only — never the
    /// user's unrelated files) and our own copies in steve-exports. The user can review/delete Archive.
    /// </summary>
    private void ArchiveOldCaptures(string? watchPath)
    {
        try
        {
            if (string.IsNullOrEmpty(watchPath) || !Directory.Exists(watchPath)) return;
            var archiveDir = Path.Combine(watchPath, "Shredder Archive");
            int moved = 0;

            // Export originals (our patterns) sitting in the watched folder.
            foreach (var f in Directory.EnumerateFiles(watchPath, "*.csv"))
                if (_exportCsvRx.IsMatch(Path.GetFileName(f)) && IsOlderThanBusinessDays(f, ArchiveAfterBusinessDays))
                    moved += MoveToArchive(f, archiveDir) ? 1 : 0;

            // Our captured copies in steve-exports.
            var steveDir = Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "Shredder", "steve-exports");
            if (Directory.Exists(steveDir))
                foreach (var f in Directory.EnumerateFiles(steveDir, "*.csv"))
                    if (IsOlderThanBusinessDays(f, ArchiveAfterBusinessDays))
                        moved += MoveToArchive(f, archiveDir) ? 1 : 0;

            if (moved > 0)
                _log.LogInformation("[FW] Archived {N} captured export file(s) >{D} business days old to {Dir}",
                    moved, ArchiveAfterBusinessDays, archiveDir);
        }
        catch (Exception ex) { _log.LogWarning(ex, "[FW] Archive sweep failed"); }
    }

    private static bool IsOlderThanBusinessDays(string path, int businessDays)
    {
        try { return BusinessDaysBetween(File.GetLastWriteTime(path), DateTime.Now) > businessDays; }
        catch { return false; }
    }

    private static int BusinessDaysBetween(DateTime from, DateTime to)
    {
        if (to <= from) return 0;
        int days = 0;
        for (var d = from.Date.AddDays(1); d <= to.Date; d = d.AddDays(1))
            if (d.DayOfWeek != DayOfWeek.Saturday && d.DayOfWeek != DayOfWeek.Sunday) days++;
        return days;
    }

    private bool MoveToArchive(string file, string archiveDir)
    {
        try
        {
            Directory.CreateDirectory(archiveDir);
            var dest = Path.Combine(archiveDir, Path.GetFileName(file));
            if (File.Exists(dest))
                dest = Path.Combine(archiveDir,
                    $"{Path.GetFileNameWithoutExtension(file)}-{DateTime.Now:yyyyMMddHHmmss}{Path.GetExtension(file)}");
            File.Move(file, dest);
            return true;
        }
        catch (Exception ex) { _log.LogWarning(ex, "[FW] Could not archive {File}", file); return false; }
    }

    /// <summary>
    /// Removes the N most-recently-modified entries from the processed-file cache so
    /// those files are picked up again on the next scan.  Keys encode lastWriteTicks as
    /// the third "|"-delimited segment; we sort descending to find the newest ones.
    /// Returns the number of keys actually removed.
    /// </summary>
    public int RemoveLastProcessedKeys(int count)
    {
        lock (_processedLock)
        {
            var ordered = _processedKeys
                .Select(k => {
                    var parts = k.Split('|');
                    long ticks = parts.Length >= 3 && long.TryParse(parts[2], out var t) ? t : 0L;
                    return (Key: k, Ticks: ticks);
                })
                .OrderByDescending(x => x.Ticks)
                .Take(count)
                .Select(x => x.Key)
                .ToList();

            foreach (var k in ordered) _processedKeys.Remove(k);
            if (ordered.Count > 0) SaveProcessedLog();
            return ordered.Count;
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
                    _notify.NotifyRfqProcessed(new RfqProcessedNotification { EventType = "ErpDocument" });
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

    /// <summary>
    /// Adds every PDF currently in the folder to the processed cache without sending any
    /// of them through AI. Called once on startup so the watcher only processes new arrivals.
    /// </summary>
    private void SeedExistingFilesAsProcessed(string folder)
    {
        var files = Directory.GetFiles(folder, "*.pdf", SearchOption.TopDirectoryOnly);
        int added = 0;
        lock (_processedLock)
        {
            foreach (var file in files)
            {
                var key = GetFileKey(file);
                if (key is not null && _processedKeys.Add(key))
                    added++;
            }
            if (added > 0) SaveProcessedLog();
        }
        _log.LogInformation("[FW] Seeded {Count} pre-existing PDF(s) as processed — only new arrivals will be extracted", added);
    }

    private void EnqueueFile(string path)
    {
        lock (_inQueueLock)
        {
            if (!_inQueue.Add(path)) return;
        }
        _fileChannel.Writer.TryWrite(path);
        _log.LogInformation("[FW] New PDF detected: {File}", Path.GetFileName(path));
    }

    /// <summary>
    /// Processes one PDF file through the ERP pipeline.
    /// Returns true if an ERP document was written to SharePoint; false if skipped or not ERP.
    /// </summary>
    private async Task<bool> ProcessFileAsync(string path, CancellationToken ct)
    {
        var fileName = Path.GetFileName(path);
        var sw = System.Diagnostics.Stopwatch.StartNew();

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

        _log.LogInformation("[FW] Timing {File}: read={Ms}ms ({Kb}KB)", fileName, sw.ElapsedMilliseconds, bytes.Length / 1024);
        sw.Restart();

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

        // Payment files are not needed — mark processed and skip.
        // (Quotations are now ingested so the T&C footer can be stamped on them — ErpFooter feature.)
        if (erpInfo.DocumentType is "Payment")
        {
            _log.LogDebug("[FW] {File} — {Type} ignored; marked as processed", fileName, erpInfo.DocumentType);
            if (key is not null) MarkProcessed(key);
            return false;
        }

        _log.LogInformation("[FW] ERP filename matched: {Type} {DocNum} in {File}",
            erpInfo.DocumentType, erpInfo.DocumentNumber ?? "(no record id)", fileName);

        var base64 = Convert.ToBase64String(bytes);

        // For PickingSlips: fire PDF enrichment on a background thread concurrently with the AI
        // call. Enrichment only needs raw bytes — PdfPig extracts the ship-to name independently.
        Task<(byte[] Enriched, string? ShipToName, IReadOnlyList<string> ProcessOps)>? enrichTask = null;
        if (erpInfo.DocumentType == "PickingSlip")
            enrichTask = Task.Run(() => PickingSlipEnricher.EnrichPickingSlip(bytes, null, _log, _processingKeywords), ct);

        // Call AI to extract detail fields (CustomerName, TotalAmount, DocumentDate, LineItems).
        // Filename already provides DocumentType and DocumentNumber — AI output for those is ignored.
        var extraction = await _ai.ExtractAsync(base64, fileName, ct);

        _log.LogInformation("[FW] Timing {File}: ai={Ms}ms", fileName, sw.ElapsedMilliseconds);
        sw.Restart();

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
            _log.LogWarning("[FW] {File} — AI returned DocumentNumber '{Num}' without HSK-/020803- prefix; discarding",
                fileName, extraction.DocumentNumber);
            extraction.DocumentNumber = null;
        }

        _log.LogInformation("[FW] ERP document: {Type} {Number} in {File}",
            extraction.DocumentType, extraction.DocumentNumber, fileName);

        // Await enrichment (already running in parallel with the AI call above).
        bool bytesModified = false;
        IReadOnlyList<string> processOps = [];
        if (enrichTask is not null)
        {
            try
            {
                var (enriched, shipToName, ops) = await enrichTask;
                _log.LogInformation("[FW] Timing {File}: enrich={Ms}ms (ran parallel with AI)", fileName, sw.ElapsedMilliseconds);
                sw.Restart();
                if (!string.IsNullOrWhiteSpace(shipToName))
                {
                    _log.LogInformation("[FW] Ship-to name from PDF: '{Name}'", shipToName);
                    extraction.CustomerName = shipToName;
                }
                if (!ReferenceEquals(enriched, bytes))
                {
                    bytes = enriched;
                    bytesModified = true;
                }
                processOps = ops;
            }
            catch (Exception ex)
            {
                _log.LogWarning(ex, "[FW] Picking-slip enrichment failed for {File}", fileName);
            }
        }

        // Sales Orders / Quotations: stamp the configurable T&C footer (white box blocks the
        // existing page-bottom boilerplate). Reuses the modified-bytes → temp-file → SP path below.
        if (_footerOptions is not null && _footerDocTypes.Contains(extraction.DocumentType ?? ""))
        {
            try
            {
                var stamped = ErpDocumentFooterService.StampFooter(bytes, _footerOptions, _log);
                if (!ReferenceEquals(stamped, bytes))
                {
                    bytes = stamped;
                    bytesModified = true;
                    _log.LogInformation("[FW] T&C footer stamped on {Type} {File}",
                        extraction.DocumentType, fileName);
                }
            }
            catch (Exception ex)
            {
                _log.LogWarning(ex, "[FW] T&C footer stamp failed for {File}", fileName);
            }
        }

        // If bytes were modified, write the stamped PDF to a temp file so the immediate notification
        // path points to the stamped version. Temp file is deleted by the background task after upload.
        string notifyPath = path;
        string? tempPath = null;
        if (bytesModified)
        {
            tempPath = Path.Combine(Path.GetTempPath(), $"Shredder_{Guid.NewGuid():N}_{fileName}");
            try
            {
                await File.WriteAllBytesAsync(tempPath, bytes, ct);
                notifyPath = tempPath;
                _log.LogDebug("[FW] Stamped temp file written: {Temp}", tempPath);
            }
            catch (Exception ex)
            {
                _log.LogWarning(ex, "[FW] Could not write stamped temp file for {File} — notifying with original", fileName);
                tempPath = null;
                notifyPath = path;
            }
        }

        // Write SP record immediately (no PDF URL yet — upload happens in background)
        string? spItemId;
        try
        {
            spItemId = await _sp.WriteErpDocumentAsync(
                extraction, fileName, DateTimeOffset.UtcNow, Environment.MachineName, Environment.UserName,
                pdfUrl: null, ct);
        }
        catch (Exception ex)
        {
            _log.LogError(ex, "[FW] SP write failed for {File}", fileName);
            if (tempPath is not null) try { File.Delete(tempPath); } catch { }
            return false;
        }

        _log.LogInformation("[FW] Timing {File}: sp-write={Ms}ms → notifying", fileName, sw.ElapsedMilliseconds);

        // Notify immediately — notifyPath is the stamped temp file (or original path when no stamping).
        // Other machines that reload later get the SP URL from the follow-up notification below.
        var receivedAt    = DateTimeOffset.UtcNow.ToString("o");
        var lineItemsJson = extraction.LineItems.Count > 0
            ? System.Text.Json.JsonSerializer.Serialize(extraction.LineItems)
            : null;
        _notify.NotifyErpDocument(new ErpBusRecord
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
            PdfUrl            = notifyPath,
            ReceivedAt        = receivedAt,
            IsArchived        = false,
            IsNew             = true,
            SourceMachine     = Environment.MachineName,
            SourceUser        = Environment.UserName,
            DeliveryAddress   = extraction.DeliveryAddress,
            DeliveryMethod    = extraction.DeliveryMethod,
            LineItemsJson     = lineItemsJson,
        });

        _log.LogInformation("[FW] Recorded {Type} {Number} ({File}) → SP {Id} (upload pending)",
            extraction.DocumentType, extraction.DocumentNumber, fileName, spItemId);

        // Background: upload PDF, patch SP URL, re-notify with SP URL, archive older duplicates.
        // Also deletes the temp file once the upload succeeds (SP now owns the stamped copy).
        var bgBytes   = bytes;
        var bgDocNum  = extraction.DocumentNumber;
        var bgSid     = spItemId;
        var bgTempPath = tempPath;
        var bgRecord = new ErpBusRecord
        {
            SpItemId          = spItemId,
            DocumentNumber    = bgDocNum,
            DocumentType      = extraction.DocumentType,
            DocumentDate      = extraction.DocumentDate,
            CustomerName      = extraction.CustomerName,
            CustomerReference = extraction.CustomerReference,
            TotalAmount       = extraction.TotalAmount,
            Currency          = extraction.Currency ?? "USD",
            FileName          = fileName,
            ReceivedAt        = receivedAt,
            IsArchived        = false,
            IsNew             = false,  // update, not new — suppress auto-jump in Focus
            SourceMachine     = Environment.MachineName,
            SourceUser        = Environment.UserName,
            DeliveryAddress   = extraction.DeliveryAddress,
            DeliveryMethod    = extraction.DeliveryMethod,
            LineItemsJson     = lineItemsJson,
        };
        _ = Task.Run(async () =>
        {
            try
            {
                var pdfUrl = await _sp.UploadErpPdfAsync(
                    bgDocNum ?? Path.GetFileNameWithoutExtension(fileName), fileName, bgBytes,
                    CancellationToken.None);
                if (pdfUrl is not null && bgSid is not null)
                {
                    await _sp.PatchErpDocumentPdfUrlAsync(bgSid, pdfUrl, CancellationToken.None);
                    bgRecord.PdfUrl = pdfUrl;
                    _notify.NotifyErpDocument(bgRecord);
                    _log.LogInformation("[FW] PDF uploaded and SP patched for {File}", fileName);
                }
            }
            catch (Exception ex) { _log.LogWarning(ex, "[FW] Background PDF upload failed for {File}", fileName); }
            finally
            {
                // SP now has the stamped PDF — delete the local temp file
                if (bgTempPath is not null)
                    try { File.Delete(bgTempPath); } catch { }
            }

            if (bgSid is not null && !string.IsNullOrEmpty(bgDocNum))
            {
                try { await _sp.ArchiveOlderErpDocumentsAsync(bgDocNum, bgSid, CancellationToken.None); }
                catch (Exception ex) { _log.LogWarning(ex, "[FW] Archive-older failed for {Number}", bgDocNum); }
            }
        });

        // For PurchaseOrders, also run the RFQ purchase-matching pipeline
        if (extraction.DocumentType == "PurchaseOrder" && !string.IsNullOrEmpty(extraction.CustomerName))
        {
            try { await TriggerPoMatchingAsync(extraction, fileName, ct); }
            catch (Exception ex) { _log.LogWarning(ex, "[FW] PO matching failed for {File}", fileName); }
        }

        // For PickingSlips: route into the Trigger Prioritize column when keywords match
        // or DeliveryMethod is literally "Delivery". WorkflowCardService handles dedup.
        if (extraction.DocumentType == "PickingSlip")
        {
            try { await _workflow.AutoCreateFromPickingSlipAsync(extraction, spItemId, processOps, ct); }
            catch (Exception ex) { _log.LogWarning(ex, "[FW] Workflow auto-create failed for {File}", fileName); }
        }

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

    // ── ERP PO → RFQ matching ─────────────────────────────────────────────────

    /// <summary>
    /// Called after a PurchaseOrder ERP document is written to SP.
    /// Writes a PurchaseOrders list record, tries to link the PO to the matching
    /// open RFQ by MSPC code or supplier fallback, then publishes a "PO" bus event
    /// so the RFQ tab updates its purchase markers.
    /// </summary>
    private async Task TriggerPoMatchingAsync(
        ErpExtraction extraction, string fileName, CancellationToken ct)
    {
        var supplierName = extraction.CustomerName!;
        var poNumber     = extraction.DocumentNumber;

        // Map ERP line items to PoLineItem; treat product codes containing '/' as MSPC codes
        var poLineItems = extraction.LineItems.Select(li => new PoLineItem
        {
            Product  = li.Description,
            Quantity = double.TryParse(li.Quantity, out var q) ? q : null,
            Mspc     = li.Code?.Contains('/') == true ? li.Code : null,
            Size     = li.Description,
        }).ToList();

        var lineItemsJson = JsonSerializer.Serialize(poLineItems,
            new JsonSerializerOptions { PropertyNamingPolicy = JsonNamingPolicy.CamelCase });

        // Try MSPC-based matching first — resolves rfqId when ERP codes are our MSPCs
        string? rfqId = null;
        var matched = await _sp.MatchAndMarkRliByMspcAsync(supplierName, poNumber, poLineItems);
        if (matched.Count > 0)
            rfqId = matched.First();

        // Fallback: find the most recently active open RFQ for this supplier
        if (rfqId is null)
            rfqId = await _sp.FindOpenRfqForSupplierAsync(supplierName, ct);

        // Write the PurchaseOrders SP record (deduped by PoNumber)
        var poSpItemId = await _sp.WritePurchaseOrderAsync(
            rfqId ?? "UNKNOWN", supplierName, poNumber,
            DateTimeOffset.UtcNow.ToString("o"), messageId: null, lineItemsJson);

        // Mark SLI rows as purchased and check for RFQ completion
        if (poSpItemId is not null && rfqId is not null)
            await _sp.UpdateRliPurchaseStatusAsync(rfqId, supplierName, poSpItemId, poLineItems, poNumber);

        // Publish PO bus event → RFQ tab updates purchase markers on affected rows
        _notify.NotifyRfqProcessed(new RfqProcessedNotification
        {
            EventType    = "PO",
            RfqId        = rfqId,
            SupplierName = supplierName,
            MessageId    = null,
            Products     = poLineItems.Select(li => new RfqNotificationProduct
            {
                Name = li.Product,
                Mspc = li.Mspc,
                Size = li.Size,
            }).ToList(),
        });

        _log.LogInformation("[FW] PO {Num}: RFQ={RfqId} supplier={Supplier} ({Items} items)",
            poNumber, rfqId ?? "unmatched", supplierName, poLineItems.Count);
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
                var count = await _sp.RemoveOldErpPdfsAsync(TimeSpan.FromDays(7), ct);
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
