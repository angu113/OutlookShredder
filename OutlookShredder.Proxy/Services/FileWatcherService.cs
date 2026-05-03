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
    private readonly ILogger<FileWatcherService> _log;

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

        // Drain channel — wakes immediately when EnqueueFile writes, no polling delay
        await foreach (var path in _fileChannel.Reader.ReadAllAsync(ct))
        {
            lock (_inQueueLock) _inQueue.Remove(path);
            try { await ProcessFileAsync(path, ct); }
            catch (Exception ex) { _log.LogError(ex, "[FW] Unhandled error processing {File}", path); }
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

        // Payment and Quotation files are not needed — mark processed and skip
        if (erpInfo.DocumentType is "Payment" or "Quotation")
        {
            _log.LogDebug("[FW] {File} — {Type} ignored; marked as processed", fileName, erpInfo.DocumentType);
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

        // Stamp picking slips with the customer name at top-left, then enrich with bold comments + callout page
        if (extraction.DocumentType == "PickingSlip" && !string.IsNullOrWhiteSpace(extraction.CustomerName))
        {
            try { bytes = StampCustomerName(bytes, extraction.CustomerName); }
            catch (Exception ex) { _log.LogWarning(ex, "[FW] PDF stamp failed for {File} — uploading original", fileName); }

            try { bytes = PickingSlipEnricher.Enrich(bytes, _log); }
            catch (Exception ex) { _log.LogWarning(ex, "[FW] Picking-slip enrichment failed for {File} — uploading without callouts", fileName); }
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
            return false;
        }

        // Notify immediately — local file path lets this machine render from disk without waiting
        // for the SharePoint upload. Other machines that reload later get the SP URL from the
        // follow-up notification sent by the background task below.
        var receivedAt = DateTimeOffset.UtcNow.ToString("o");
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
            PdfUrl            = path,   // local file path — same machine renders instantly
            ReceivedAt        = receivedAt,
            IsArchived        = false,
            IsNew             = true,
            SourceMachine     = Environment.MachineName,
            SourceUser        = Environment.UserName,
        });

        _log.LogInformation("[FW] Recorded {Type} {Number} ({File}) → SP {Id} (upload pending)",
            extraction.DocumentType, extraction.DocumentNumber, fileName, spItemId);

        // Background: upload PDF, patch SP URL, re-notify with SP URL, archive older duplicates
        var bgBytes  = bytes;
        var bgDocNum = extraction.DocumentNumber;
        var bgSid    = spItemId;
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

        return true;
    }

    // ── PDF stamping ──────────────────────────────────────────────────────────

    // PdfSharp 6.x requires an explicit font resolver on Windows — set once.
    private static readonly object _fontResolverLock = new();
    private static bool _fontResolverSet;

    private static void EnsureFontResolver()
    {
        if (_fontResolverSet) return;
        lock (_fontResolverLock)
        {
            if (_fontResolverSet) return;
            PdfSharp.Fonts.GlobalFontSettings.FontResolver = new ArialFontResolver();
            _fontResolverSet = true;
        }
    }

    // Reads Arial variants directly from %WINDIR%\Fonts — works in service/task context.
    private sealed class ArialFontResolver : PdfSharp.Fonts.IFontResolver
    {
        private static readonly string FontsDir =
            Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Windows), "Fonts");

        public PdfSharp.Fonts.FontResolverInfo? ResolveTypeface(string familyName, bool isBold, bool isItalic)
        {
            if (!familyName.Equals("Arial", StringComparison.OrdinalIgnoreCase)) return null;
            var face = (isBold, isItalic) switch
            {
                (true,  true)  => "arialbi",
                (true,  false) => "arialbd",
                (false, true)  => "ariali",
                _              => "arial",
            };
            return new PdfSharp.Fonts.FontResolverInfo(face);
        }

        public byte[]? GetFont(string faceName) =>
            new[] { $"{faceName}.ttf", $"{faceName}.ttc", $"{faceName}.otf" }
                .Select(f => Path.Combine(FontsDir, f))
                .Where(File.Exists)
                .Select(File.ReadAllBytes)
                .FirstOrDefault();
    }

    /// <summary>
    /// Draws a bold customer-name label (2.5" wide × 1" tall) at the top-left of page 1.
    /// Text wraps across lines and the font size shrinks until everything fits.
    /// Returns the modified PDF bytes, or throws so the caller can fall back to the original.
    /// </summary>
    private static byte[] StampCustomerName(byte[] pdfBytes, string customerName)
    {
        EnsureFontResolver();

        using var inputStream = new System.IO.MemoryStream(pdfBytes);
        var doc  = PdfSharp.Pdf.IO.PdfReader.Open(inputStream, PdfSharp.Pdf.IO.PdfDocumentOpenMode.Modify);
        var page = doc.Pages[0];

        // XGraphics.FromPdfPage: top-left origin, Y increases downward.
        const double boxX      = 18;   // 0.25" from left
        const double boxY      = 18;   // 0.25" from top
        const double boxWidth  = 180;  // 2.5"
        const double boxHeight = 72;   // 1"
        const double pad       = 6;    // inner padding
        double textAreaW = boxWidth  - pad * 2;
        double textAreaH = boxHeight - pad * 2;

        using var gfx = PdfSharp.Drawing.XGraphics.FromPdfPage(page);

        // Find the largest font (max 16pt, min 7pt) at which all wrapped lines fit the box.
        var (lines, font) = FitTextInBox(gfx, customerName, textAreaW, textAreaH);

        var boxRect = new PdfSharp.Drawing.XRect(boxX, boxY, boxWidth, boxHeight);
        gfx.DrawRectangle(PdfSharp.Drawing.XBrushes.White, boxRect);
        gfx.DrawRectangle(new PdfSharp.Drawing.XPen(PdfSharp.Drawing.XColors.Black, 0.75), boxRect);

        // Centre the line block vertically, each line centred horizontally.
        double lineH   = font.GetHeight();
        double blockH  = lines.Count * lineH;
        double startY  = boxY + pad + (textAreaH - blockH) / 2.0;

        foreach (var line in lines)
        {
            var lineRect = new PdfSharp.Drawing.XRect(boxX + pad, startY, textAreaW, lineH);
            gfx.DrawString(line, font, PdfSharp.Drawing.XBrushes.Black,
                lineRect, PdfSharp.Drawing.XStringFormats.TopCenter);
            startY += lineH;
        }

        using var outputStream = new System.IO.MemoryStream();
        doc.Save(outputStream);
        return outputStream.ToArray();
    }

    private static (List<string> Lines, PdfSharp.Drawing.XFont Font) FitTextInBox(
        PdfSharp.Drawing.XGraphics gfx, string text, double maxW, double maxH)
    {
        for (double size = 16; size >= 7; size -= 1)
        {
            var font  = new PdfSharp.Drawing.XFont("Arial", size, PdfSharp.Drawing.XFontStyleEx.Bold);
            var lines = WrapWords(gfx, font, text, maxW);
            if (lines.Count * font.GetHeight() <= maxH)
                return (lines, font);
        }
        var fallback = new PdfSharp.Drawing.XFont("Arial", 7, PdfSharp.Drawing.XFontStyleEx.Bold);
        return (WrapWords(gfx, fallback, text, maxW), fallback);
    }

    private static List<string> WrapWords(
        PdfSharp.Drawing.XGraphics gfx, PdfSharp.Drawing.XFont font, string text, double maxW)
    {
        var lines   = new List<string>();
        var current = "";
        foreach (var word in text.Split(' ', StringSplitOptions.RemoveEmptyEntries))
        {
            var candidate = string.IsNullOrEmpty(current) ? word : current + " " + word;
            if (gfx.MeasureString(candidate, font).Width <= maxW)
            {
                current = candidate;
            }
            else
            {
                if (!string.IsNullOrEmpty(current)) lines.Add(current);
                current = word;
            }
        }
        if (!string.IsNullOrEmpty(current)) lines.Add(current);
        return lines.Count > 0 ? lines : [text];
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
