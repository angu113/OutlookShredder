using Microsoft.AspNetCore.Mvc;
using OutlookShredder.Proxy.Services;
using OutlookShredder.Proxy.Services.Drawing;

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
    /// Returns page dimensions + fractional bounding box of the Description (Special Instructions)
    /// cell in a picking slip PDF. url may be a SharePoint HTTPS URL or an absolute local file path.
    /// </summary>
    [HttpGet("/api/erp/stamp-bounds")]
    public async Task<IActionResult> GetStampBounds([FromQuery] string url, CancellationToken ct)
    {
        if (string.IsNullOrWhiteSpace(url))
            return BadRequest(new { error = "url parameter is required" });

        try
        {
            var bytes = url.StartsWith("http", StringComparison.OrdinalIgnoreCase)
                ? await _sp.DownloadSpFileAsync(url, ct)
                : await System.IO.File.ReadAllBytesAsync(url, ct);

            var dims   = PickingSlipEnricher.GetPageDimensions(bytes);
            var bounds = PickingSlipEnricher.ExtractDescriptionBoxBounds(bytes);

            var pages = dims.Select((d, i) => new
            {
                pageIndex   = i,
                widthPt     = Math.Round(d.WidthPt,  2),
                heightPt    = Math.Round(d.HeightPt, 2),
                widthIn     = Math.Round(d.WidthPt  / 72.0, 4),
                heightIn    = Math.Round(d.HeightPt / 72.0, 4),
            }).ToList();

            if (bounds is null)
                return Ok(new { pages, descriptionBox = (object?)null });

            var b = bounds.Value;
            var pg = dims.ElementAtOrDefault(b.PageIndex);
            return Ok(new
            {
                pages,
                descriptionBox = new
                {
                    pageIndex   = b.PageIndex,
                    leftFrac    = Math.Round(b.LeftFrac,   4),
                    topFrac     = Math.Round(b.TopFrac,    4),
                    widthFrac   = Math.Round(b.WidthFrac,  4),
                    heightFrac  = Math.Round(b.HeightFrac, 4),
                    // Absolute inches on that page (for direct coordinate comparison)
                    leftIn      = Math.Round(b.LeftFrac   * pg.WidthPt  / 72.0, 4),
                    topIn       = Math.Round(b.TopFrac    * pg.HeightPt / 72.0, 4),
                    widthIn     = Math.Round(b.WidthFrac  * pg.WidthPt  / 72.0, 4),
                    heightIn    = Math.Round(b.HeightFrac * pg.HeightPt / 72.0, 4),
                },
            });
        }
        catch (Exception ex)
        {
            _log.LogWarning(ex, "[ERP] stamp-bounds failed for {Url}", url);
            return StatusCode(502, new { error = ex.Message });
        }
    }

    public sealed class ProductBoxesRequest
    {
        public string? Url { get; set; }
        public string? DocType { get; set; }
        public List<ProductBoxLineItem>? LineItems { get; set; }
    }

    public sealed class ProductBoxLineItem
    {
        public string? Code { get; set; }
        public string? Description { get; set; }
    }

    /// <summary>
    /// Returns an anchor box per product found in the PDF so the UI can render a "+ Stock"
    /// marker beside each one. Picking slips are detected by the leading MSPC code; other
    /// doc types are matched against the supplied line items. url may be an SP HTTPS URL or
    /// an absolute local file path.
    /// </summary>
    [HttpPost("/api/erp/product-boxes")]
    public async Task<IActionResult> GetProductBoxes([FromBody] ProductBoxesRequest req, CancellationToken ct)
    {
        if (req is null || string.IsNullOrWhiteSpace(req.Url))
            return BadRequest(new { error = "url is required" });

        try
        {
            var bytes = req.Url.StartsWith("http", StringComparison.OrdinalIgnoreCase)
                ? await _sp.DownloadSpFileAsync(req.Url, ct)
                : await System.IO.File.ReadAllBytesAsync(req.Url, ct);

            var hints = (req.LineItems ?? new())
                .Select(li => new PickingSlipEnricher.ProductHint(li.Code, li.Description))
                .ToList();

            var boxes = PickingSlipEnricher.ExtractProductBoxes(bytes, req.DocType, hints);

            return Ok(new
            {
                products = boxes.Select(b => new
                {
                    pageIndex   = b.PageIndex,
                    leftFrac    = Math.Round(b.LeftFrac,   4),
                    topFrac     = Math.Round(b.TopFrac,    4),
                    widthFrac   = Math.Round(b.WidthFrac,  4),
                    heightFrac  = Math.Round(b.HeightFrac, 4),
                    productName = b.ProductName,
                    mspc        = b.Mspc,
                }).ToList(),
            });
        }
        catch (Exception ex)
        {
            _log.LogWarning(ex, "[ERP] product-boxes failed for {Url}", req.Url);
            return StatusCode(502, new { error = ex.Message });
        }
    }

    /// <summary>
    /// Returns page count and dimensions (points + inches) for any PDF.
    /// url may be a SharePoint HTTPS URL or an absolute local file path.
    /// </summary>
    [HttpGet("/api/erp/page-info")]
    public async Task<IActionResult> GetPageInfo([FromQuery] string url, CancellationToken ct)
    {
        if (string.IsNullOrWhiteSpace(url))
            return BadRequest(new { error = "url parameter is required" });

        try
        {
            var bytes = url.StartsWith("http", StringComparison.OrdinalIgnoreCase)
                ? await _sp.DownloadSpFileAsync(url, ct)
                : await System.IO.File.ReadAllBytesAsync(url, ct);

            var dims = PickingSlipEnricher.GetPageDimensions(bytes);
            return Ok(new
            {
                pageCount = dims.Count,
                pages = dims.Select((d, i) => new
                {
                    pageIndex = i,
                    widthPt   = Math.Round(d.WidthPt,  2),
                    heightPt  = Math.Round(d.HeightPt, 2),
                    widthIn   = Math.Round(d.WidthPt  / 72.0, 4),
                    heightIn  = Math.Round(d.HeightPt / 72.0, 4),
                }),
            });
        }
        catch (Exception ex)
        {
            _log.LogWarning(ex, "[ERP] page-info failed for {Url}", url);
            return StatusCode(502, new { error = ex.Message });
        }
    }

    public sealed class FooterPreviewRequest
    {
        /// <summary>Local file path or SharePoint HTTPS URL of the source PDF.</summary>
        public string? Url { get; set; }
        /// <summary>Override footer text. Falls back to config ErpFooter:Text when omitted.</summary>
        public string? Text { get; set; }
        public double? FontSizePt { get; set; }
        public double? SideMarginPt { get; set; }
        public double? BottomMarginPt { get; set; }
        public double? BoxHeightPt { get; set; }
        public bool? EveryPage { get; set; }
        public int? OnlyPageIndex { get; set; }
        public bool? TopRule { get; set; }
        public bool? Center { get; set; }
    }

    /// <summary>
    /// Stamps the T&amp;C footer onto a PDF and returns the stamped bytes — a visual-tuning
    /// harness for the ErpFooter feature. Geometry params override config so the look can be
    /// iterated without redeploying. url may be a local path or an SP HTTPS URL.
    /// </summary>
    [HttpPost("/api/erp/footer-preview")]
    public async Task<IActionResult> FooterPreview([FromBody] FooterPreviewRequest req, CancellationToken ct)
    {
        if (req is null || string.IsNullOrWhiteSpace(req.Url))
            return BadRequest(new { error = "url is required (local path or SP HTTPS URL)" });

        var text = !string.IsNullOrWhiteSpace(req.Text) ? req.Text : _config["ErpFooter:Text"];
        if (string.IsNullOrWhiteSpace(text))
            return BadRequest(new { error = "no text supplied and ErpFooter:Text is not configured" });

        try
        {
            var bytes = req.Url.StartsWith("http", StringComparison.OrdinalIgnoreCase)
                ? await _sp.DownloadSpFileAsync(req.Url, ct)
                : await System.IO.File.ReadAllBytesAsync(req.Url, ct);

            var opts = new ErpDocumentFooterService.FooterOptions
            {
                Text           = text,
                FontSizePt     = req.FontSizePt     ?? _config.GetValue("ErpFooter:FontSizePt",     7.0),
                SideMarginPt   = req.SideMarginPt   ?? _config.GetValue("ErpFooter:SideMarginPt",   36.0),
                BottomMarginPt = req.BottomMarginPt ?? _config.GetValue("ErpFooter:BottomMarginPt", 14.0),
                BoxHeightPt    = req.BoxHeightPt    ?? _config.GetValue("ErpFooter:BoxHeightPt",    30.0),
                EveryPage      = req.EveryPage      ?? _config.GetValue("ErpFooter:EveryPage",      true),
                OnlyPageIndex  = req.OnlyPageIndex,
                TopRule        = req.TopRule        ?? _config.GetValue("ErpFooter:TopRule",        true),
                Center         = req.Center         ?? _config.GetValue("ErpFooter:Center",         true),
            };

            var stamped = ErpDocumentFooterService.StampFooter(bytes, opts, _log);
            return File(stamped, "application/pdf");
        }
        catch (Exception ex)
        {
            _log.LogWarning(ex, "[ERP] footer-preview failed for {Url}", req.Url);
            return StatusCode(502, new { error = ex.Message });
        }
    }

    /// <summary>
    /// Proxies a SharePoint PDF download using app-only credentials.
    /// Clients cannot fetch SharePoint WebUrls directly (no auth); this endpoint adds the Bearer token.
    /// </summary>
    // Short-TTL cache of the assembled (downloaded + optionally FAB-appended) PDF, keyed by url+appendFabs.
    // Re-opens (and the upload/enrich follow-up re-renders) skip the SharePoint download AND the FAB render
    // — the dominant cost for FAB-heavy / multi-page slips (#4). Bounded by lazy eviction.
    private static readonly System.Collections.Concurrent.ConcurrentDictionary<string, (byte[] Bytes, DateTimeOffset At)> _pdfCache = new();
    private static readonly TimeSpan _pdfCacheTtl = TimeSpan.FromMinutes(10);

    [HttpGet("/api/erp/pdf")]
    public async Task<IActionResult> GetPdf([FromQuery] string url, [FromQuery] bool appendFabs, [FromQuery] string? customerName, CancellationToken ct)
    {
        if (string.IsNullOrWhiteSpace(url))
            return BadRequest(new { error = "url query parameter is required" });

        var cacheKey = $"{url}|{appendFabs}|{customerName ?? ""}";
        if (_pdfCache.TryGetValue(cacheKey, out var hit) && DateTimeOffset.UtcNow - hit.At <= _pdfCacheTtl)
            return File(hit.Bytes, "application/pdf");

        try
        {
            var bytes = await _sp.DownloadSpFileAsync(url, ct);
            if (appendFabs)
            {
                try { bytes = PickingSlipFabAppender.AppendFabDrawings(bytes, _log, customerName: customerName); }
                catch (Exception ex) { _log.LogWarning(ex, "[ERP] FAB drawing append failed for {Url}", url); }
            }
            var now = DateTimeOffset.UtcNow;
            _pdfCache[cacheKey] = (bytes, now);
            foreach (var kv in _pdfCache)
                if (now - kv.Value.At > _pdfCacheTtl) _pdfCache.TryRemove(kv.Key, out _);
            return File(bytes, "application/pdf");
        }
        catch (Exception ex)
        {
            _log.LogWarning(ex, "[ERP] PDF proxy download failed for {Url}", url);
            return StatusCode(502, new { error = ex.Message });
        }
    }

    /// <summary>
    /// Develops the slip's deduped <c>FAB:</c> notes into ONE combined DXF (parts laid out left-to-right,
    /// 1" apart, bottom-aligned) and returns the bytes (base64) plus the part slugs. The client saves
    /// this as <c>{HSK#}.dxf</c> in the OneDrive CAD folder and stamps the slip. Returns
    /// <c>ok=false</c> (200) when the slip has no developable FAB notes so the client just skips
    /// generation; 502 only on a download/build error.
    /// </summary>
    [HttpGet("/api/erp/fab-dxf")]
    public async Task<IActionResult> GetFabDxf([FromQuery] string url, CancellationToken ct)
    {
        if (string.IsNullOrWhiteSpace(url))
            return BadRequest(new { error = "url query parameter is required" });

        try
        {
            var bytes = await _sp.DownloadSpFileAsync(url, ct);
            PickingSlipEnricher.EnsureFontResolver();   // FlatPattern.Develop is geometry-only, but keep parity

            var notes = PickingSlipFabAppender.GetFabNotes(bytes, _log);
            if (notes.Count == 0) return Ok(new { ok = false, reason = "no FAB notes" });

            var built = FabDxfBuilder.Build(notes, _log);
            if (built is null) return Ok(new { ok = false, reason = "no developable parts" });

            _log.LogInformation("[FAB-DXF] built combined DXF for {Url}: {N} part(s) [{Parts}]",
                url, built.Parts.Count, string.Join(", ", built.Parts));
            return Ok(new
            {
                ok        = true,
                partCount = built.Parts.Count,
                parts     = built.Parts,
                dxfBase64 = Convert.ToBase64String(built.Dxf),
            });
        }
        catch (Exception ex)
        {
            _log.LogWarning(ex, "[FAB-DXF] generation failed for {Url}", url);
            return StatusCode(502, new { error = ex.Message });
        }
    }

    /// <summary>
    /// Returns a single ERP document record from SharePoint by its SP item ID.
    /// Used by Shredder to backfill LineItemsJson when a bus notification arrived without it.
    /// </summary>
    [HttpGet("/api/erp/documents/{spItemId}")]
    public async Task<IActionResult> GetDocument(string spItemId, CancellationToken ct)
    {
        var doc = await _sp.GetErpDocumentByIdAsync(spItemId, ct);
        if (doc is null) return NotFound();
        return Ok(doc);
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

    /// <summary>Rollout migration: backfill ALL native *Dt dateTime columns on this service's lists
    /// (ErpDocuments ReceivedAt + DocumentDate, PhoneCallLog ReceivedAt, Messages MsgTime) from their legacy
    /// text columns. Registry-driven + idempotent; the self-healing sweep runs the same set automatically.
    /// (Inquiry-cluster columns: POST /api/inquiries/backfill-datetime-columns.)</summary>
    [HttpPost("/api/erp/backfill-datetime-columns")]
    public async Task<IActionResult> BackfillDateTimeColumns(CancellationToken ct)
    {
        var (scanned, patched, failed) = await _sp.BackfillAllDateTimeColumnsAsync(ct);
        return Ok(new { scanned, patched, failed });
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

    /// <summary>
    /// Removes the N most-recently-modified entries from the in-memory processed-file cache
    /// so those files are picked up again on the next scan without touching any other files.
    /// Defaults to count=1.  Use before running a targeted smoke-test scan.
    /// </summary>
    [HttpDelete("/api/erp/processed-cache")]
    public IActionResult RemoveLastProcessedKeys([FromQuery] int count = 1)
    {
        var removed = _fw.RemoveLastProcessedKeys(count);
        _log.LogInformation("[ERP] Removed {Count} most-recent key(s) from processed-file cache", removed);
        return Ok(new { removed });
    }
}
