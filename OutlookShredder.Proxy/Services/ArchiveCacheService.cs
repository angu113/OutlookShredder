using System.Text.Json;
using OutlookShredder.Proxy.Models;

namespace OutlookShredder.Proxy.Services;

/// <summary>
/// Persistent in-memory cache of Complete=true RFQ data.
///
/// Holds:
///   _refs    — all Complete=true RFQ References (for ref-level search)
///   _rli     — RFQ Line Items keyed by RFQ ID (for product/MSPC search)
///
/// SLI rows are NOT stored here; they come from SliCacheService at search time
/// (filtered to the matched RFQ IDs), avoiding a duplicate SP read.
///
/// Disk persistence: %LOCALAPPDATA%\ShredderData\Cache\v1\archive.json
/// Schema change → bump ArchiveSchemaVersion → new dir → old cache silently ignored.
/// </summary>
public sealed class ArchiveCacheService : IHostedService, ICacheStatusProvider
{
    private const int ArchiveSchemaVersion = 1;
    private static readonly TimeSpan DefaultFullRefresh = TimeSpan.FromDays(7);
    private static readonly TimeSpan DeltaBuffer        = TimeSpan.FromMinutes(10);

    private readonly SharePointService             _sp;
    private readonly SliCacheService               _sli;
    private readonly ILogger<ArchiveCacheService>  _log;
    private readonly DiskBackedCache<ArchiveCacheFile> _disk;

    private readonly SemaphoreSlim _sem = new(1, 1);

    private volatile List<ArchiveRfqRef> _refs = [];
    private volatile Dictionary<string, List<RliContextItem>> _rli = new(StringComparer.OrdinalIgnoreCase);
    private DateTime? _cacheBuiltUtc;
    private DateTime? _lastDeltaUtc;
    private bool      _isLoading;

    // ICacheStatusProvider
    public string   Name          => "archive";
    public string   DisplayName   => "Archive RFQ Data";
    public int      SchemaVersion => ArchiveSchemaVersion;
    public int      ItemCount     => _refs.Count;
    public DateTime? CacheBuiltUtc => _cacheBuiltUtc;
    public DateTime? LastDeltaUtc  => _lastDeltaUtc;
    public bool     IsLoading     => _isLoading;

    // Search access
    public IReadOnlyList<ArchiveRfqRef>                             Refs => _refs;
    public IReadOnlyDictionary<string, List<RliContextItem>>        Rli  => _rli;

    private static readonly JsonSerializerOptions JsonOpts = new()
    {
        PropertyNameCaseInsensitive = true,
        WriteIndented               = false,
    };

    public ArchiveCacheService(
        SharePointService sp,
        SliCacheService sli,
        ILogger<ArchiveCacheService> log)
    {
        _sp  = sp;
        _sli = sli;
        _log = log;

        var dir = Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
            "ShredderData", "Cache", $"v{ArchiveSchemaVersion}");
        _disk = new DiskBackedCache<ArchiveCacheFile>(dir, "archive", log, JsonOpts);

        TryWarmFromDisk();
    }

    private void TryWarmFromDisk()
    {
        var file = _disk.TryLoad();
        if (file is null || file.SchemaVersion != ArchiveSchemaVersion) return;

        _refs          = file.Refs ?? [];
        _rli           = new Dictionary<string, List<RliContextItem>>(
            file.RliByRfqId ?? [], StringComparer.OrdinalIgnoreCase);
        _cacheBuiltUtc = file.CacheBuiltUtc;
        _lastDeltaUtc  = file.LastDeltaUtc;

        _log.LogInformation("[ArchiveCache] Warmed from disk — {Count} refs, built {At:yyyy-MM-dd}",
            _refs.Count, _cacheBuiltUtc);
    }

    // ── IHostedService ────────────────────────────────────────────────────────

    public Task StartAsync(CancellationToken ct)
    {
        _ = Task.Run(() => BackgroundWarmAsync(ct), ct);
        return Task.CompletedTask;
    }

    public Task StopAsync(CancellationToken ct) => Task.CompletedTask;

    private async Task BackgroundWarmAsync(CancellationToken ct)
    {
        try
        {
            await Task.Delay(TimeSpan.FromSeconds(10), ct); // let SharePoint pre-warm settle first
            if (ct.IsCancellationRequested) return;

            var now = DateTime.UtcNow;
            var needsFull = _cacheBuiltUtc is null
                         || now - _cacheBuiltUtc > DefaultFullRefresh;

            if (needsFull)
                await FetchAndReplaceAsync(ct);
            else
                await RunDeltaAsync(ct);
        }
        catch (OperationCanceledException) { }
        catch (Exception ex)
        {
            _log.LogWarning(ex, "[ArchiveCache] Background warm failed (non-fatal)");
        }
    }

    // ── ICacheStatusProvider ─────────────────────────────────────────────────

    public async Task ForceRebuildAsync(CancellationToken ct = default)
        => await FetchAndReplaceAsync(ct);

    public async Task ForceDeltaAsync(CancellationToken ct = default)
        => await RunDeltaAsync(ct);

    // ── Internal refresh ──────────────────────────────────────────────────────

    private async Task FetchAndReplaceAsync(CancellationToken ct)
    {
        await _sem.WaitAsync(ct);
        _isLoading = true;
        try
        {
            var sw = System.Diagnostics.Stopwatch.StartNew();
            _log.LogInformation("[ArchiveCache] Full rebuild started");

            var refs = await _sp.ReadCompletedRfqReferencesAsync(ct);
            var rli  = await _sp.ReadAllRliAsync(ct);

            _refs          = refs;
            _rli           = rli;
            _cacheBuiltUtc = DateTime.UtcNow;
            _lastDeltaUtc  = DateTime.UtcNow;

            _log.LogInformation("[ArchiveCache] Full rebuild done — {Count} refs in {Ms}ms",
                refs.Count, sw.ElapsedMilliseconds);

            await SaveToDiskAsync();
        }
        finally
        {
            _isLoading = false;
            _sem.Release();
        }
    }

    private async Task RunDeltaAsync(CancellationToken ct)
    {
        await _sem.WaitAsync(ct);
        _isLoading = true;
        try
        {
            var since = (_lastDeltaUtc ?? _cacheBuiltUtc ?? DateTime.UtcNow.AddDays(-1))
                        .Subtract(DeltaBuffer);

            _log.LogInformation("[ArchiveCache] Delta refresh since {Since:u}", since);

            var newRefs = await _sp.ReadCompletedRfqReferencesSinceAsync(since, ct);
            if (newRefs.Count > 0)
            {
                var existing = new HashSet<string>(_refs.Select(r => r.RfqId), StringComparer.OrdinalIgnoreCase);
                var added    = newRefs.Where(r => !existing.Contains(r.RfqId)).ToList();
                if (added.Count > 0)
                {
                    _refs = [.._refs, ..added];
                    // Fetch RLI for newly added RFQs
                    foreach (var r in added)
                    {
                        try
                        {
                            var rliItems = await _sp.ReadRfqLineItemsByRfqIdAsync(r.RfqId);
                            if (rliItems.Count > 0)
                                _rli[r.RfqId] = rliItems;
                        }
                        catch { /* non-fatal */ }
                    }
                    _log.LogInformation("[ArchiveCache] Delta: merged {Count} new refs", added.Count);
                }
            }

            _lastDeltaUtc = DateTime.UtcNow;
            await SaveToDiskAsync();
        }
        finally
        {
            _isLoading = false;
            _sem.Release();
        }
    }

    private Task SaveToDiskAsync()
    {
        var file = new ArchiveCacheFile
        {
            SchemaVersion = ArchiveSchemaVersion,
            CacheBuiltUtc = _cacheBuiltUtc,
            LastDeltaUtc  = _lastDeltaUtc,
            Refs          = _refs,
            RliByRfqId    = new Dictionary<string, List<RliContextItem>>(_rli),
        };
        return _disk.SaveAsync(file);
    }

    // ── Bus event handlers (called by ExtractController / RfqNotificationService) ──

    /// <summary>Called when an RFQ is marked Complete — add it to the archive.</summary>
    public async Task OnRfqCompletedAsync(string rfqId)
    {
        try
        {
            var existing = _refs.FirstOrDefault(r =>
                string.Equals(r.RfqId, rfqId, StringComparison.OrdinalIgnoreCase));
            if (existing is not null) return; // already in archive

            var refs = await _sp.ReadCompletedRfqReferencesAsync(CancellationToken.None);
            var match = refs.FirstOrDefault(r =>
                string.Equals(r.RfqId, rfqId, StringComparison.OrdinalIgnoreCase));
            if (match is null) return;

            _refs = [.._refs, match];

            var rliItems = await _sp.ReadRfqLineItemsByRfqIdAsync(rfqId);
            if (rliItems.Count > 0) _rli[rfqId] = rliItems;

            _log.LogInformation("[ArchiveCache] RFQ_COMPLETED: appended {RfqId}", rfqId);
            _ = SaveToDiskAsync();
        }
        catch (Exception ex)
        {
            _log.LogWarning(ex, "[ArchiveCache] OnRfqCompleted({RfqId}) failed", rfqId);
        }
    }

    /// <summary>Called when an RFQ is reactivated — remove it from the archive.</summary>
    public void OnRfqReactivated(string rfqId)
    {
        var before = _refs.Count;
        _refs = _refs.Where(r => !string.Equals(r.RfqId, rfqId, StringComparison.OrdinalIgnoreCase)).ToList();
        _rli.Remove(rfqId);
        if (_refs.Count < before)
        {
            _log.LogInformation("[ArchiveCache] RFQ_REACTIVATED: removed {RfqId}", rfqId);
            _ = SaveToDiskAsync();
        }
    }

    // ── Search ────────────────────────────────────────────────────────────────

    public ArchiveSearchResponse Search(ArchiveSearchRequest req)
    {
        static bool Sub(string? haystack, string? needle)
            => needle is not null && haystack is not null &&
               haystack.Contains(needle, StringComparison.OrdinalIgnoreCase);

        // Get all SLI rows from SliCacheService for use in supplier/product filter
        var allSli  = _sli.GetSnapshot() ?? [];
        var sliByRfq = allSli
            .GroupBy(r => r.TryGetValue("JobReference", out var v) ? v?.ToString() ?? "" : "")
            .ToDictionary(g => g.Key, g => g.ToList(), StringComparer.OrdinalIgnoreCase);

        // 1. Filter refs by ref-level criteria
        IEnumerable<ArchiveRfqRef> filtered = _refs;

        if (!string.IsNullOrWhiteSpace(req.RfqId))
            filtered = filtered.Where(r => r.RfqId.StartsWith(req.RfqId, StringComparison.OrdinalIgnoreCase));
        if (!string.IsNullOrWhiteSpace(req.CustomerName))
            filtered = filtered.Where(r => Sub(r.CustomerName, req.CustomerName));
        if (!string.IsNullOrWhiteSpace(req.Requester))
            filtered = filtered.Where(r => Sub(r.Requester, req.Requester));
        if (!string.IsNullOrWhiteSpace(req.HskNumber))
            filtered = filtered.Where(r => Sub(r.HskNumber, req.HskNumber));
        if (!string.IsNullOrWhiteSpace(req.NotesContains))
            filtered = filtered.Where(r => Sub(r.Notes, req.NotesContains));
        if (req.DateFrom.HasValue)
            filtered = filtered.Where(r => r.DateSent >= req.DateFrom);
        if (req.DateTo.HasValue)
            filtered = filtered.Where(r => r.DateSent <= req.DateTo);

        // Apply cursor for pagination (DateSent < cursor, since we sort DESC)
        if (req.Cursor.HasValue)
            filtered = filtered.Where(r => r.DateSent < req.Cursor);

        // 2. For supplier / product filters, also check SLI + RLI
        var filteredList = filtered.ToList();

        if (!string.IsNullOrWhiteSpace(req.SupplierName) || !string.IsNullOrWhiteSpace(req.Product))
        {
            filteredList = filteredList.Where(r =>
            {
                if (!string.IsNullOrWhiteSpace(req.SupplierName))
                {
                    if (!sliByRfq.TryGetValue(r.RfqId, out var rows)) return false;
                    if (!rows.Any(row =>
                            Sub(row.TryGetValue("SupplierName", out var sn) ? sn?.ToString() : null,
                                req.SupplierName)))
                        return false;
                }

                if (!string.IsNullOrWhiteSpace(req.Product))
                {
                    // Check SLI ProductName
                    bool inSli = sliByRfq.TryGetValue(r.RfqId, out var rows) &&
                        rows.Any(row =>
                            Sub(row.TryGetValue("ProductName", out var pn) ? pn?.ToString() : null,
                                req.Product) ||
                            Sub(row.TryGetValue("CatalogProductName", out var cpn) ? cpn?.ToString() : null,
                                req.Product));
                    // Check RLI ProductName or MSPC
                    bool inRli = _rli.TryGetValue(r.RfqId, out var rliRows) &&
                        rliRows.Any(ri =>
                            Sub(ri.ProductName, req.Product) ||
                            Sub(ri.Mspc,        req.Product));
                    if (!inSli && !inRli) return false;
                }
                return true;
            }).ToList();
        }

        // 3. Sort DESC by DateSent
        filteredList = filteredList
            .OrderByDescending(r => r.DateSent ?? DateTime.MinValue)
            .ToList();

        var totalCount = filteredList.Count;

        // 4. Take one page
        var page     = filteredList.Take(req.PageSize).ToList();
        var pageIds  = page.Select(r => r.RfqId).ToHashSet(StringComparer.OrdinalIgnoreCase);
        var lastDate = page.LastOrDefault()?.DateSent;

        // 5. Gather SLI for matched RFQs
        var sliRows = allSli
            .Where(r => pageIds.Contains(r.TryGetValue("JobReference", out var v) ? v?.ToString() ?? "" : ""))
            .ToList();

        // 6. Gather RLI for matched RFQs
        var rliRows = new List<object>();
        foreach (var rfqId in pageIds)
        {
            if (!_rli.TryGetValue(rfqId, out var items)) continue;
            foreach (var item in items)
                rliRows.Add(new { rfqId, mspc = item.Mspc, product = item.ProductName,
                    units = item.Quantity, sizeOfUnits = item.SizeOfUnits, notes = item.Notes });
        }

        return new ArchiveSearchResponse
        {
            Sli        = sliRows,
            Rli        = rliRows,
            Refs       = page,
            TotalCount = totalCount,
            NextCursor = page.Count == req.PageSize ? lastDate : null,
        };
    }
}
