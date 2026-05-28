using System.Text.Json;

namespace OutlookShredder.Proxy.Services;

/// <summary>
/// Two-level cache for the full joined SLI+SR dataset.
///
/// L1 (memory): volatile reference swap — TryGet is non-blocking and safe to call at any time.
/// L2 (disk):   JSON file at %LOCALAPPDATA%\Shredder\Proxy\cache\sli.json — loaded at
///              startup so the first request after a proxy restart is served immediately
///              without waiting for SP pagination. Written (fire-and-forget) after each
///              successful SP refresh.
///
/// Thread safety: SemaphoreSlim(1) serialises SP population so only one fetch runs at a time.
/// TTL is 5 minutes; forceRefresh=true bypasses it and re-fetches from SP live.
/// </summary>
public sealed class SliCacheService : ICacheStatusProvider
{
    private readonly SharePointService           _sp;
    private readonly ILogger<SliCacheService>    _log;
    private readonly DiskBackedCache<List<Dictionary<string, object?>>> _disk;

    private volatile List<Dictionary<string, object?>>? _items;
    private DateTime  _populatedAt  = DateTime.MinValue;
    private DateTime? _cacheBuilt;
    private static readonly TimeSpan Ttl = TimeSpan.FromMinutes(5);
    private readonly SemaphoreSlim _sem = new(1, 1);

    // ICacheStatusProvider
    public string    Name          => "sli";
    public string    DisplayName   => "Live Quote Data";
    public int       SchemaVersion => 1;
    public int       ItemCount     => _items?.Count ?? 0;
    public DateTime? CacheBuiltUtc => _cacheBuilt;
    public DateTime? LastDeltaUtc  => _populatedAt == DateTime.MinValue ? null : _populatedAt;
    public bool      IsLoading     => _sem.CurrentCount == 0;

    public async Task ForceRebuildAsync(CancellationToken ct = default)
        => await PopulateAsync(force: true, ct);
    public async Task ForceDeltaAsync(CancellationToken ct = default)
        => await PopulateAsync(force: true, ct);

    private static readonly JsonSerializerOptions SliJsonOpts = new()
    {
        Converters = { ObjectConverter.Instance }
    };

    public SliCacheService(SharePointService sp, ILogger<SliCacheService> log)
    {
        _sp  = sp;
        _log = log;

        var cacheDir = Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
            "ShredderData", "Cache", "v1");
        _disk = new DiskBackedCache<List<Dictionary<string, object?>>>(cacheDir, "sli", log, SliJsonOpts);

        // Warm L1 from disk immediately so early requests don't block on SP pagination
        var fromDisk = _disk.TryLoad();
        if (fromDisk is { Count: > 0 })
        {
            _items       = fromDisk;
            _populatedAt = DateTime.UtcNow;
            _log.LogInformation("[SliCache] Warmed from disk — {Count} rows (background refresh follows)", fromDisk.Count);
        }
    }

    /// <summary>Returns the cached list when fresh, otherwise null (cache miss).</summary>
    public List<Dictionary<string, object?>>? TryGet()
        => (_items is not null && DateTime.UtcNow < _populatedAt + Ttl) ? _items : null;

    /// <summary>Returns whatever is in memory regardless of TTL. Null if never populated.</summary>
    public List<Dictionary<string, object?>>? GetSnapshot() => _items;

    /// <summary>
    /// Populates the cache from SP.  Idempotent when already fresh unless
    /// <paramref name="force"/> is true.  Safe to call concurrently.
    /// </summary>
    public async Task PopulateAsync(bool force = false, CancellationToken ct = default)
    {
        if (!force && TryGet() is not null) return;

        await _sem.WaitAsync(ct);
        try
        {
            if (!force && TryGet() is not null) return; // double-check under lock

            var sw  = System.Diagnostics.Stopwatch.StartNew();
            var all = new List<Dictionary<string, object?>>();
            string? cursor = null;
            int pages = 0;

            do
            {
                var (items, next) = await _sp.ReadSupplierItemsAsync(5000, cursor, skipDedup: false);
                all.AddRange(items);
                cursor = next;
                pages++;
            }
            while (cursor is not null && !ct.IsCancellationRequested);

            if (ct.IsCancellationRequested) return;

            _items       = all;
            _populatedAt = DateTime.UtcNow;
            _cacheBuilt  ??= DateTime.UtcNow;
            _log.LogInformation("[SliCache] Populated {Count} rows in {Pages} pages ({Ms}ms)",
                all.Count, pages, sw.ElapsedMilliseconds);

            _ = _disk.SaveAsync(all); // fire-and-forget; exceptions are caught inside SaveAsync
        }
        catch (Exception ex)
        {
            _log.LogWarning(ex, "[SliCache] Population failed — cache remains as-is");
        }
        finally
        {
            _sem.Release();
        }
    }

    /// <summary>Clears the in-memory cache so the next TryGet returns null.</summary>
    public void Invalidate()
    {
        _items       = null;
        _populatedAt = DateTime.MinValue;
        _log.LogInformation("[SliCache] Invalidated");
    }

    /// <summary>
    /// Replaces cached rows for a single RFQ without invalidating the rest of the cache.
    /// Called after writing SP rows so the cache immediately reflects the change;
    /// subsequent full-load requests serve correct data without a round-trip to SP.
    /// </summary>
    public void MergeRfqRows(string rfqId, IReadOnlyList<Dictionary<string, object?>> freshRows)
    {
        var current = _items;
        if (current is null) return; // cache is cold — PopulateAsync will handle it

        _items = current
            .Where(r => !MatchesRfq(r, rfqId))
            .Concat(freshRows)
            .ToList();

        _log.LogInformation("[SliCache] MergeRfqRows({RfqId}): {Count} fresh rows merged", rfqId, freshRows.Count);
    }

    /// <summary>
    /// Removes all cached rows for a single RFQ without touching the rest.
    /// Used when SliRows are unavailable in the bus message (fallback path).
    /// </summary>
    public void InvalidateRfq(string rfqId)
    {
        var current = _items;
        if (current is null) return;

        _items = current.Where(r => !MatchesRfq(r, rfqId)).ToList();
        _log.LogInformation("[SliCache] InvalidateRfq({RfqId}): rows removed", rfqId);
    }

    private static bool MatchesRfq(Dictionary<string, object?> row, string rfqId) =>
        (row.TryGetValue("JobReference", out var j) &&
         string.Equals(j?.ToString(), rfqId, StringComparison.OrdinalIgnoreCase)) ||
        (row.TryGetValue("RFQ_ID", out var r) &&
         string.Equals(r?.ToString(), rfqId, StringComparison.OrdinalIgnoreCase));
}
