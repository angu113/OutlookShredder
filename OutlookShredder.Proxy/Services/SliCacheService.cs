using OutlookShredder.Proxy.Services;

namespace OutlookShredder.Proxy.Services;

/// <summary>
/// In-memory cache for the full joined SLI+SR dataset.  Populated at startup by
/// iterating all SharePoint pages once; subsequent /api/items calls are served
/// from this list in O(1) with no SP round-trips.
///
/// TTL is 5 minutes.  forceRefresh=true on /api/items (manual Refresh button)
/// bypasses the cache and re-fetches from SP live, then repopulates the cache.
///
/// Thread safety: a SemaphoreSlim(1) serialises population; TryGet is a
/// non-blocking volatile read safe to call at any time.
/// </summary>
public sealed class SliCacheService
{
    private readonly SharePointService           _sp;
    private readonly ILogger<SliCacheService>    _log;

    private volatile List<Dictionary<string, object?>>? _items;
    private DateTime  _populatedAt = DateTime.MinValue;
    private static readonly TimeSpan Ttl = TimeSpan.FromMinutes(5);
    private readonly SemaphoreSlim _sem = new(1, 1);

    public SliCacheService(SharePointService sp, ILogger<SliCacheService> log)
    {
        _sp  = sp;
        _log = log;
    }

    /// <summary>Returns the cached list when fresh, otherwise null (cache miss).</summary>
    public List<Dictionary<string, object?>>? TryGet()
        => (_items is not null && DateTime.UtcNow < _populatedAt + Ttl) ? _items : null;

    /// <summary>
    /// Populates the cache from SP.  Idempotent when already fresh unless
    /// <paramref name="force"/> is true.  Safe to call concurrently.
    /// </summary>
    public async Task PopulateAsync(bool force = false, CancellationToken ct = default)
    {
        // Fast path: already fresh and not forced.
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
            _log.LogInformation("[SliCache] Populated {Count} rows in {Pages} pages ({Ms}ms)",
                all.Count, pages, sw.ElapsedMilliseconds);
        }
        catch (Exception ex)
        {
            _log.LogWarning(ex, "[SliCache] Population failed — cache remains empty");
        }
        finally
        {
            _sem.Release();
        }
    }

    /// <summary>Clears the cache so the next TryGet returns null.</summary>
    public void Invalidate()
    {
        _items       = null;
        _populatedAt = DateTime.MinValue;
        _log.LogInformation("[SliCache] Invalidated");
    }
}
