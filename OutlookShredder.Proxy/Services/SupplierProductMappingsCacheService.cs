using OutlookShredder.Proxy.Models;

namespace OutlookShredder.Proxy.Services;

/// <summary>
/// Singleton background service that loads the SupplierProductMappings SP list into memory
/// and refreshes every 5 minutes. Provides O(1) lookup by (supplierName, supplierTerm).
///
/// This cache is the highest-priority path in MatchProductAsync — user-confirmed mappings
/// from the MSPC remap dialog (≠ badge) take precedence over token scoring.
/// </summary>
public class SupplierProductMappingsCacheService : BackgroundService
{
    private readonly SharePointService _sp;
    private readonly ILogger<SupplierProductMappingsCacheService> _log;

    // Keyed by (supplierName, supplierTerm) — both lowercased for case-insensitive lookup.
    // Reference swap is atomic — no lock needed for reads.
    private volatile IReadOnlyDictionary<(string, string), SupplierProductMappingEntry> _cache
        = new Dictionary<(string, string), SupplierProductMappingEntry>();

    public SupplierProductMappingsCacheService(
        SharePointService sp,
        ILogger<SupplierProductMappingsCacheService> log)
    {
        _sp  = sp;
        _log = log;
    }

    protected override async Task ExecuteAsync(CancellationToken stoppingToken)
    {
        while (!stoppingToken.IsCancellationRequested)
        {
            await RefreshAsync();
            try { await Task.Delay(TimeSpan.FromMinutes(5), stoppingToken); }
            catch (OperationCanceledException) { break; }
        }
    }

    private async Task RefreshAsync()
    {
        try
        {
            var entries = await _sp.ReadSupplierProductMappingsAsync();
            var dict = new Dictionary<(string, string), SupplierProductMappingEntry>(entries.Count);
            foreach (var e in entries)
            {
                var key = (e.SupplierName.ToLowerInvariant(), e.SupplierTerm.ToLowerInvariant());
                dict[key] = e; // last write wins on duplicate keys
            }
            _cache = dict;
            _log.LogInformation("[MappingsCache] Refreshed — {Count} mapping(s) loaded", dict.Count);
        }
        catch (Exception ex)
        {
            _log.LogWarning(ex, "[MappingsCache] Refresh failed — stale cache will be used");
        }
    }

    /// <summary>
    /// Returns the cached mapping for (supplierName, supplierTerm), or null if not found.
    /// Lookup is case-insensitive on both keys. Call site does NOT need to lock.
    /// </summary>
    public SupplierProductMappingEntry? TryGetMapping(string supplierName, string supplierTerm)
    {
        var key = (supplierName.ToLowerInvariant(), supplierTerm.ToLowerInvariant());
        return _cache.TryGetValue(key, out var entry) ? entry : null;
    }

    /// <summary>
    /// Forces a synchronous cache refresh. Called after a new mapping is written so
    /// the next MatchProductAsync call sees the new entry without waiting 5 minutes.
    /// </summary>
    public Task ForceRefreshAsync() => RefreshAsync();
}
