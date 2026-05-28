using OutlookShredder.Proxy.Models;

namespace OutlookShredder.Proxy.Services;

/// <summary>
/// Implemented by every persistent cache service so CacheController can query
/// all registered caches for status, rebuild, and delta operations.
/// </summary>
public interface ICacheStatusProvider
{
    string Name { get; }
    string DisplayName { get; }
    int SchemaVersion { get; }
    int ItemCount { get; }
    DateTime? CacheBuiltUtc { get; }
    DateTime? LastDeltaUtc { get; }
    bool IsLoading { get; }

    Task ForceRebuildAsync(CancellationToken ct = default);
    Task ForceDeltaAsync(CancellationToken ct = default);
}

public static class CacheStatusHelper
{
    public static CacheStatusDto ToDto(this ICacheStatusProvider p, int fullRefreshIntervalDays)
    {
        var status = ComputeStatus(p, fullRefreshIntervalDays);
        return new CacheStatusDto
        {
            Name          = p.Name,
            DisplayName   = p.DisplayName,
            SchemaVersion = p.SchemaVersion,
            ItemCount     = p.ItemCount,
            CacheBuiltUtc = p.CacheBuiltUtc,
            LastDeltaUtc  = p.LastDeltaUtc,
            Status        = status,
        };
    }

    private static string ComputeStatus(ICacheStatusProvider p, int fullRefreshIntervalDays)
    {
        if (p.IsLoading) return "loading";
        if (p.CacheBuiltUtc is null) return "cold";
        var now = DateTime.UtcNow;
        var fullInterval = TimeSpan.FromDays(fullRefreshIntervalDays);
        if (now - p.CacheBuiltUtc > fullInterval) return "aged";
        var lastDelta = p.LastDeltaUtc ?? p.CacheBuiltUtc;
        if (now - lastDelta > fullInterval / 2) return "stale";
        return "fresh";
    }
}
