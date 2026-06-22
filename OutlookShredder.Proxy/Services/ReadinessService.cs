using System.Text.Json;

namespace OutlookShredder.Proxy.Services;

/// <summary>
/// Startup readiness tracker. A cache is "ready" once it has warmed (CacheBuiltUtc != null); SharePoint
/// is ready once <see cref="SharePointService.IsPrewarmed"/>. Serves the snapshot for GET /api/ready and
/// PUSHES a <c>cache-ready</c> SSE event the instant each cache warms (and <c>all-ready</c> once they all
/// have) over the existing /api/events stream — so the client can gate each sub-app on its specific cache
/// instead of firing requests into a cold backend and surfacing timeouts/errors.
/// </summary>
public sealed class ReadinessService : BackgroundService
{
    private readonly IEnumerable<ICacheStatusProvider> _caches;
    private readonly SharePointService _sp;
    private readonly RfqNotificationService _notify;
    private readonly ILogger<ReadinessService> _log;

    private readonly HashSet<string> _announced = new(StringComparer.OrdinalIgnoreCase);
    private bool _allAnnounced;

    public ReadinessService(IEnumerable<ICacheStatusProvider> caches, SharePointService sp,
        RfqNotificationService notify, ILogger<ReadinessService> log)
    {
        _caches = caches;
        _sp = sp;
        _notify = notify;
        _log = log;
    }

    public sealed record Entry(string Id, string Label, bool Ready, int ItemCount);

    /// <summary>Current per-service readiness — SharePoint pre-warm + every registered cache.</summary>
    public IReadOnlyList<Entry> Snapshot()
    {
        var list = new List<Entry> { new("sharepoint", "SharePoint", _sp.IsPrewarmed, 0) };
        foreach (var c in _caches)
            list.Add(new Entry(c.Name, c.DisplayName, c.CacheBuiltUtc is not null, c.ItemCount));
        return list;
    }

    public bool AllReady() => Snapshot().All(e => e.Ready);

    protected override async Task ExecuteAsync(CancellationToken ct)
    {
        while (!ct.IsCancellationRequested)
        {
            try { AnnounceTransitions(); }
            catch (Exception ex) { _log.LogDebug(ex, "[Ready] announce failed"); }
            try { await Task.Delay(500, ct); }
            catch (OperationCanceledException) { break; }
        }
    }

    private void AnnounceTransitions()
    {
        foreach (var e in Snapshot())
            if (e.Ready && _announced.Add(e.Id))
            {
                _notify.BroadcastLocal("cache-ready",
                    JsonSerializer.Serialize(new { id = e.Id, label = e.Label, itemCount = e.ItemCount }));
                _log.LogInformation("[Ready] {Id} ready ({Count} items)", e.Id, e.ItemCount);
            }

        if (!_allAnnounced && AllReady())
        {
            _allAnnounced = true;
            _notify.BroadcastLocal("all-ready", "{}");
            _log.LogInformation("[Ready] all caches warm");
        }
    }
}
