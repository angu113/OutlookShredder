using Microsoft.Extensions.Hosting;
using Microsoft.Extensions.Logging;
using OutlookShredder.Proxy.Models;

namespace OutlookShredder.Proxy.Services;

/// <summary>
/// In-memory serving cache for the <c>SalesOrderHistory</c> list (indexed by normalized customer name) plus
/// the resumable bulk-import runner. Mirrors <see cref="CustomerCacheService"/> / SupplierUnreadIndexService:
/// build once at startup, atomic <c>volatile</c> snapshot (readers never block), a 5-minute safety re-scan,
/// and a cold fallback to a live SP read before the first build completes. Serves the Raptor incoming-call
/// card via <see cref="GetOrdersForCustomerAsync"/>.
/// </summary>
public sealed class SalesOrderHistoryService : IHostedService
{
    private readonly SharePointService _sp;
    private readonly ILogger<SalesOrderHistoryService> _log;

    private volatile Snapshot _snap = Snapshot.Empty;
    private volatile bool _built;
    private Timer? _timer;

    private static readonly TimeSpan SafetyInterval = TimeSpan.FromMinutes(5);

    public SalesOrderHistoryService(SharePointService sp, ILogger<SalesOrderHistoryService> log)
    {
        _sp  = sp;
        _log = log;
        // Signal 1 — local write-through: an order appended on THIS proxy (FileWatcher) merges into the
        // serving cache directly (a proxy filters out its own bus messages, so it must self-update).
        _sp.SalesOrderAppended += MergeOrder;
    }

    public async Task StartAsync(CancellationToken ct)
    {
        // Critical-path: StartAsync awaits this SP read before the host finishes starting.
        await StartupTimings.MeasureAsync("sales-order-history", null, () => SafeRefreshAsync(), critical: true);
        _timer = new Timer(_ => _ = SafeRefreshAsync(), null, SafetyInterval, SafetyInterval);
    }

    public Task StopAsync(CancellationToken ct)
    {
        _sp.SalesOrderAppended -= MergeOrder;
        _timer?.Dispose();
        return Task.CompletedTask;
    }

    // ── Serving ──────────────────────────────────────────────────────────────

    /// <summary>The customer's recent orders + summary header, newest-first, capped to <paramref name="top"/>.
    /// Served from the in-memory snapshot (cold fallback to a one-off live read during warmup).</summary>
    public async Task<CustomerOrdersResponse> GetOrdersForCustomerAsync(
        string customer, int top, CancellationToken ct = default)
    {
        var snap = _built ? _snap : await BuildColdAsync(ct);
        snap.ByCustomer.TryGetValue(NormName(customer), out var orders);
        return Project(orders ?? [], top);
    }

    private static CustomerOrdersResponse Project(IReadOnlyList<SalesOrderRecord> orders, int top)
    {
        var sorted = orders
            .OrderByDescending(o => o.OrderDate ?? DateTimeOffset.MinValue)
            .ToList();

        double lifetimeNet = sorted.Sum(o => o.DisplayAmount ?? 0);
        var lastDate = sorted.FirstOrDefault(o => o.OrderDate is not null)?.OrderDate;

        var summary = new CustomerOrdersSummary(
            sorted.Count, Math.Round(lifetimeNet, 2), lastDate?.ToString("yyyy-MM-dd"));

        var dtos = sorted
            .Take(Math.Max(0, top))
            .Select(o => new CustomerOrderDto(
                o.OrderId,
                o.OrderDate?.ToString("yyyy-MM-dd"),
                o.DisplayAmount,
                o.NetAmount,
                o.GrossAmount,
                o.Status))
            .ToList();

        return new CustomerOrdersResponse(summary, dtos);
    }

    // ── Cache build ───────────────────────────────────────────────────────────

    public async Task RefreshAsync(CancellationToken ct = default)
    {
        var snap = await LoadSnapshotAsync(ct);
        _snap  = snap;
        _built = true;
        _log.LogInformation("[SOHistory] cache built: {Orders} orders across {Customers} customers",
            snap.Count, snap.ByCustomer.Count);
    }

    private async Task SafeRefreshAsync()
    {
        try { await RefreshAsync(CancellationToken.None); }
        catch (Exception ex) { _log.LogWarning(ex, "[SOHistory] cache refresh failed — old data retained"); }
    }

    private async Task<Snapshot> BuildColdAsync(CancellationToken ct)
    {
        try
        {
            var snap = await LoadSnapshotAsync(ct);
            _snap  = snap;
            _built = true;
            return snap;
        }
        catch (Exception ex)
        {
            _log.LogWarning(ex, "[SOHistory] cold load failed — serving empty");
            return Snapshot.Empty;
        }
    }

    private async Task<Snapshot> LoadSnapshotAsync(CancellationToken ct)
    {
        var rows = await _sp.ReadAllSalesOrdersAsync(ct);
        var byCustomer = rows
            .GroupBy(r => NormName(r.CustomerName), StringComparer.OrdinalIgnoreCase)
            .ToDictionary(g => g.Key, g => (IReadOnlyList<SalesOrderRecord>)g.ToList(),
                StringComparer.OrdinalIgnoreCase);
        return new Snapshot(byCustomer, rows.Count);
    }

    // ── Write-through merge (signals 1 & 2) ─────────────────────────────────────

    private readonly object _mergeLock = new();

    /// <summary>Merges a single order into the serving cache, insert-if-absent by OrderId, with NO SharePoint
    /// read — O(1) per order (the misdesign to avoid is a full 35k-row refresh per order per proxy). Used by
    /// both the local write-through (<see cref="SharePointService.SalesOrderAppended"/>) and the cross-proxy
    /// bus branch. Copy-on-write of the one affected customer's list keeps reads lock-free.</summary>
    public void MergeOrder(SalesOrderRecord rec)
    {
        if (rec is null || string.IsNullOrWhiteSpace(rec.OrderId) || string.IsNullOrWhiteSpace(rec.CustomerName))
            return;

        lock (_mergeLock)
        {
            if (!_built) return;   // cold — the next build/refresh reads the persisted row from SP anyway
            var snap = _snap;
            var key  = NormName(rec.CustomerName);
            snap.ByCustomer.TryGetValue(key, out var existing);
            existing ??= [];
            if (existing.Any(o => string.Equals(o.OrderId, rec.OrderId, StringComparison.OrdinalIgnoreCase)))
                return;   // already present — insert-if-absent

            var newList = new List<SalesOrderRecord>(existing) { rec };
            var newDict = new Dictionary<string, IReadOnlyList<SalesOrderRecord>>(
                snap.ByCustomer, StringComparer.OrdinalIgnoreCase) { [key] = newList };
            _snap = new Snapshot(newDict, snap.Count + 1);
            _log.LogDebug("[SOHistory] merged order {Order} for {Customer} (cache now {N} orders)",
                rec.OrderId, rec.CustomerName, snap.Count + 1);
        }
    }

    // ── Bulk import (background, resumable) ─────────────────────────────────────

    private readonly object _importLock = new();
    private volatile SalesOrderImportStatus _import = SalesOrderImportStatus.Idle;

    public SalesOrderImportStatus ImportStatus => _import;

    /// <summary>Kicks off a background bulk load of the already-parsed rows. Returns false (with a message)
    /// when a load is already running. The serving cache is refreshed when the load completes.</summary>
    public bool StartImport(IReadOnlyList<CustomerImportService.SalesOrderRow> rows, out string message)
    {
        lock (_importLock)
        {
            if (_import.Running)
            {
                message = "A SalesOrderHistory load is already running.";
                return false;
            }
            _import = SalesOrderImportStatus.Started(rows.Count);
        }
        message = $"Started background load of {rows.Count} parsed row(s).";
        _ = Task.Run(() => RunImportAsync(rows));
        return true;
    }

    private async Task RunImportAsync(IReadOnlyList<CustomerImportService.SalesOrderRow> rows)
    {
        try
        {
            var (added, already, failed) = await _sp.InsertSalesOrdersAsync(
                rows,
                onAdded: n => { if (n % 250 == 0) _import = _import with { Added = n }; },
                CancellationToken.None);

            _import = _import with
            {
                Running = false, Added = added, AlreadyPresent = already, Failed = failed,
                Message = $"Done: {added} added, {already} already present, {failed} failed.",
            };
            await SafeRefreshAsync();   // publish the freshly-loaded rows to the serving cache
        }
        catch (Exception ex)
        {
            _log.LogError(ex, "[SOHistory] bulk load failed");
            _import = _import with { Running = false, Message = $"Failed: {ex.Message}" };
        }
    }

    // ── Helpers ────────────────────────────────────────────────────────────────

    // Collapse-whitespace + lowercase — the dup-bug lesson from the contact import, and it dodges OData
    // apostrophe-escaping for names like "Steve John's Landscaping". Matches SharePointService.NormName.
    private static string NormName(string s) =>
        string.Join(' ', (s ?? "").Split((char[]?)null, StringSplitOptions.RemoveEmptyEntries)).ToLowerInvariant();

    private sealed record Snapshot(
        Dictionary<string, IReadOnlyList<SalesOrderRecord>> ByCustomer, int Count)
    {
        public static Snapshot Empty { get; } = new(new(StringComparer.OrdinalIgnoreCase), 0);
    }
}

/// <summary>Live status of the background bulk load, surfaced by the import status endpoint.</summary>
public sealed record SalesOrderImportStatus(
    bool Running, int Total, int Added, int AlreadyPresent, int Failed, string Message)
{
    public static SalesOrderImportStatus Idle { get; } = new(false, 0, 0, 0, 0, "idle");
    public static SalesOrderImportStatus Started(int total) => new(true, total, 0, 0, 0, "running");
}
