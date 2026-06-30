using System.Text.Json;
using Microsoft.Extensions.Hosting;
using Microsoft.Extensions.Logging;
using OutlookShredder.Proxy.Models;

namespace OutlookShredder.Proxy.Services;

/// <summary>
/// In-memory serving cache for the <c>SalesOrderHistory</c> list (indexed by normalized customer name) plus
/// the resumable bulk-import runner. Serves the Raptor incoming-call card via <see cref="GetOrdersForCustomerAsync"/>.
///
/// <para><b>Disk-backed (no startup SP read).</b> SalesOrderHistory is a big (~35k row), slowly-changing,
/// MANUAL-load dataset, so reading it all from SharePoint on every startup cost ~29s and used to block the
/// host. Instead the snapshot is loaded from an on-disk cache (instant) and only fully (re)loaded from SP on
/// a manual delta load (<c>POST /api/sales-orders/import-history</c>). Between loads, freshness rides the
/// write-through (<see cref="SharePointService.SalesOrderAppended"/> as ERP SalesOrder docs arrive), persisted
/// on a debounce. A machine with no disk cache yet is seeded ONCE from SP in the background.</para>
/// </summary>
public sealed class SalesOrderHistoryService : IHostedService
{
    private const int SchemaVersion = 1;

    private readonly SharePointService _sp;
    private readonly ILogger<SalesOrderHistoryService> _log;
    private readonly DiskBackedCache<SalesOrderCacheFile> _disk;

    private volatile Snapshot _snap = Snapshot.Empty;
    private volatile bool _built;
    private DateTime? _lastLoadUtc;      // delta marker — when SP was last fully (re)loaded
    private volatile bool _dirty;        // write-through merges awaiting a disk persist
    private Timer? _persistTimer;

    private static readonly JsonSerializerOptions JsonOpts = new()
    {
        PropertyNameCaseInsensitive = true,
        WriteIndented               = false,
    };

    public SalesOrderHistoryService(SharePointService sp, ILogger<SalesOrderHistoryService> log)
    {
        _sp  = sp;
        _log = log;
        var dir = Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
            "ShredderData", "Cache", $"v{SchemaVersion}");
        _disk = new DiskBackedCache<SalesOrderCacheFile>(dir, "salesorders", log, JsonOpts);
        // Signal — local write-through: an order appended on THIS proxy (FileWatcher) merges into the serving
        // cache directly (a proxy filters out its own bus messages, so it must self-update).
        _sp.SalesOrderAppended += MergeOrder;
    }

    /// <summary>The delta marker: when SharePoint was last fully (re)loaded into this cache. Null when the
    /// cache has never been loaded (fresh machine, no manual load yet).</summary>
    public DateTime? LastLoadUtc => _lastLoadUtc;
    public int OrderCount => _snap.Count;

    public Task StartAsync(CancellationToken ct)
    {
        // Load the serving cache from DISK — instant — rather than the ~29s SharePoint read this used to do
        // on every startup. Keep the on-disk copy until the next manual load refreshes it.
        var file = _disk.TryLoad();
        if (file is { SchemaVersion: SchemaVersion, Rows.Count: > 0 })
        {
            _snap        = Build(file.Rows);
            _built       = true;
            _lastLoadUtc = file.LastLoadUtc;
            _log.LogInformation("[SOHistory] warmed from disk — {Orders} orders across {Customers} customers (last SP load {At:u})",
                _snap.Count, _snap.ByCustomer.Count, _lastLoadUtc);
        }
        else
        {
            // No usable disk cache (first deploy of this change, or wiped) — seed it ONCE from SP in the
            // background so the card still has data, without blocking startup or re-reading on later restarts.
            _ = Task.Run(async () =>
            {
                try
                {
                    await StartupTimings.MeasureAsync("sales-order-history", "seed", () => RefreshFromSpAsync(CancellationToken.None));
                    _log.LogInformation("[SOHistory] seeded disk cache from SharePoint (first run)");
                }
                catch (Exception ex) { _log.LogWarning(ex, "[SOHistory] initial seed failed — empty until next load"); }
            });
        }

        // Persist write-through merges to disk on a debounce (no SharePoint traffic; the ~29s SP re-read that
        // used to run here every 5 minutes is gone — a manual load is the only full refresh).
        _persistTimer = new Timer(_ => { if (_dirty) { _dirty = false; _ = PersistAsync(); } },
            null, TimeSpan.FromSeconds(30), TimeSpan.FromSeconds(30));
        return Task.CompletedTask;
    }

    public Task StopAsync(CancellationToken ct)
    {
        _sp.SalesOrderAppended -= MergeOrder;
        _persistTimer?.Dispose();
        return _dirty ? PersistAsync() : Task.CompletedTask;
    }

    // ── Serving ──────────────────────────────────────────────────────────────

    /// <summary>The customer's recent orders + summary header, newest-first, capped to <paramref name="top"/>.
    /// Served from the in-memory (disk-loaded) snapshot — empty until the cache is seeded / first loaded.</summary>
    public Task<CustomerOrdersResponse> GetOrdersForCustomerAsync(string customer, int top, CancellationToken ct = default)
    {
        _snap.ByCustomer.TryGetValue(NormName(customer), out var orders);
        return Task.FromResult(Project(orders ?? [], top));
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

    // ── Cache build / persistence ───────────────────────────────────────────────

    private static Snapshot Build(IReadOnlyList<SalesOrderRecord> rows)
    {
        var byCustomer = rows
            .GroupBy(r => NormName(r.CustomerName), StringComparer.OrdinalIgnoreCase)
            .ToDictionary(g => g.Key, g => (IReadOnlyList<SalesOrderRecord>)g.ToList(),
                StringComparer.OrdinalIgnoreCase);
        return new Snapshot(byCustomer, rows.Count);
    }

    /// <summary>Full SharePoint (re)load → serving snapshot → disk + marker. Used to seed a fresh machine; the
    /// manual delta load updates the snapshot from its merged set instead, avoiding a second SP read.</summary>
    private async Task RefreshFromSpAsync(CancellationToken ct = default)
    {
        var rows = await _sp.ReadAllSalesOrdersAsync(ct);
        _snap        = Build(rows);
        _built       = true;
        _lastLoadUtc = DateTime.UtcNow;
        await PersistAsync();
        _log.LogInformation("[SOHistory] cache (re)built from SP: {Orders} orders across {Customers} customers",
            _snap.Count, _snap.ByCustomer.Count);
    }

    private Task PersistAsync() => _disk.SaveAsync(new SalesOrderCacheFile
    {
        SchemaVersion = SchemaVersion,
        LastLoadUtc   = _lastLoadUtc,
        Rows          = _snap.AllRows(),
    });

    // ── Write-through merge (single order appended via an ERP doc) ──────────────

    private readonly object _mergeLock = new();

    /// <summary>Merges a single order into the serving cache, insert-if-absent by OrderId, with NO SharePoint
    /// read — O(1) per order. Marks the cache dirty so the debounce timer persists it to disk.</summary>
    public void MergeOrder(SalesOrderRecord rec)
    {
        if (rec is null || string.IsNullOrWhiteSpace(rec.OrderId) || string.IsNullOrWhiteSpace(rec.CustomerName))
            return;

        lock (_mergeLock)
        {
            if (!_built) return;   // cold — the next seed/load reads the persisted row from SP anyway
            var snap = _snap;
            var key  = NormName(rec.CustomerName);
            snap.ByCustomer.TryGetValue(key, out var existing);
            existing ??= [];
            if (existing.Any(o => string.Equals(o.OrderId, rec.OrderId, StringComparison.OrdinalIgnoreCase)))
                return;   // already present — insert-if-absent

            var newList = new List<SalesOrderRecord>(existing) { rec };
            var newDict = new Dictionary<string, IReadOnlyList<SalesOrderRecord>>(
                snap.ByCustomer, StringComparer.OrdinalIgnoreCase) { [key] = newList };
            _snap  = new Snapshot(newDict, snap.Count + 1);
            _dirty = true;
            _log.LogDebug("[SOHistory] merged order {Order} for {Customer} (cache now {N} orders)",
                rec.OrderId, rec.CustomerName, snap.Count + 1);
        }
    }

    // ── Delta load (manual, background, resumable) ──────────────────────────────

    private readonly object _importLock = new();
    private volatile SalesOrderImportStatus _import = SalesOrderImportStatus.Idle;

    public SalesOrderImportStatus ImportStatus => _import;

    /// <summary>Kicks off a background DELTA load of the parsed export: only new + changed rows are written.
    /// Returns false (with a message) when a load is already running. The serving cache + disk are refreshed
    /// from the merged set when the load completes (no extra SP read), and the delta marker is advanced.</summary>
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
        message = $"Started background delta load of {rows.Count} parsed row(s).";
        _ = Task.Run(() => RunImportAsync(rows));
        return true;
    }

    private async Task RunImportAsync(IReadOnlyList<CustomerImportService.SalesOrderRow> rows)
    {
        try
        {
            var result = await _sp.UpsertSalesOrdersAsync(
                rows, onProgress: n => { _import = _import with { Added = n }; }, CancellationToken.None);

            // Publish the merged set to the serving cache + disk, and advance the delta marker — no second SP read.
            _snap        = Build(result.Merged);
            _built       = true;
            _lastLoadUtc = DateTime.UtcNow;
            _dirty       = false;
            await PersistAsync();

            _import = _import with
            {
                Running = false, Added = result.Added, Changed = result.Changed,
                Unchanged = result.Unchanged, Failed = result.Failed,
                Message = $"Done: {result.Added} new, {result.Changed} changed, {result.Unchanged} unchanged, {result.Failed} failed.",
            };
        }
        catch (Exception ex)
        {
            _log.LogError(ex, "[SOHistory] delta load failed");
            _import = _import with { Running = false, Message = $"Failed: {ex.Message}" };
        }
    }

    // ── Helpers ────────────────────────────────────────────────────────────────

    // Collapse-whitespace + lowercase — matches SharePointService.NormName / the contact-import dedup lesson.
    private static string NormName(string s) =>
        string.Join(' ', (s ?? "").Split((char[]?)null, StringSplitOptions.RemoveEmptyEntries)).ToLowerInvariant();

    private sealed record Snapshot(
        Dictionary<string, IReadOnlyList<SalesOrderRecord>> ByCustomer, int Count)
    {
        public static Snapshot Empty { get; } = new(new(StringComparer.OrdinalIgnoreCase), 0);

        /// <summary>All rows flattened, for persisting the disk cache.</summary>
        public List<SalesOrderRecord> AllRows() => ByCustomer.Values.SelectMany(v => v).ToList();
    }
}

/// <summary>Live status of the background delta load, surfaced by the import status endpoint.</summary>
public sealed record SalesOrderImportStatus(
    bool Running, int Total, int Added, int Changed, int Unchanged, int Failed, string Message)
{
    public static SalesOrderImportStatus Idle { get; } = new(false, 0, 0, 0, 0, 0, "idle");
    public static SalesOrderImportStatus Started(int total) => new(true, total, 0, 0, 0, 0, "running");
}
