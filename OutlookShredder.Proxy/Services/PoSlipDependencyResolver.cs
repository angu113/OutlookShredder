using OutlookShredder.Proxy.Models;

namespace OutlookShredder.Proxy.Services;

/// <summary>
/// Pure dependency rules between purchase orders and the sales-order (picking-slip) cards they fulfil.
///
/// A slip card depends on a PO when the slip's HSK-SO (its DocumentNumber) is one of the PO's captured
/// SalesOrders. The slip's board date is PINNED to the PO's receipt date — BoardDate override, else the
/// supplier ETA (ExpectedDate). When material is split across several POs the slip is gated by the LATEST
/// of their receipt dates (all material must land first). A received PO (MaterialReceivedAt set) no longer
/// governs — the slip is released. Pure + order-independent: it reconciles whatever POs/slips exist now,
/// so it doesn't matter whether the PO or the slip was created first.
/// </summary>
public static class PoSlipDependency
{
    /// <summary>The PO's normalized receipt date ("yyyy-MM-dd"), or null when it doesn't govern
    /// (received, or unscheduled with no BoardDate/ETA).</summary>
    public static string? ReceiptDate(PurchaseOrderRecord po)
    {
        if (!string.IsNullOrWhiteSpace(po.MaterialReceivedAt)) return null;   // received → released
        var raw = !string.IsNullOrWhiteSpace(po.BoardDate)     ? po.BoardDate
                : !string.IsNullOrWhiteSpace(po.ExpectedDate)  ? po.ExpectedDate
                : null;
        return DateTimeOffset.TryParse(raw, out var d) ? d.ToString("yyyy-MM-dd") : null;
    }

    /// <summary>Canonical HSK-SO tokens this PO fulfils, from its SalesOrders CSV.</summary>
    public static IEnumerable<string> SalesOrdersOf(PurchaseOrderRecord po) =>
        (po.SalesOrders ?? "").Split([',', ';'],
            StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries);

    /// <summary>Map of HSK-SO → latest governing PO receipt date across all active POs.</summary>
    public static Dictionary<string, string> LatestReceiptBySalesOrder(IEnumerable<PurchaseOrderRecord> pos)
    {
        var map = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        foreach (var po in pos)
        {
            var date = ReceiptDate(po);
            if (date is null) continue;
            foreach (var so in SalesOrdersOf(po))
                if (!map.TryGetValue(so, out var cur) || string.CompareOrdinal(date, cur) > 0)
                    map[so] = date;   // "yyyy-MM-dd" lexical order == chronological
        }
        return map;
    }

    /// <summary>The date a slip with the given HSK-SO must be pinned to, or null if it's free
    /// (no governing active PO). Used to enforce/snap a manual slip move.</summary>
    public static string? PinnedDateFor(string? slipSalesOrder, IEnumerable<PurchaseOrderRecord> pos)
    {
        if (string.IsNullOrWhiteSpace(slipSalesOrder)) return null;
        return LatestReceiptBySalesOrder(pos).TryGetValue(slipSalesOrder.Trim(), out var d) ? d : null;
    }

    /// <summary>
    /// Slips whose board date is out of sync with their governing PO and the date they must move to.
    /// Only SCHEDULED slips (AssignedDate != "") are pinned — an unscheduled (Prioritize) slip is left
    /// for the user to place, at which point the move is enforced. Completed slips are ignored.
    /// </summary>
    public static IReadOnlyList<(WorkflowCard Slip, string PinnedDate)> ComputePins(
        IEnumerable<PurchaseOrderRecord> pos, IEnumerable<WorkflowCard> slips)
    {
        var latest = LatestReceiptBySalesOrder(pos);
        var result = new List<(WorkflowCard, string)>();
        foreach (var slip in slips)
        {
            if (slip.IsCompleted) continue;
            var so = (slip.DocumentNumber ?? "").Trim();
            if (so.Length == 0) continue;
            if (string.IsNullOrEmpty(slip.AssignedDate)) continue;          // unscheduled — leave in Prioritize
            if (!latest.TryGetValue(so, out var pin)) continue;             // not governed by any active PO
            if (!string.Equals(slip.AssignedDate, pin, StringComparison.Ordinal))
                result.Add((slip, pin));
        }
        return result;
    }
}

/// <summary>
/// Enforces <see cref="PoSlipDependency"/> across the board: pins dependent slip cards to their governing
/// PO's receipt date. Runs on a timer (the order-independence guard — catches a PO or slip whose
/// counterpart is added later) and is also invoked immediately when a PO moves/receives or a slip moves.
/// Slip updates flow through <see cref="WorkflowCardService"/> so the cache stays warm and clients get a
/// live "Updated" bus event.
/// </summary>
public class PoSlipDependencyResolver : IHostedService
{
    private readonly SharePointService                   _sp;
    private readonly WorkflowCardService                 _wf;
    private readonly ILogger<PoSlipDependencyResolver>   _log;
    private readonly SemaphoreSlim                       _gate = new(1, 1);

    // POs from the last reconcile — used by PinnedDateForSlipAsync so a slip-move check doesn't hit Graph
    // on every drag. At most one reconcile-interval stale; the timer corrects any drift.
    private volatile IReadOnlyList<PurchaseOrderRecord> _posCache = [];
    private Timer? _timer;

    public PoSlipDependencyResolver(SharePointService sp, WorkflowCardService wf,
        ILogger<PoSlipDependencyResolver> log)
    {
        _sp = sp; _wf = wf; _log = log;
    }

    public Task StartAsync(CancellationToken ct)
    {
        _timer = new Timer(_ => _ = ReconcileAsync(), null,
            TimeSpan.FromSeconds(20), TimeSpan.FromSeconds(30));
        return Task.CompletedTask;
    }

    public Task StopAsync(CancellationToken ct) { _timer?.Dispose(); return Task.CompletedTask; }

    /// <summary>Pin every scheduled slip to its governing PO's (latest) receipt date. Idempotent; safe to
    /// call often. Overlapping calls are coalesced (a running reconcile already covers the latest state).</summary>
    public async Task ReconcileAsync(CancellationToken ct = default)
    {
        if (!await _gate.WaitAsync(0, ct)) return;   // a reconcile is already in flight
        try
        {
            var pos   = await _sp.ReadPurchaseOrdersAsync();
            _posCache = pos;
            var slips = await _wf.GetAllAsync();
            var pins  = PoSlipDependency.ComputePins(pos, slips);
            foreach (var (slip, pin) in pins)
            {
                await _wf.UpdateAsync(slip.SpItemId,
                    new UpdateWorkflowCardRequest { AssignedDate = pin }, ct);
                _log.LogInformation("[PO-DEP] Pinned slip {Doc} ({Id}): {Old} -> {New} (PO receipt date)",
                    slip.DocumentNumber, slip.SpItemId,
                    string.IsNullOrEmpty(slip.AssignedDate) ? "(none)" : slip.AssignedDate, pin);
            }
            if (pins.Count > 0)
                _log.LogInformation("[PO-DEP] Reconcile pinned {N} slip(s) to their PO receipt date", pins.Count);
        }
        catch (Exception ex) { _log.LogWarning(ex, "[PO-DEP] Reconcile failed"); }
        finally { _gate.Release(); }
    }

    /// <summary>The date a slip must sit on given its HSK-SO, or null if free (no governing active PO).
    /// Uses the PO list cached by the last reconcile to avoid a Graph round-trip per slip move.</summary>
    public async Task<string?> PinnedDateForSlipAsync(string? slipSalesOrder)
    {
        var pos = _posCache;
        if (pos.Count == 0) { pos = await _sp.ReadPurchaseOrdersAsync(); _posCache = pos; }
        return PoSlipDependency.PinnedDateFor(slipSalesOrder, pos);
    }
}
