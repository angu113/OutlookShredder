using System;
using System.Collections.Generic;
using System.Threading;
using System.Threading.Tasks;
using Microsoft.Extensions.Hosting;
using Microsoft.Extensions.Logging;

namespace OutlookShredder.Proxy.Services;

/// <summary>
/// Hosted, in-memory index of inbound supplier-message rows — the minimal projection
/// <see cref="SupplierReadModel.Tally"/> consumes (Rfq, Supplier, MessageId, MsgTime), built from BOTH
/// SupplierResponses (all inbound) and SupplierConversations (Direction=="in") via
/// <see cref="SharePointService.ScanInboundRowsAsync"/>.
///
/// Why: the per-user unread tally was re-scanning both SharePoint lists on EVERY call (~5–7 s, growing with
/// volume). The row set is identical for every user — only the per-user profile (watermark + read/unread
/// sets) differs, and that is applied in <c>Tally</c> at read time. So one shared index turns every per-user
/// tally into an in-memory pass (sub-100 ms) with zero SP round-trips on the hot path.
///
/// Freshness: built once at startup, then refreshed (a) on a 5-min safety timer, (b) write-through when THIS
/// proxy writes a new inbound row (<see cref="SharePointService.InboundChanged"/>, trailing-debounced so a
/// burst collapses to one rebuild), and (c) cross-proxy when a peer publishes an SR event on the rfq-updates
/// bus (wired in <c>RfqNotificationService.OnProxyBusMessageAsync</c>). The comms-start cutoff is NOT applied
/// here — <c>Tally</c> floors at read time, so a CommsDataStartDate change needs no rebuild.
///
/// Cold start: <see cref="TryGetRows"/> returns null until the first build completes, so callers fall back to
/// a live scan (one slow call) rather than flashing an empty badge.
/// </summary>
public sealed class SupplierUnreadIndexService : BackgroundService, ICacheStatusProvider
{
    private readonly SharePointService _sp;
    private readonly ILogger<SupplierUnreadIndexService> _log;

    private volatile IReadOnlyList<SupplierReadModel.ReadRow> _rows = Array.Empty<SupplierReadModel.ReadRow>();
    private volatile bool _built;
    private DateTime? _builtUtc;
    private DateTime? _lastRefreshUtc;
    private readonly SemaphoreSlim _gate = new(1, 1);

    private readonly object _debounceLock = new();
    private Timer? _debounce;

    private static readonly TimeSpan SafetyInterval = TimeSpan.FromMinutes(5);
    private static readonly TimeSpan DebounceDelay   = TimeSpan.FromSeconds(2);

    public SupplierUnreadIndexService(SharePointService sp, ILogger<SupplierUnreadIndexService> log)
    {
        _sp  = sp;
        _log = log;
        // Local write-through: any inbound SR/conversation change re-builds the index (trailing-debounced).
        _sp.InboundChanged += RequestRefresh;
    }

    /// <summary>The current inbound-row snapshot, or null when the index has not built yet — the caller then
    /// falls back to a live scan rather than reporting an empty/incorrect tally.</summary>
    public IReadOnlyList<SupplierReadModel.ReadRow>? TryGetRows() => _built ? _rows : null;

    /// <summary>Rebuilds the index from a single full scan of both inbound sources. Serialized via a gate so
    /// concurrent triggers (timer + write-through + bus) don't double-fetch.</summary>
    public async Task RefreshAsync(CancellationToken ct = default)
    {
        await _gate.WaitAsync(ct);
        try
        {
            var sw   = System.Diagnostics.Stopwatch.StartNew();
            var rows = await _sp.ScanInboundRowsAsync(ct);
            _rows           = rows;
            _built          = true;
            _builtUtc     ??= DateTime.UtcNow;
            _lastRefreshUtc = DateTime.UtcNow;
            _log.LogInformation("[UnreadIndex] rebuilt {Count} inbound rows in {Ms}ms", rows.Count, sw.ElapsedMilliseconds);
        }
        finally { _gate.Release(); }
    }

    /// <summary>Fire-and-forget, trailing-debounced rebuild for write-through + bus hooks: never blocks the
    /// caller, and a burst of triggers (e.g. a bulk import or reprocess) collapses to a single rebuild fired
    /// once the burst quiets. The 5-min safety timer covers a burst that never pauses.</summary>
    public void RequestRefresh()
    {
        lock (_debounceLock)
        {
            _debounce ??= new Timer(_ => _ = SafeRefreshAsync(), null, Timeout.InfiniteTimeSpan, Timeout.InfiniteTimeSpan);
            _debounce.Change(DebounceDelay, Timeout.InfiniteTimeSpan);   // re-arm on each trigger (trailing)
        }
    }

    private async Task SafeRefreshAsync()
    {
        try { await RefreshAsync(); }
        catch (Exception ex) { _log.LogWarning(ex, "[UnreadIndex] debounced refresh failed"); }
    }

    protected override async Task ExecuteAsync(CancellationToken stoppingToken)
    {
        // Build once at startup, then a 5-min safety re-scan (write-through + bus keep it fresh between).
        while (!stoppingToken.IsCancellationRequested)
        {
            try { await RefreshAsync(stoppingToken); }
            catch (OperationCanceledException) { break; }
            catch (Exception ex) { _log.LogWarning(ex, "[UnreadIndex] scheduled refresh failed"); }
            try { await Task.Delay(SafetyInterval, stoppingToken); }
            catch (OperationCanceledException) { break; }
        }
    }

    public override void Dispose()
    {
        _sp.InboundChanged -= RequestRefresh;
        _debounce?.Dispose();
        _gate.Dispose();
        base.Dispose();
    }

    // ── ICacheStatusProvider (Home health dashboard) ───────────────────────────
    public string    Name          => "unread-index";
    public string    DisplayName   => "Unread Index";
    public int       SchemaVersion => 1;
    public int       ItemCount     => _rows.Count;
    public DateTime? CacheBuiltUtc => _builtUtc;
    public DateTime? LastDeltaUtc  => _lastRefreshUtc;
    public bool      IsLoading     => _gate.CurrentCount == 0;
    public Task ForceRebuildAsync(CancellationToken ct = default) => RefreshAsync(ct);
    public Task ForceDeltaAsync(CancellationToken ct = default)   => RefreshAsync(ct);
}
