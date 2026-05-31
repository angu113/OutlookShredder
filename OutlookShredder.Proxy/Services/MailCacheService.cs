using System.Collections.Concurrent;
using System.Text.Json;
using OutlookShredder.Proxy.Models;

namespace OutlookShredder.Proxy.Services;

/// <summary>
/// In-memory read model for the mail workbench, mirroring <see cref="ArchiveCacheService"/>:
///   • L1 (memory): every captured <see cref="MailItemRow"/> + its current <see cref="MailClassRow"/>,
///     constrained to messages received on/after <c>CommsDataStartDate</c> (ShredderConfig SP list,
///     the mail analogue of RfqDataStartDate).
///   • L2 (disk): a JSON snapshot at %LOCALAPPDATA%\ShredderData\Cache\v1\mail.json, loaded instantly
///     at startup so the tree/grid serve from memory before the first SP read returns.
///
/// The workbench serves <c>GetTreeAsync</c>/<c>GetItemsAsync</c> from this cache (no per-request SP
/// roundtrip) and upserts/removes here on every write. A periodic full reconcile catches anything
/// missed (e.g. a peer proxy's writes not delivered via bus). Cross-proxy coherency: peer "Mail"
/// bus events are applied via <see cref="ApplyBusItem"/>/<see cref="ApplyBusDelete"/>.
/// </summary>
public sealed class MailCacheService : IHostedService, ICacheStatusProvider
{
    private readonly SharePointService _sp;
    private readonly IConfiguration _config;
    private readonly ILogger<MailCacheService> _log;
    private readonly DiskBackedCache<MailCacheFile> _disk;
    private readonly SemaphoreSlim _refreshGate = new(1, 1);

    // L1: keyed by MailItemId.
    private readonly ConcurrentDictionary<string, MailItemRow>  _items = new(StringComparer.Ordinal);
    private readonly ConcurrentDictionary<string, MailClassRow> _class = new(StringComparer.Ordinal);

    private Timer? _timer;
    private volatile bool _warmed;
    private volatile bool _loading;
    private DateTimeOffset _startCutoff = DateTimeOffset.MinValue;
    public DateTimeOffset? LastRefreshAt { get; private set; }
    public bool Warmed => _warmed;

    // ── ICacheStatusProvider (surfaces in Tools | System | Data Cache) ──────────────
    public string Name        => "mail";
    public string DisplayName => "Mail Classify";
    public int    SchemaVersion => 1;
    public int    ItemCount   => _items.Count;
    public DateTime? CacheBuiltUtc => LastRefreshAt?.UtcDateTime;
    public DateTime? LastDeltaUtc  => LastRefreshAt?.UtcDateTime;
    public bool   IsLoading   => _loading;
    public Task ForceRebuildAsync(CancellationToken ct = default) => ForceRefreshAsync(ct);
    public Task ForceDeltaAsync(CancellationToken ct = default)   => ForceRefreshAsync(ct);

    // How far back to cache when CommsDataStartDate is unset.
    private const int DefaultLookbackDays = 90;
    // Full reconcile cadence.
    private static readonly TimeSpan ReconcileInterval = TimeSpan.FromMinutes(10);

    private static readonly JsonSerializerOptions _diskOpts = new()
    {
        PropertyNamingPolicy = JsonNamingPolicy.CamelCase,
        DefaultIgnoreCondition = System.Text.Json.Serialization.JsonIgnoreCondition.WhenWritingNull,
    };

    public MailCacheService(SharePointService sp, IConfiguration config, ILogger<MailCacheService> log)
    {
        _sp = sp; _config = config; _log = log;
        var dir = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
            "ShredderData", "Cache", "v1");
        _disk = new DiskBackedCache<MailCacheFile>(dir, "mail", log, _diskOpts);
    }

    public Task StartAsync(CancellationToken ct)
    {
        // Serve the disk snapshot immediately (instant cold-start), then warm from SP in the background.
        try
        {
            var snap = _disk.TryLoad();
            if (snap is not null)
            {
                foreach (var i in snap.Items ?? []) if (i.MailItemId.Length > 0) _items[i.MailItemId] = i;
                foreach (var c in snap.Class ?? []) if (c.MailItemId.Length > 0) _class[c.MailItemId] = c;
                _warmed = _items.Count > 0;
                _log.LogInformation("[MailCache] loaded disk snapshot: {Items} items, {Class} classifications", _items.Count, _class.Count);
            }
        }
        catch (Exception ex) { _log.LogWarning(ex, "[MailCache] disk snapshot load failed"); }

        // Delay the first SP refresh ~8s so startup isn't gated on it.
        _timer = new Timer(async _ => await SafeRefreshAsync(), null, TimeSpan.FromSeconds(8), ReconcileInterval);
        return Task.CompletedTask;
    }

    public Task StopAsync(CancellationToken ct) { _timer?.Dispose(); return Task.CompletedTask; }

    private async Task SafeRefreshAsync()
    {
        try { await ForceRefreshAsync(); }
        catch (Exception ex) { _log.LogWarning(ex, "[MailCache] scheduled refresh failed"); }
    }

    /// <summary>Full reconcile from SP, constrained by CommsDataStartDate. Swaps the in-memory set atomically.</summary>
    public async Task ForceRefreshAsync(CancellationToken ct = default)
    {
        await _refreshGate.WaitAsync(ct);
        _loading = true;
        try
        {
            _startCutoff = await ResolveStartCutoffAsync();
            var items = await _sp.ReadMailItemsAsync(ct);
            var cls   = await _sp.ReadCurrentClassificationsAsync(ct);

            var kept = items.Where(i => i.MailItemId.Length > 0 && WithinCutoff(i.ReceivedAt))
                            .ToDictionary(i => i.MailItemId, i => i, StringComparer.Ordinal);
            var keptCls = cls.Where(c => kept.ContainsKey(c.MailItemId))
                             .ToDictionary(c => c.MailItemId, c => c, StringComparer.Ordinal);

            // Atomic-ish swap: clear then repopulate under the gate (readers tolerate a brief gap).
            _items.Clear(); foreach (var kv in kept)    _items[kv.Key] = kv.Value;
            _class.Clear(); foreach (var kv in keptCls) _class[kv.Key] = kv.Value;
            _warmed = true;
            LastRefreshAt = DateTimeOffset.UtcNow;
            _log.LogInformation("[MailCache] reconciled from SP: {Items} items, {Class} classifications (cutoff {Cutoff:yyyy-MM-dd})",
                _items.Count, _class.Count, _startCutoff);
            await PersistAsync();
        }
        finally { _loading = false; _refreshGate.Release(); }
    }

    private bool WithinCutoff(string receivedIso)
    {
        if (_startCutoff <= DateTimeOffset.MinValue) return true;
        return !DateTimeOffset.TryParse(receivedIso, out var d) || d >= _startCutoff;
    }

    private async Task<DateTimeOffset> ResolveStartCutoffAsync()
    {
        try
        {
            var cfg = await _sp.GetShredderConfigAsync("CommsDataStartDate");
            if (cfg.HasValue && DateTimeOffset.TryParse(cfg.Value.Value, out var d)) return d;
        }
        catch (Exception ex) { _log.LogDebug(ex, "[MailCache] CommsDataStartDate read failed — using default lookback"); }
        return DateTimeOffset.UtcNow.AddDays(-DefaultLookbackDays);
    }

    // ── Reads (served to workbench) ──────────────────────────────────────────────

    public IReadOnlyCollection<MailItemRow>  GetItems()   => _items.Values.ToList();
    public IReadOnlyCollection<MailClassRow> GetCurrents() => _class.Values.ToList();
    public bool TryGetItem(string mailItemId, out MailItemRow row) => _items.TryGetValue(mailItemId, out row!);
    public MailClassRow? GetClass(string mailItemId) => _class.TryGetValue(mailItemId, out var c) ? c : null;

    // ── Writes (local, in-process) ───────────────────────────────────────────────

    public void UpsertItem(MailItemRow row)
    {
        if (row.MailItemId.Length == 0) return;
        if (!WithinCutoff(row.ReceivedAt)) return;   // outside the cached window — ignore
        _items[row.MailItemId] = row;
        _ = PersistAsync();
    }

    public void UpsertClass(MailClassRow row)
    {
        if (row.MailItemId.Length == 0) return;
        if (!_items.ContainsKey(row.MailItemId)) return;  // only track classes for cached items
        _class[row.MailItemId] = row;
        _ = PersistAsync();
    }

    public void SetCompleted(string mailItemId, bool completed, string? completedAtIso, string? by = null)
    {
        if (_items.TryGetValue(mailItemId, out var row))
        {
            row.Completed = completed;
            row.CompletedAt = completed ? (completedAtIso ?? DateTimeOffset.UtcNow.ToString("o")) : null;
            row.CompletedBy = completed ? by : null;
            _ = PersistAsync();
        }
    }

    public void SetRead(string mailItemId, bool read, string? by, string? readAtIso = null)
    {
        if (_items.TryGetValue(mailItemId, out var row))
        {
            row.IsRead = read;
            row.ReadAt = read ? (readAtIso ?? DateTimeOffset.UtcNow.ToString("o")) : null;
            row.ReadBy = read ? by : null;
            _ = PersistAsync();
        }
    }

    public void SetReceived(string mailItemId, string receivedIso)
    {
        if (_items.TryGetValue(mailItemId, out var row)) { row.ReceivedAt = receivedIso; _ = PersistAsync(); }
    }

    public void SetConversation(string mailItemId, string conversationId)
    {
        if (_items.TryGetValue(mailItemId, out var row)) { row.ConversationId = conversationId; _ = PersistAsync(); }
    }

    public void SetClaim(string mailItemId, string? claimedBy, string? claimedAtIso)
    {
        if (_items.TryGetValue(mailItemId, out var row))
        {
            row.ClaimedBy = claimedBy;
            row.ClaimedAt = string.IsNullOrEmpty(claimedBy) ? null : (claimedAtIso ?? DateTimeOffset.UtcNow.ToString("o"));
            _ = PersistAsync();
        }
    }

    public void Remove(string mailItemId)
    {
        _items.TryRemove(mailItemId, out _);
        _class.TryRemove(mailItemId, out _);
        _ = PersistAsync();
    }

    /// <summary>Wipes the cache (dev purge). Peers reconcile on their next scheduled refresh.</summary>
    public void ClearAll()
    {
        _items.Clear(); _class.Clear();
        _ = PersistAsync();
    }

    // ── Cross-proxy apply (from peer "Mail" bus events) ──────────────────────────

    public void ApplyBusItem(MailBusItem b)
    {
        if (string.IsNullOrEmpty(b.MailItemId)) return;
        UpsertItem(new MailItemRow
        {
            SpId = b.SpId, MailItemId = b.MailItemId, WrapperGraphId = b.WrapperGraphId, ConversationId = b.ConversationId,
            SourceType = b.SourceType, SourceMailbox = b.SourceMailbox,
            FromAddress = b.FromAddress, FromName = b.FromName, Subject = b.Subject,
            ReceivedAt = b.ReceivedAt, HasAttachments = b.HasAttachments,
            AttachmentsJson = b.AttachmentsJson, Completed = b.Completed, CompletedAt = b.CompletedAt,
            CompletedBy = b.CompletedBy, IsRead = b.IsRead, ReadAt = b.ReadAt, ReadBy = b.ReadBy,
            ClaimedBy = b.ClaimedBy, ClaimedAt = b.ClaimedAt,
        });
        UpsertClass(new MailClassRow
        {
            MailItemId = b.MailItemId, IsCurrent = true, CategoryPath = b.CategoryPath,
            OtherLabel = b.OtherLabel, Confidence = b.Confidence, KeywordTags = b.KeywordTags,
            PoNumber = b.PoNumber, SoNumber = b.SoNumber, Amount = b.Amount,
        });
    }

    public void ApplyBusDelete(string mailItemId) => Remove(mailItemId);

    /// <summary>Builds a compact bus snapshot for an item (merges its current classification).</summary>
    public MailBusItem ToBusItem(MailItemRow row, MailClassRow? cls)
    {
        cls ??= GetClass(row.MailItemId);
        return new MailBusItem
        {
            MailItemId = row.MailItemId, SpId = row.SpId, WrapperGraphId = row.WrapperGraphId, ConversationId = row.ConversationId,
            SourceType = row.SourceType, SourceMailbox = row.SourceMailbox,
            FromAddress = row.FromAddress, FromName = row.FromName, Subject = row.Subject,
            ReceivedAt = row.ReceivedAt, HasAttachments = row.HasAttachments,
            AttachmentsJson = row.AttachmentsJson, Completed = row.Completed, CompletedAt = row.CompletedAt,
            CompletedBy = row.CompletedBy, IsRead = row.IsRead, ReadAt = row.ReadAt, ReadBy = row.ReadBy,
            ClaimedBy = row.ClaimedBy, ClaimedAt = row.ClaimedAt,
            CategoryPath = cls?.CategoryPath ?? "Other", OtherLabel = cls?.OtherLabel,
            Confidence = cls?.Confidence ?? 0, KeywordTags = cls?.KeywordTags ?? "",
            PoNumber = cls?.PoNumber, SoNumber = cls?.SoNumber, Amount = cls?.Amount,
        };
    }

    private Task PersistAsync() =>
        _disk.SaveAsync(new MailCacheFile
        {
            CacheBuiltUtc = DateTimeOffset.UtcNow,
            Items = _items.Values.ToList(),
            Class = _class.Values.ToList(),
        });

    public sealed class MailCacheFile
    {
        public DateTimeOffset? CacheBuiltUtc { get; set; }
        public List<MailItemRow>  Items { get; set; } = [];
        public List<MailClassRow> Class { get; set; } = [];
    }
}
