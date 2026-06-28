using System.Collections.Concurrent;
using System.Text.Json;
using Microsoft.Extensions.Hosting;
using OutlookShredder.Proxy.Models;
using OutlookShredder.Proxy.Services.Storage;

namespace OutlookShredder.Proxy.Services;

/// <summary>
/// L1 (memory) + L2 (disk) cache of every ACTIVE inquiry (anything not Closed/Spam) and its full conversation,
/// so Pulse renders without a per-open SharePoint round-trip. Mirrors <see cref="MailCacheService"/>:
///   • L1: <see cref="InquiryCacheEntry"/> per CINQ id (inquiry header + messages + notes + quotations + drafts
///     + contact). The CRM card is recomputed on read so it stays fresh as the catalog changes.
///   • L2: a JSON snapshot at %LOCALAPPDATA%\ShredderData\Cache\v1\inquiries.json for instant cold-start.
///   • Attachments are pre-fetched to %LOCALAPPDATA%\ShredderData\Cache\v1\inquiry-media\{id}\ so the media
///     endpoint serves from local disk, not Graph.
/// Freshness is event-driven (write → SP → bus): the local proxy applies each mutation in place (no SP read on
/// the write path); a peer proxy's bus event triggers a targeted SP re-read (<see cref="RefreshOneAsync"/>).
/// Closed/Spam are never cached — they fall through to a direct SP read (the "load on demand" path).
/// </summary>
public sealed class SmsInquiryCacheService : IHostedService, ICacheStatusProvider
{
    private readonly IInquiryStore       _store;
    private readonly IMessageStore       _messages;
    private readonly CustomerCacheService _crm;
    private readonly IConfiguration      _config;
    private readonly ILogger<SmsInquiryCacheService> _log;
    private readonly DiskBackedCache<InquiryCacheFile> _disk;
    private readonly string _mediaCacheDir;
    private readonly SemaphoreSlim _gate = new(1, 1);

    // L1, keyed by CINQ id. Each entry's lists are mutated/copied under lock(entry).
    private readonly ConcurrentDictionary<string, InquiryCacheEntry> _entries = new(StringComparer.Ordinal);

    private Timer? _warmTimer;
    private Timer? _persistTimer;
    private volatile bool _warmed;
    private volatile bool _loading;
    private volatile bool _dirty;
    private DateTimeOffset? _lastRefreshAt;

    private static readonly JsonSerializerOptions _diskOpts = new() { PropertyNameCaseInsensitive = true };

    public SmsInquiryCacheService(IInquiryStore store, IMessageStore messages, CustomerCacheService crm,
        IConfiguration config, ILogger<SmsInquiryCacheService> log)
    {
        _store    = store;
        _messages = messages;
        _crm      = crm;
        _config   = config;
        _log      = log;

        var cacheDir = Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
            "ShredderData", "Cache", "v1");
        _disk          = new DiskBackedCache<InquiryCacheFile>(cacheDir, "inquiries", log, _diskOpts);
        _mediaCacheDir = Path.Combine(cacheDir, "inquiry-media");
    }

    // ── ICacheStatusProvider (surfaces in Tools | System | Data Cache) ──────────────
    public string    Name          => "inquiries";
    public string    DisplayName   => "SMS Inquiries";
    public int       SchemaVersion => 1;
    public int       ItemCount     => _entries.Count;
    public DateTime? CacheBuiltUtc => _lastRefreshAt?.UtcDateTime;
    public DateTime? LastDeltaUtc  => _lastRefreshAt?.UtcDateTime;
    public bool      IsLoading     => _loading;
    public Task ForceRebuildAsync(CancellationToken ct = default) => WarmAllAsync(ct);
    public Task ForceDeltaAsync(CancellationToken ct = default)   => WarmAllAsync(ct);

    public bool Warmed => _warmed;

    public Task StartAsync(CancellationToken ct)
    {
        // Serve the disk snapshot immediately (instant cold-start), then warm from SP in the background.
        try
        {
            var snap = _disk.TryLoad();
            if (snap?.Entries is { Count: > 0 } entries)
            {
                foreach (var e in entries)
                    if (!string.IsNullOrEmpty(e.Inquiry.Id)) _entries[e.Inquiry.Id] = e;
                _warmed = _entries.Count > 0;
                _lastRefreshAt = snap.CacheBuiltUtc;
                _log.LogInformation("[InquiryCache] loaded disk snapshot: {N} active inquiries", _entries.Count);
            }
        }
        catch (Exception ex) { _log.LogWarning(ex, "[InquiryCache] disk snapshot load failed"); }

        _warmTimer    = new Timer(async _ => await SafeWarmAsync(), null, TimeSpan.FromSeconds(10), TimeSpan.FromMinutes(10));
        _persistTimer = new Timer(_ => { if (_dirty) { _dirty = false; _ = _disk.SaveAsync(Snapshot()); } },
            null, TimeSpan.FromSeconds(30), TimeSpan.FromSeconds(30));
        return Task.CompletedTask;
    }

    public Task StopAsync(CancellationToken ct)
    {
        _warmTimer?.Dispose();
        _persistTimer?.Dispose();
        return _dirty ? _disk.SaveAsync(Snapshot()) : Task.CompletedTask;
    }

    private async Task SafeWarmAsync()
    {
        try { await WarmAllAsync(); }
        catch (Exception ex) { _log.LogWarning(ex, "[InquiryCache] background warm failed"); }
    }

    /// <summary>Full reconcile from SharePoint: load every active inquiry's detail, swap the cache, persist,
    /// and pre-fetch attachments. Reconciles away anything that became Closed/Spam or was deleted out-of-band.</summary>
    public async Task WarmAllAsync(CancellationToken ct = default)
    {
        await _gate.WaitAsync(ct);
        _loading = true;
        try
        {
            var all    = await _store.GetInquiriesAsync(null, null, ct);
            var active = all.Where(i => IsActive(i.Status)).ToList();

            var fresh = new Dictionary<string, InquiryCacheEntry>(StringComparer.Ordinal);
            foreach (var inq in active)
            {
                try { fresh[inq.Id] = await BuildEntryAsync(inq, ct); }
                catch (Exception ex) { _log.LogWarning(ex, "[InquiryCache] build {Id} failed", inq.Id); }
            }

            _entries.Clear();
            foreach (var kv in fresh) _entries[kv.Key] = kv.Value;
            _warmed = true;
            _lastRefreshAt = DateTimeOffset.UtcNow;
            _log.LogInformation("[InquiryCache] warmed {N} active inquiries from SP", _entries.Count);
            await _disk.SaveAsync(Snapshot());
            _dirty = false;
            _ = Task.Run(() => PrefetchAllMediaAsync(CancellationToken.None));
        }
        finally { _loading = false; _gate.Release(); }
    }

    // ── Reads (the fast path) ───────────────────────────────────────────────────────
    public async Task<InquiryDetail?> GetDetailAsync(string inquiryId, CancellationToken ct = default)
    {
        if (_entries.TryGetValue(inquiryId, out var hit)) return ToDetail(hit);

        // Miss: Closed/Spam (never cached) or not yet warmed → read straight from SP. Cache if it's active.
        var inquiry = await _store.GetInquiryByIdAsync(inquiryId, ct);
        if (inquiry is null) return null;
        var entry = await BuildEntryAsync(inquiry, ct);
        if (IsActive(inquiry.Status)) { _entries[inquiryId] = entry; MarkDirty(); _ = PrefetchEntryMediaAsync(entry, ct); }
        return ToDetail(entry);
    }

    public async Task<IReadOnlyList<Inquiry>> ListAsync(string? status, string? query, CancellationToken ct = default)
    {
        // The active set is cached; Closed/Spam (and any explicit closed query) fall through to SP.
        if (_warmed && status is not (InquiryStatus.Closed or InquiryStatus.Spam))
        {
            IEnumerable<Inquiry> q = Snapshot().Entries.Select(e => e.Inquiry);
            if (status is InquiryStatus.Open or InquiryStatus.Quoted)
                q = q.Where(i => string.Equals(i.Status, status, StringComparison.OrdinalIgnoreCase));
            if (!string.IsNullOrWhiteSpace(query))
            {
                var ql = query.Trim();
                q = q.Where(i =>
                    (i.CustomerPhone?.Contains(ql, StringComparison.OrdinalIgnoreCase) ?? false) ||
                    (i.Id?.Contains(ql, StringComparison.OrdinalIgnoreCase) ?? false) ||
                    (i.CustomerName?.Contains(ql, StringComparison.OrdinalIgnoreCase) ?? false));
            }
            return q.OrderByDescending(i => i.LastMessageAt ?? "", StringComparer.Ordinal).ToList();
        }
        return await _store.GetInquiriesAsync(status, query, ct);
    }

    /// <summary>Serves an attachment from the local disk cache, falling back to SharePoint (and caching the
    /// bytes) on a miss. Name is an opaque key the pipeline minted; the caller validates it first.</summary>
    public async Task<(string ContentType, byte[] Bytes)?> GetMediaAsync(string inquiryId, string name, CancellationToken ct = default)
    {
        var path = Path.Combine(_mediaCacheDir, inquiryId, name);
        if (File.Exists(path))
        {
            try { return (InquiryRules.MimeForName(name), await File.ReadAllBytesAsync(path, ct)); }
            catch (Exception ex) { _log.LogWarning(ex, "[InquiryCache] local media read failed {Path}", path); }
        }
        var got = await _messages.GetMediaAsync(inquiryId, name, ct);
        if (got is not null) await WriteMediaToDiskAsync(inquiryId, name, got.Value.Bytes, ct);
        return got;
    }

    // ── In-place updates (local proxy, no SP read) ──────────────────────────────────
    public void ApplyInquiry(Inquiry inq)
    {
        if (string.IsNullOrEmpty(inq.Id)) return;
        if (!IsActive(inq.Status)) { Evict(inq.Id); return; }
        var e = _entries.GetOrAdd(inq.Id, _ => new InquiryCacheEntry());
        lock (e) { e.Inquiry = inq; }
        MarkDirty();
    }

    public void ApplyMessage(string inquiryId, MessageRecord msg)
    {
        if (!_entries.TryGetValue(inquiryId, out var e)) return;
        lock (e)
        {
            var idx = e.Messages.FindIndex(m =>
                (msg.SpItemId is not null && m.SpItemId == msg.SpItemId) ||
                (!string.IsNullOrEmpty(msg.ExternalId) && m.ExternalId == msg.ExternalId));
            if (idx >= 0) e.Messages[idx] = msg; else e.Messages.Add(msg);
        }
        MarkDirty();
        _ = PrefetchMessageMediaAsync(inquiryId, msg, CancellationToken.None);
    }

    public void ApplyDraft(InquiryDraft d)
    {
        if (!_entries.TryGetValue(d.InquiryId, out var e)) return;
        lock (e)
        {
            var idx = d.SpItemId is not null ? e.Drafts.FindIndex(x => x.SpItemId == d.SpItemId) : -1;
            if (idx >= 0) e.Drafts[idx] = d; else e.Drafts.Add(d);
        }
        MarkDirty();
    }

    public void ApplyNote(InquiryNote n)
    {
        if (!_entries.TryGetValue(n.InquiryId, out var e)) return;
        lock (e) { e.Notes.Add(n); }
        MarkDirty();
    }

    public void ApplyQuotation(InquiryQuotation qn)
    {
        if (!_entries.TryGetValue(qn.InquiryId, out var e)) return;
        lock (e)
        {
            if (!e.Quotations.Any(x => string.Equals(x.HskNumber, qn.HskNumber, StringComparison.OrdinalIgnoreCase)))
                e.Quotations.Add(qn);
        }
        MarkDirty();
    }

    public void SetDraftStatus(string inquiryId, int spItemId, string status)
    {
        if (!_entries.TryGetValue(inquiryId, out var e)) return;
        lock (e) { var d = e.Drafts.FirstOrDefault(x => x.SpItemId == spItemId); if (d is not null) d.Status = status; }
        MarkDirty();
    }

    // ── Unread state (Phase 7) — IsRead is button-only; the inquiry counter = unread INBOUND messages ─────
    public void SetMessageRead(string inquiryId, int spItemId, bool read)
    {
        if (!_entries.TryGetValue(inquiryId, out var e)) return;
        lock (e) { var m = e.Messages.FirstOrDefault(x => x.SpItemId == spItemId); if (m is not null) m.IsRead = read; }
        MarkDirty();
    }

    public void SetAllRead(string inquiryId, bool read)
    {
        if (!_entries.TryGetValue(inquiryId, out var e)) return;
        lock (e) { foreach (var m in e.Messages) m.IsRead = read; }
        MarkDirty();
    }

    /// <summary>Count of unread INBOUND messages for an inquiry (the badge/counter source); -1 if not cached.</summary>
    public int UnreadCount(string inquiryId)
    {
        if (!_entries.TryGetValue(inquiryId, out var e)) return -1;
        lock (e) { return e.Messages.Count(m => string.Equals(m.Direction, "in", StringComparison.OrdinalIgnoreCase) && !m.IsRead); }
    }

    /// <summary>Total unread INBOUND messages across all active inquiries — the app-level badge (Phase 7b).</summary>
    public int TotalUnread()
    {
        var total = 0;
        foreach (var e in _entries.Values)
            lock (e) total += e.Messages.Count(m => string.Equals(m.Direction, "in", StringComparison.OrdinalIgnoreCase) && !m.IsRead);
        return total;
    }

    public void Evict(string inquiryId)
    {
        if (_entries.TryRemove(inquiryId, out _)) MarkDirty();
    }

    /// <summary>Targeted SP re-read for one inquiry — used when a PEER proxy reports a change over the bus
    /// (we don't have the mutated objects). Evicts if it became Closed/Spam or vanished.</summary>
    public async Task RefreshOneAsync(string inquiryId, CancellationToken ct = default)
    {
        try
        {
            var inquiry = await _store.GetInquiryByIdAsync(inquiryId, ct);
            if (inquiry is null || !IsActive(inquiry.Status)) { Evict(inquiryId); return; }
            var entry = await BuildEntryAsync(inquiry, ct);
            _entries[inquiryId] = entry;
            MarkDirty();
            _ = PrefetchEntryMediaAsync(entry, ct);
        }
        catch (Exception ex) { _log.LogWarning(ex, "[InquiryCache] refresh {Id} failed", inquiryId); }
    }

    // ── Helpers ─────────────────────────────────────────────────────────────────────
    private static bool IsActive(string? status) =>
        !string.Equals(status, InquiryStatus.Closed, StringComparison.OrdinalIgnoreCase) &&
        !string.Equals(status, InquiryStatus.Spam,   StringComparison.OrdinalIgnoreCase);

    private async Task<InquiryCacheEntry> BuildEntryAsync(Inquiry inquiry, CancellationToken ct)
    {
        var messages   = await _messages.GetByInquiryAsync(inquiry.Id, 500, ct);
        var notes      = await _store.GetNotesByInquiryAsync(inquiry.Id, ct);
        var quotations = await _store.GetQuotationsByInquiryAsync(inquiry.Id, ct);
        var drafts     = await _store.GetDraftsByInquiryAsync(inquiry.Id, ct);
        var contact    = await _store.GetContactAsync(inquiry.CustomerPhone, ct);
        // Phase 7: per-message IsRead is the source of truth — derive the inquiry counter from it so the list
        // counter and the app badge always agree (reconciles pre-Phase-7 rows where UnreadCount was zeroed
        // on open without marking the messages read).
        inquiry.UnreadCount = messages.Count(m => string.Equals(m.Direction, "in", StringComparison.OrdinalIgnoreCase) && !m.IsRead);
        return new InquiryCacheEntry
        {
            Inquiry = inquiry, Messages = [.. messages], Notes = [.. notes],
            Quotations = [.. quotations], Drafts = [.. drafts], Contact = contact,
        };
    }

    private InquiryDetail ToDetail(InquiryCacheEntry e)
    {
        var crm  = _crm.LookupByPhone(e.Inquiry.CustomerPhone);
        var card = new CustomerCard(
            crm?.BusinessPartner ?? e.Inquiry.CustomerName,
            crm?.ContactName ?? e.Inquiry.ContactName,
            crm?.PopupMessage,
            IsFirstTime: crm is null);
        lock (e)
        {
            return new InquiryDetail(e.Inquiry, [.. e.Messages], [.. e.Notes], [.. e.Quotations], [.. e.Drafts], e.Contact, card);
        }
    }

    private InquiryCacheFile Snapshot() => new()
    {
        CacheBuiltUtc = _lastRefreshAt,
        Entries = _entries.Values.Select(e => { lock (e) { return Clone(e); } }).ToList(),
    };

    private static InquiryCacheEntry Clone(InquiryCacheEntry e) => new()
    {
        Inquiry = e.Inquiry, Messages = [.. e.Messages], Notes = [.. e.Notes],
        Quotations = [.. e.Quotations], Drafts = [.. e.Drafts], Contact = e.Contact,
    };

    private void MarkDirty() => _dirty = true;

    // ── Attachment pre-fetch (local disk) ───────────────────────────────────────────
    private async Task PrefetchAllMediaAsync(CancellationToken ct)
    {
        foreach (var e in _entries.Values)
            try { await PrefetchEntryMediaAsync(e, ct); } catch { /* best effort */ }
    }

    private async Task PrefetchEntryMediaAsync(InquiryCacheEntry e, CancellationToken ct)
    {
        List<MessageRecord> msgs;
        lock (e) { msgs = [.. e.Messages]; }
        foreach (var m in msgs) await PrefetchMessageMediaAsync(e.Inquiry.Id, m, ct);
    }

    private async Task PrefetchMessageMediaAsync(string inquiryId, MessageRecord msg, CancellationToken ct)
    {
        foreach (var media in msg.Media)
        {
            if (string.IsNullOrEmpty(media.Name)) continue;
            var path = Path.Combine(_mediaCacheDir, inquiryId, media.Name);
            if (File.Exists(path)) continue;
            try
            {
                var got = await _messages.GetMediaAsync(inquiryId, media.Name, ct);
                if (got is not null) await WriteMediaToDiskAsync(inquiryId, media.Name, got.Value.Bytes, ct);
            }
            catch (Exception ex) { _log.LogWarning(ex, "[InquiryCache] prefetch media {Name} failed", media.Name); }
        }
    }

    private async Task WriteMediaToDiskAsync(string inquiryId, string name, byte[] bytes, CancellationToken ct)
    {
        try
        {
            var dir = Path.Combine(_mediaCacheDir, inquiryId);
            Directory.CreateDirectory(dir);
            await File.WriteAllBytesAsync(Path.Combine(dir, name), bytes, ct);
        }
        catch (Exception ex) { _log.LogWarning(ex, "[InquiryCache] media disk write failed {Id}/{Name}", inquiryId, name); }
    }
}

/// <summary>One cached active inquiry: header + full conversation + sidecar data. The CRM card is NOT stored
/// (recomputed on read so it stays current).</summary>
public sealed class InquiryCacheEntry
{
    public Inquiry                Inquiry    { get; set; } = new();
    public List<MessageRecord>    Messages   { get; set; } = [];
    public List<InquiryNote>      Notes      { get; set; } = [];
    public List<InquiryQuotation> Quotations { get; set; } = [];
    public List<InquiryDraft>     Drafts     { get; set; } = [];
    public MessagingContact?      Contact    { get; set; }
}

public sealed class InquiryCacheFile
{
    public DateTimeOffset?         CacheBuiltUtc { get; set; }
    public List<InquiryCacheEntry> Entries       { get; set; } = [];
}
