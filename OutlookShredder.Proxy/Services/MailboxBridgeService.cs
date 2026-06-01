using System.Text.RegularExpressions;
using Azure.Identity;
using Microsoft.Graph;
using Microsoft.Graph.Models;
using MimeKit;
using OutlookShredder.Proxy.Models;

namespace OutlookShredder.Proxy.Services;

/// <summary>
/// Read-side of the mailbox bridge (see wip/mailbox-bridge.md). Polls a mirror folder
/// on a mailbox we own (e.g. store@mithril → Inbox/Hackensack-Mirror) that a server-side
/// "Forward as attachment" rule fills with copies of mail sent to a mailbox we cannot
/// authenticate into (e.g. hackensack@metalsupermarkets.com).
///
/// Each mirrored message wraps the original supplier email as an embedded
/// message/rfc822 part. We fetch the wrapper's raw MIME (Graph $value) and parse it with
/// MimeKit so the UI shows the real sender/subject/body/attachments instead of
/// "from hackensack, 1 attachment". Parsed headers/bodies are cached in memory and served
/// to the controller; attachment bytes are fetched on demand.
///
/// The public API keys on the WATCHED upn (hackensack@metal); the bridge maps that to the
/// destination mailbox + folder internally. Outbound send is Phase 1.1 and not here yet.
///
/// Auth reuses SharePoint:TenantId/ClientId/ClientSecret (app-only, Mail.ReadWrite), same
/// as MailService. Config: MailboxBridge:PollIntervalSeconds, MailboxBridge:Mailboxes[].
/// </summary>
public sealed class MailboxBridgeService : BackgroundService
{
    private readonly IConfiguration _config;
    private readonly ILogger<MailboxBridgeService> _log;
    private GraphServiceClient? _graph;

    private readonly List<MailboxConfig> _mailboxes;
    private readonly int _pollSeconds;

    // Per-watched-upn live state + parsed-message cache (newest first, capped).
    private sealed class MailboxState
    {
        public required MailboxConfig Config;
        public string? FolderId;
        public MailboxStatus Status = new();
        public readonly object Lock = new();
        // Wrapper message id → parsed view (forward-as-attachment items only).
        public readonly Dictionary<string, CachedMessage> ById = new(StringComparer.Ordinal);
        // Ids classified as "not a forward-as-attachment" (inline forwards etc.) so they
        // are ignored per the product rule and never re-fetched on subsequent polls.
        public readonly HashSet<string> Skipped = new(StringComparer.Ordinal);
        // Wrapper ids already parsed + cached (cache keys are composite "{wrapperId}#{index}" for
        // multi-.eml forwards, so we track parsed wrappers separately from the ById cache).
        public readonly HashSet<string> ParsedWrappers = new(StringComparer.Ordinal);
        // Newest receivedDateTime processed — drives incremental polling (fetch only mail at/after
        // this minus an overlap). Null until the first seed poll has run.
        public DateTimeOffset? HighWater;

        // ── Outbound source (self-BCC'd workbench sends, filed to OutboundFolderPath) ──
        // These are DIRECT messages (the outbound email itself), parsed directly — not the embedded
        // message/rfc822 path used for the inbound mirror. Keyed by the message's own Graph id.
        public string? OutboundFolderId;
        public bool OutboundFolderMissingLogged;
        public readonly Dictionary<string, CachedMessage> OutboundById = new(StringComparer.Ordinal);
        public readonly HashSet<string> OutboundParsed = new(StringComparer.Ordinal);
        public DateTimeOffset? OutboundHighWater;
    }

    private sealed class CachedMessage
    {
        public required MailboxMessageHeader Header;
        public required MailboxMessageBody   Body;
    }

    private const int CacheCap = 250;

    // Cross-proxy claim/dedup via categories on the mirror (wrapper) message — same pattern as the
    // RFQ poller (RFQ-Claiming/RFQ-Processed). A proxy claims a message before capturing so other
    // proxies skip it; marks it processed once fully landed. Distinct names from the RFQ poller (the
    // Forwards folder is excluded from the RFQ scan, but distinct categories avoid any confusion).
    private const string ClaimCategory = "Inbox-Claiming";
    private const string DoneCategory  = "Inbox-Processed";

    // Overlap subtracted from the high-water mark on incremental polls — tolerates out-of-order
    // delivery / commit lag; re-fetched messages in the overlap window dedup cheaply against
    // the ById/Skipped caches.
    private static readonly TimeSpan IncrementalOverlap = TimeSpan.FromMinutes(2);

    private readonly Dictionary<string, MailboxState> _states = new(StringComparer.OrdinalIgnoreCase);

    public MailboxBridgeService(IConfiguration config, ILogger<MailboxBridgeService> log)
    {
        _config = config;
        _log = log;
        _pollSeconds = int.TryParse(_config["MailboxBridge:PollIntervalSeconds"], out var s) && s > 0 ? s : 30;
        _mailboxes = _config.GetSection("MailboxBridge:Mailboxes").Get<MailboxConfig[]>()?.ToList() ?? [];

        foreach (var mb in _mailboxes.Where(m => !string.IsNullOrWhiteSpace(m.WatchedUpn)))
        {
            _states[mb.WatchedUpn] = new MailboxState
            {
                Config = mb,
                Status = new MailboxStatus { WatchedUpn = mb.WatchedUpn, DisplayName = mb.DisplayName },
            };
        }
    }

    private GraphServiceClient GetGraph()
    {
        if (_graph is not null) return _graph;
        var tenantId     = _config["SharePoint:TenantId"]     ?? throw new InvalidOperationException("SharePoint:TenantId not set");
        var clientId     = _config["SharePoint:ClientId"]     ?? throw new InvalidOperationException("SharePoint:ClientId not set");
        var clientSecret = _config["SharePoint:ClientSecret"] ?? throw new InvalidOperationException("SharePoint:ClientSecret not set");
        _graph = new GraphServiceClient(new ClientSecretCredential(tenantId, clientId, clientSecret),
            ["https://graph.microsoft.com/.default"]);
        return _graph;
    }

    // ── Poll loop ──────────────────────────────────────────────────────────────

    protected override async Task ExecuteAsync(CancellationToken ct)
    {
        if (_states.Count == 0)
        {
            _log.LogInformation("[MailboxBridge] No mailboxes configured — bridge idle");
            return;
        }

        _log.LogInformation("[MailboxBridge] Watching {Count} mailbox(es), poll every {Secs}s",
            _states.Count, _pollSeconds);

        try
        {
            while (!ct.IsCancellationRequested)
            {
                foreach (var state in _states.Values)
                {
                    try { await PollMailboxAsync(state, ct); }
                    catch (OperationCanceledException) when (ct.IsCancellationRequested) { throw; }
                    catch (Exception ex)
                    {
                        lock (state.Lock)
                        {
                            state.Status.PollSucceeded = false;
                            state.Status.LastError = ex.Message;
                            state.Status.LastPollAt = DateTimeOffset.UtcNow;
                        }
                        _log.LogWarning(ex, "[MailboxBridge] Poll failed for {Upn}", state.Config.WatchedUpn);
                    }
                }
                await Task.Delay(TimeSpan.FromSeconds(_pollSeconds), ct);
            }
        }
        catch (OperationCanceledException) when (ct.IsCancellationRequested)
        {
            _log.LogInformation("[MailboxBridge] Shutdown requested — bridge stopped cleanly");
        }
    }

    private async Task PollMailboxAsync(MailboxState state, CancellationToken ct)
    {
        var cfg = state.Config;
        state.FolderId ??= await ResolveFolderIdAsync(cfg.DestinationUpn, cfg.DestinationFolderPath, ct);
        if (state.FolderId is null)
            throw new InvalidOperationException(
                $"Mirror folder '{cfg.DestinationFolderPath}' not found in {cfg.DestinationUpn}");

        DateTimeOffset? highWater;
        lock (state.Lock) highWater = state.HighWater;

        var page = await GetGraph().Users[cfg.DestinationUpn].MailFolders[state.FolderId].Messages
            .GetAsync(req =>
            {
                req.QueryParameters.Select  = ["id", "subject", "from", "receivedDateTime", "isRead",
                                               "hasAttachments", "bodyPreview", "categories"];
                req.QueryParameters.Top     = 50;
                req.QueryParameters.Orderby = ["receivedDateTime desc"];

                // Incremental ("delta") fetch: after the first poll, pull only messages at/after the
                // high-water mark minus a small overlap, instead of re-listing the newest 50 every
                // cycle. This keeps fast polling cheap as the mirror folder grows. The first poll
                // (no high-water) seeds the most-recent page so the list view is populated. New mail
                // always sorts newest-first, so a single Top=50 page never misses arrivals between polls.
                if (highWater is { } hw)
                {
                    var since = hw.Subtract(IncrementalOverlap).UtcDateTime.ToString("yyyy-MM-ddTHH:mm:ssZ");
                    req.QueryParameters.Filter = $"receivedDateTime ge {since}";
                }
            }, ct);

        var msgs = page?.Value ?? [];

        foreach (var m in msgs)
        {
            if (m.Id is null) continue;

            // Cross-proxy dedup: skip messages any proxy has already claimed or processed.
            if (m.Categories is not null &&
                (m.Categories.Contains(DoneCategory, StringComparer.OrdinalIgnoreCase) ||
                 m.Categories.Contains(ClaimCategory, StringComparer.OrdinalIgnoreCase)))
                continue;

            bool known, skipped;
            lock (state.Lock) { known = state.ParsedWrappers.Contains(m.Id); skipped = state.Skipped.Contains(m.Id); }
            if (known) continue;     // wrapper already parsed + cached this session
            if (skipped) continue;

            // Forward-as-attachment always carries the original as an .eml attachment, so a
            // wrapper with no attachments is an inline forward — ignore it without fetching MIME.
            if (m.HasAttachments != true)
            {
                lock (state.Lock) state.Skipped.Add(m.Id);
                continue;
            }

            try
            {
                var parsedList = await ParseMirrorMessagesAsync(cfg.DestinationUpn, m, ct);
                lock (state.Lock)
                {
                    if (parsedList.Count == 0)
                    {
                        // Has attachments but no embedded message/rfc822 — an inline forward
                        // whose original carried files. Not a relayed message; ignore it.
                        state.Skipped.Add(m.Id);
                    }
                    else
                    {
                        state.ParsedWrappers.Add(m.Id);
                        foreach (var p in parsedList) state.ById[p.Header.Id] = p;
                        PruneCache(state);
                    }
                }
            }
            catch (OperationCanceledException) when (ct.IsCancellationRequested) { throw; }
            catch (Exception ex)
            {
                _log.LogWarning(ex, "[MailboxBridge] Could not parse mirror message {Id} in {Upn}",
                    m.Id, cfg.WatchedUpn);
            }
        }

        // Bound the Skipped set: drop ids no longer in the polled window. The poll only ever
        // fetches the most-recent page (ordered receivedDateTime desc), so once an inline-forward
        // ages out of that window it can never reappear — keeping it skipped is pointless and the
        // set would otherwise grow unbounded as the mirror folder accumulates mail. Intersecting
        // with the current page caps it at ≤ the page size; the worst case if a dropped id ever
        // resurfaces is one redundant re-classification.
        var pageIds = msgs.Where(m => m.Id is not null).Select(m => m.Id!).ToHashSet(StringComparer.Ordinal);

        lock (state.Lock)
        {
            state.Skipped.IntersectWith(pageIds);
            state.ParsedWrappers.IntersectWith(pageIds);

            // Advance the high-water mark to the newest message seen this poll (incremental polls
            // re-fetch the overlap window, so this only moves forward).
            foreach (var m in msgs)
                if (m.ReceivedDateTime is { } r && (state.HighWater is null || r > state.HighWater))
                    state.HighWater = r;

            state.Status.PollSucceeded = true;
            state.Status.LastError = null;
            state.Status.LastPollAt = DateTimeOffset.UtcNow;
            state.Status.MessageCount = state.ById.Count;
            state.Status.UnreadCount = state.ById.Values.Count(c => !c.Header.IsRead);
        }

        // Second source: the dedicated folder of self-BCC'd workbench sends (direct messages).
        if (!string.IsNullOrWhiteSpace(cfg.OutboundFolderPath))
        {
            try { await PollOutboundAsync(state, ct); }
            catch (OperationCanceledException) when (ct.IsCancellationRequested) { throw; }
            catch (Exception ex) { _log.LogWarning(ex, "[MailboxBridge] outbound poll failed for {Upn}", cfg.WatchedUpn); }
        }
    }

    /// <summary>
    /// Polls the OutboundFolderPath for self-BCC'd workbench sends and parses each as a DIRECT
    /// outbound message (the message itself is the email — no embedded message/rfc822). Threading
    /// uses the same ResolveConversationId logic so an outbound reply keys to the thread it answers.
    /// </summary>
    private async Task PollOutboundAsync(MailboxState state, CancellationToken ct)
    {
        var cfg = state.Config;
        state.OutboundFolderId ??= await ResolveFolderIdAsync(cfg.DestinationUpn, cfg.OutboundFolderPath, ct);
        if (state.OutboundFolderId is null)
        {
            // Folder not created yet (Phase 0 mail rule pending) — skip quietly, log once.
            if (!state.OutboundFolderMissingLogged)
            {
                _log.LogInformation("[MailboxBridge] outbound folder '{Path}' not found in {Upn} yet — skipping outbound poll until it exists",
                    cfg.OutboundFolderPath, cfg.DestinationUpn);
                state.OutboundFolderMissingLogged = true;
            }
            return;
        }

        DateTimeOffset? highWater;
        lock (state.Lock) highWater = state.OutboundHighWater;

        var page = await GetGraph().Users[cfg.DestinationUpn].MailFolders[state.OutboundFolderId].Messages
            .GetAsync(req =>
            {
                req.QueryParameters.Select  = ["id", "subject", "from", "receivedDateTime", "isRead", "hasAttachments", "categories"];
                req.QueryParameters.Top     = 50;
                req.QueryParameters.Orderby = ["receivedDateTime desc"];
                if (highWater is { } hw)
                    req.QueryParameters.Filter = $"receivedDateTime ge {hw.Subtract(IncrementalOverlap).UtcDateTime:yyyy-MM-ddTHH:mm:ssZ}";
            }, ct);

        var msgs = page?.Value ?? [];
        foreach (var m in msgs)
        {
            if (m.Id is null) continue;
            if (m.Categories is not null &&
                (m.Categories.Contains(DoneCategory, StringComparer.OrdinalIgnoreCase) ||
                 m.Categories.Contains(ClaimCategory, StringComparer.OrdinalIgnoreCase)))
                continue;
            bool known; lock (state.Lock) known = state.OutboundParsed.Contains(m.Id);
            if (known) continue;

            try
            {
                var parsed = await ParseDirectMessageAsync(cfg.DestinationUpn, m, ct);
                lock (state.Lock)
                {
                    state.OutboundParsed.Add(m.Id);
                    state.OutboundById[parsed.Header.Id] = parsed;
                    PruneOutbound(state);
                }
            }
            catch (OperationCanceledException) when (ct.IsCancellationRequested) { throw; }
            catch (Exception ex) { _log.LogWarning(ex, "[MailboxBridge] could not parse outbound message {Id}", m.Id); }
        }

        var pageIds = msgs.Where(m => m.Id is not null).Select(m => m.Id!).ToHashSet(StringComparer.Ordinal);
        lock (state.Lock)
        {
            state.OutboundParsed.IntersectWith(pageIds);
            foreach (var m in msgs)
                if (m.ReceivedDateTime is { } r && (state.OutboundHighWater is null || r > state.OutboundHighWater))
                    state.OutboundHighWater = r;
        }
    }

    /// <summary>Parses a direct (non-wrapped) message into an outbound CachedMessage.</summary>
    private async Task<CachedMessage> ParseDirectMessageAsync(string destUpn, Message m, CancellationToken ct)
    {
        await using var mime = await GetGraph().Users[destUpn].Messages[m.Id!].Content.GetAsync(cancellationToken: ct)
            ?? throw new InvalidOperationException("Graph returned no MIME content");
        var src = await MimeMessage.LoadAsync(mime, ct);

        var fromMb      = src.From?.Mailboxes?.FirstOrDefault();
        var receivedIso = ResolveOriginalReceivedIso(src, m.ReceivedDateTime);
        var bodyText    = ExtractPlainBody(src);
        var attachments = src.Attachments.OfType<MimePart>().Select(p => new MailboxAttachmentMeta
        {
            Name        = p.FileName ?? p.ContentType?.Name ?? "attachment",
            ContentType = p.ContentType?.MimeType ?? "application/octet-stream",
            Size        = p.ContentDisposition?.Size ?? 0,
        }).ToList();
        var subject = src.Subject ?? m.Subject ?? "(no subject)";

        return new CachedMessage
        {
            Header = new MailboxMessageHeader
            {
                Id = m.Id!, Subject = subject, FromAddress = fromMb?.Address ?? "", FromName = fromMb?.Name ?? "",
                ReceivedAt = receivedIso, IsRead = true, HasAttachments = attachments.Count > 0,
                Preview = Truncate(bodyText, 200), Direction = "out",
            },
            Body = new MailboxMessageBody
            {
                Id = m.Id!, InternetMessageId = src.MessageId ?? "",
                Subject = subject, FromAddress = fromMb?.Address ?? "", FromName = fromMb?.Name ?? "",
                ToLine = src.To?.ToString() ?? "", CcLine = src.Cc?.ToString() ?? "",
                SourceMailbox = fromMb?.Address ?? destUpn, ConversationId = ResolveConversationId(src),
                ReceivedAt = receivedIso, IsRead = true, BodyText = bodyText, Attachments = attachments,
                Direction = "out",
            },
        };
    }

    /// <summary>
    /// Fetches the wrapper MIME and parses EVERY embedded original (message/rfc822) — a single
    /// forward may carry multiple .eml attachments (bulk history import). Returns one CachedMessage
    /// per embedded email, each with a composite id "{wrapperId}#{index}" (the bare wrapper id when
    /// there is exactly one, for backward-compatibility) and the embedded email's own Internet
    /// Message-ID (the global dedup key). Empty list ⇒ inline forward (ignored).
    /// </summary>
    /// <summary>
    /// The time the ORIGINAL email was received, as UTC ISO-8601. Prefers the topmost "Received:" header
    /// (the recipient server's delivery stamp), then the "Date:" header, then the wrapper's arrival time
    /// as a last resort. UTC-normalized so ISO string sorting equals chronological order.
    /// </summary>
    public static string ResolveOriginalReceivedIso(MimeMessage src, DateTimeOffset? wrapperReceived)
    {
        foreach (var h in src.Headers)
        {
            if (!h.Field.Equals("Received", StringComparison.OrdinalIgnoreCase)) continue;
            var semi = h.Value.LastIndexOf(';');
            if (semi >= 0 && semi < h.Value.Length - 1 &&
                MimeKit.Utils.DateUtils.TryParse(h.Value[(semi + 1)..].Trim(), out var rdt))
                return rdt.ToUniversalTime().ToString("o");
            break; // only the topmost (most recent = delivery) Received header
        }
        if (src.Date.Year > 1971) return src.Date.ToUniversalTime().ToString("o");
        return (wrapperReceived ?? src.Date).ToUniversalTime().ToString("o");
    }

    /// <summary>
    /// Stable per-conversation key from the original email's headers: the Outlook Thread-Index
    /// conversation prefix (first 22 decoded bytes) → References root id → In-Reply-To → own Message-ID.
    /// Lets all messages of one back-and-forth thread group together.
    /// </summary>
    public static string ResolveConversationId(MimeMessage src)
    {
        var ti = src.Headers["Thread-Index"];
        if (!string.IsNullOrWhiteSpace(ti))
        {
            try
            {
                var bytes = Convert.FromBase64String(ti.Trim());
                if (bytes.Length >= 22) return "ti:" + Convert.ToHexString(bytes, 0, 22).ToLowerInvariant();
            }
            catch { /* not valid base64 — fall through */ }
        }
        var root = src.References?.FirstOrDefault();
        if (!string.IsNullOrWhiteSpace(root)) return "ref:" + root.Trim().Trim('<', '>').ToLowerInvariant();
        if (!string.IsNullOrWhiteSpace(src.InReplyTo)) return "ref:" + src.InReplyTo.Trim().Trim('<', '>').ToLowerInvariant();
        if (!string.IsNullOrWhiteSpace(src.MessageId)) return "msg:" + src.MessageId.Trim().Trim('<', '>').ToLowerInvariant();
        return "";
    }

    private async Task<List<CachedMessage>> ParseMirrorMessagesAsync(string destUpn, Message wrapper, CancellationToken ct)
    {
        await using var mime = await GetGraph().Users[destUpn].Messages[wrapper.Id!].Content.GetAsync(cancellationToken: ct)
            ?? throw new InvalidOperationException("Graph returned no MIME content");
        var wrapperMsg = await MimeMessage.LoadAsync(mime, ct);

        var embedded = FindAllEmbeddedMessages(wrapperMsg.Body);
        if (embedded.Count == 0) return [];

        // The wrapper's From is the franchise mailbox that forwarded this (hackensack@ / awathen@ / …),
        // i.e. the originating source mailbox. Fall back to the watched bridge mailbox if absent.
        var sourceMailbox = wrapperMsg.From?.Mailboxes?.FirstOrDefault()?.Address
                            ?? wrapper.From?.EmailAddress?.Address
                            ?? destUpn;

        var result = new List<CachedMessage>(embedded.Count);
        for (int i = 0; i < embedded.Count; i++)
        {
            var src    = embedded[i];
            var id     = embedded.Count == 1 ? wrapper.Id! : $"{wrapper.Id}#{i}";
            var fromMb = src.From?.Mailboxes?.FirstOrDefault();
            // Use the ORIGINAL email's delivery time (from its headers), NOT the wrapper's arrival in
            // the mirror — otherwise backfilled/dumped mail all sorts as "just now" (forward time).
            var receivedIso = ResolveOriginalReceivedIso(src, wrapper.ReceivedDateTime);
            var bodyText = ExtractPlainBody(src);
            var attachments = src.Attachments
                .OfType<MimePart>()
                .Select(p => new MailboxAttachmentMeta
                {
                    Name        = p.FileName ?? p.ContentType?.Name ?? "attachment",
                    ContentType = p.ContentType?.MimeType ?? "application/octet-stream",
                    Size        = p.ContentDisposition?.Size ?? 0,
                })
                .ToList();
            var subject = src.Subject ?? wrapper.Subject ?? "(no subject)";

            result.Add(new CachedMessage
            {
                Header = new MailboxMessageHeader
                {
                    Id = id, Subject = subject, FromAddress = fromMb?.Address ?? "", FromName = fromMb?.Name ?? "",
                    ReceivedAt = receivedIso, IsRead = wrapper.IsRead == true,
                    HasAttachments = attachments.Count > 0, Preview = Truncate(bodyText, 200),
                },
                Body = new MailboxMessageBody
                {
                    Id = id, InternetMessageId = src.MessageId ?? "",
                    Subject = subject, FromAddress = fromMb?.Address ?? "", FromName = fromMb?.Name ?? "",
                    ToLine = src.To?.ToString() ?? "", CcLine = src.Cc?.ToString() ?? "",
                    SourceMailbox = sourceMailbox, ConversationId = ResolveConversationId(src),
                    ReceivedAt = receivedIso, IsRead = wrapper.IsRead == true,
                    BodyText = bodyText, Attachments = attachments,
                },
            });
        }
        return result;
    }

    // ── Public read facade (called by MailboxController) ─────────────────────────

    public IReadOnlyList<MailboxStatus> GetStatuses()
    {
        var list = new List<MailboxStatus>();
        foreach (var s in _states.Values)
            lock (s.Lock) list.Add(CloneStatus(s.Status));
        return list;
    }

    /// <summary>v1 folder tree: a single "Inbox" node representing the mirror folder.</summary>
    public List<FolderNode>? GetFolders(string watchedUpn)
    {
        if (!_states.TryGetValue(watchedUpn, out var s)) return null;
        lock (s.Lock)
        {
            return
            [
                new FolderNode
                {
                    Id          = "inbox",
                    DisplayName = "Inbox",
                    TotalCount  = s.Status.MessageCount,
                    UnreadCount = s.Status.UnreadCount,
                }
            ];
        }
    }

    public List<MailboxMessageHeader>? GetMessages(string watchedUpn, int top)
    {
        if (!_states.TryGetValue(watchedUpn, out var s)) return null;
        lock (s.Lock)
        {
            return s.ById.Values
                .Select(c => c.Header)
                .OrderByDescending(h => h.ReceivedAt, StringComparer.Ordinal)
                .Take(top)
                .ToList();
        }
    }

    public MailboxMessageBody? GetMessage(string watchedUpn, string id)
    {
        if (!_states.TryGetValue(watchedUpn, out var s)) return null;
        lock (s.Lock)
        {
            if (s.ById.TryGetValue(id, out var c)) return c.Body;
            if (s.OutboundById.TryGetValue(id, out var o)) return o.Body;
            return null;
        }
    }

    /// <summary>Outbound (self-BCC'd workbench send) message bodies, newest first.</summary>
    public List<MailboxMessageBody> GetOutboundMessages(string watchedUpn, int top)
    {
        if (!_states.TryGetValue(watchedUpn, out var s)) return [];
        lock (s.Lock)
        {
            return s.OutboundById.Values
                .Select(c => c.Body)
                .OrderByDescending(b => b.ReceivedAt, StringComparer.Ordinal)
                .Take(top)
                .ToList();
        }
    }

    /// <summary>
    /// Paginates the ENTIRE mirror folder (not just the polled top-50 window) and returns the parsed
    /// body of every forward-as-attachment message. Used by the workbench backfill to capture the
    /// full history in one pass — the live poll only keeps the recent window in its cache. Fetches +
    /// parses each message's MIME, so it's expensive; intended as an on-demand background sweep.
    /// </summary>
    public async Task<List<MailboxMessageBody>> EnumerateAllForwardedAsync(string watchedUpn, CancellationToken ct = default)
    {
        if (!_states.TryGetValue(watchedUpn, out var state)) return [];
        var cfg = state.Config;
        state.FolderId ??= await ResolveFolderIdAsync(cfg.DestinationUpn, cfg.DestinationFolderPath, ct);
        if (state.FolderId is null) return [];

        var result = new List<MailboxMessageBody>();
        var page = await GetGraph().Users[cfg.DestinationUpn].MailFolders[state.FolderId].Messages
            .GetAsync(req =>
            {
                req.QueryParameters.Select  = ["id", "subject", "from", "receivedDateTime", "isRead", "hasAttachments", "categories"];
                req.QueryParameters.Top     = 50;
                req.QueryParameters.Orderby = ["receivedDateTime desc"];
            }, ct);

        while (page?.Value is not null)
        {
            foreach (var m in page.Value)
            {
                if (m.Id is null || m.HasAttachments != true) continue;   // inline forwards skipped (no .eml)
                if (m.Categories is not null &&
                    (m.Categories.Contains(DoneCategory, StringComparer.OrdinalIgnoreCase) ||
                     m.Categories.Contains(ClaimCategory, StringComparer.OrdinalIgnoreCase)))
                    continue;   // already claimed/processed by some proxy
                try
                {
                    foreach (var p in await ParseMirrorMessagesAsync(cfg.DestinationUpn, m, ct))
                        result.Add(p.Body);
                }
                catch (OperationCanceledException) when (ct.IsCancellationRequested) { throw; }
                catch (Exception ex) { _log.LogWarning(ex, "[MailboxBridge] backfill parse failed for {Id}", m.Id); }
            }
            if (string.IsNullOrEmpty(page.OdataNextLink)) break;
            page = await GetGraph().Users[cfg.DestinationUpn].MailFolders[state.FolderId].Messages
                .WithUrl(page.OdataNextLink).GetAsync(cancellationToken: ct);
        }

        _log.LogInformation("[MailboxBridge] backfill enumerated {N} forward-as-attachment items in {Upn}",
            result.Count, watchedUpn);
        return result;
    }

    /// <summary>Re-fetches the wrapper MIME and returns the named attachment bytes from the embedded original.</summary>
    public async Task<(string ContentType, byte[] Bytes, string FileName)?> GetAttachmentAsync(
        string watchedUpn, string messageId, string attachmentName, CancellationToken ct)
    {
        if (!_states.TryGetValue(watchedUpn, out var s)) return null;
        var destUpn = s.Config.DestinationUpn;

        var (wrapperId, index) = SplitId(messageId);
        await using var mime = await GetGraph().Users[destUpn].Messages[wrapperId].Content.GetAsync(cancellationToken: ct);
        if (mime is null) return null;
        var wrapperMsg = await MimeMessage.LoadAsync(mime, ct);
        var embedded = FindAllEmbeddedMessages(wrapperMsg.Body);
        var src = index < embedded.Count ? embedded[index] : (embedded.FirstOrDefault() ?? wrapperMsg);

        var part = src.Attachments
            .OfType<MimePart>()
            .FirstOrDefault(p => string.Equals(p.FileName ?? p.ContentType?.Name, attachmentName, StringComparison.OrdinalIgnoreCase));
        if (part is null) return null;

        using var ms = new MemoryStream();
        await part.Content.DecodeToAsync(ms, ct);
        return (part.ContentType?.MimeType ?? "application/octet-stream", ms.ToArray(),
                part.FileName ?? attachmentName);
    }

    /// <summary>
    /// Atomically-ish claims the wrapper message for this proxy by adding the claim category.
    /// Returns false if it is already processed or already claimed (another proxy has it). Same
    /// caveat as the RFQ poller: Graph PATCH is not a true compare-and-set, so a narrow race remains
    /// (caller keeps the in-memory wrapper-id dedup as a backstop; a periodic dedup is the safety net).
    /// </summary>
    public async Task<bool> TryClaimAsync(string watchedUpn, string messageId, CancellationToken ct = default)
    {
        if (!_states.TryGetValue(watchedUpn, out var s)) return false;
        var dest = s.Config.DestinationUpn;
        var msg  = await GetGraph().Users[dest].Messages[messageId].GetAsync(r => r.QueryParameters.Select = ["categories"], ct);
        var cats = msg?.Categories?.ToList() ?? [];
        if (cats.Contains(DoneCategory,  StringComparer.OrdinalIgnoreCase)) return false;  // already landed
        if (cats.Contains(ClaimCategory, StringComparer.OrdinalIgnoreCase)) return false;  // already claimed
        cats.Add(ClaimCategory);
        await GetGraph().Users[dest].Messages[messageId].PatchAsync(new Message { Categories = cats }, cancellationToken: ct);
        return true;
    }

    /// <summary>Marks the wrapper message processed (adds the done category, clears the claim).</summary>
    public async Task MarkProcessedAsync(string watchedUpn, string messageId, CancellationToken ct = default)
    {
        if (!_states.TryGetValue(watchedUpn, out var s)) return;
        var dest = s.Config.DestinationUpn;
        var msg  = await GetGraph().Users[dest].Messages[messageId].GetAsync(r => r.QueryParameters.Select = ["categories"], ct);
        var cats = (msg?.Categories ?? []).Where(c => !string.Equals(c, ClaimCategory, StringComparison.OrdinalIgnoreCase)).ToList();
        if (!cats.Contains(DoneCategory, StringComparer.OrdinalIgnoreCase)) cats.Add(DoneCategory);
        await GetGraph().Users[dest].Messages[messageId].PatchAsync(new Message { Categories = cats }, cancellationToken: ct);
    }

    /// <summary>
    /// Strips the Inbox-Claiming/Inbox-Processed categories from every mirror message and resets the
    /// in-memory poll state so everything re-surfaces for a fresh re-capture (used by purge/reset).
    /// </summary>
    public async Task<int> ResetClaimCategoriesAsync(string watchedUpn, CancellationToken ct = default)
    {
        if (!_states.TryGetValue(watchedUpn, out var state)) return 0;
        var cfg = state.Config;
        state.FolderId ??= await ResolveFolderIdAsync(cfg.DestinationUpn, cfg.DestinationFolderPath, ct);
        if (state.FolderId is null) return 0;

        int cleared = 0;
        var page = await GetGraph().Users[cfg.DestinationUpn].MailFolders[state.FolderId].Messages
            .GetAsync(req => { req.QueryParameters.Select = ["id", "categories"]; req.QueryParameters.Top = 50; }, ct);
        while (page?.Value is not null)
        {
            foreach (var m in page.Value)
            {
                if (m.Id is null || m.Categories is null) continue;
                if (!m.Categories.Contains(ClaimCategory, StringComparer.OrdinalIgnoreCase) &&
                    !m.Categories.Contains(DoneCategory,  StringComparer.OrdinalIgnoreCase)) continue;
                var kept = m.Categories.Where(c => !string.Equals(c, ClaimCategory, StringComparison.OrdinalIgnoreCase)
                                                && !string.Equals(c, DoneCategory,  StringComparison.OrdinalIgnoreCase)).ToList();
                await GetGraph().Users[cfg.DestinationUpn].Messages[m.Id].PatchAsync(new Message { Categories = kept }, cancellationToken: ct);
                cleared++;
            }
            if (string.IsNullOrEmpty(page.OdataNextLink)) break;
            page = await GetGraph().Users[cfg.DestinationUpn].MailFolders[state.FolderId].Messages
                .WithUrl(page.OdataNextLink).GetAsync(cancellationToken: ct);
        }

        lock (state.Lock)
        {
            state.ById.Clear(); state.ParsedWrappers.Clear(); state.Skipped.Clear(); state.HighWater = null;
            state.OutboundById.Clear(); state.OutboundParsed.Clear(); state.OutboundHighWater = null;
        }
        _log.LogInformation("[MailboxBridge] reset: cleared claim/processed categories on {N} message(s)", cleared);
        return cleared;
    }

    /// <summary>Returns the raw .eml bytes of the embedded original message (for archival).</summary>
    public async Task<byte[]?> GetRawEmlAsync(string watchedUpn, string messageId, CancellationToken ct = default)
    {
        if (!_states.TryGetValue(watchedUpn, out var s)) return null;
        var (wrapperId, index) = SplitId(messageId);
        await using var mime = await GetGraph().Users[s.Config.DestinationUpn].Messages[wrapperId].Content.GetAsync(cancellationToken: ct);
        if (mime is null) return null;
        var wrapperMsg = await MimeMessage.LoadAsync(mime, ct);
        var embedded = FindAllEmbeddedMessages(wrapperMsg.Body);
        if (index >= embedded.Count) return null;
        var src = embedded[index];
        using var ms = new MemoryStream();
        await src.WriteToAsync(ms, ct);
        return ms.ToArray();
    }

    /// <summary>Attachment bytes from a DIRECT outbound message (the message itself, not an embedded original).</summary>
    public async Task<(string ContentType, byte[] Bytes, string FileName)?> GetDirectAttachmentAsync(
        string watchedUpn, string messageId, string attachmentName, CancellationToken ct)
    {
        if (!_states.TryGetValue(watchedUpn, out var s)) return null;
        await using var mime = await GetGraph().Users[s.Config.DestinationUpn].Messages[messageId].Content.GetAsync(cancellationToken: ct);
        if (mime is null) return null;
        var msg = await MimeMessage.LoadAsync(mime, ct);
        var part = msg.Attachments.OfType<MimePart>()
            .FirstOrDefault(p => string.Equals(p.FileName ?? p.ContentType?.Name, attachmentName, StringComparison.OrdinalIgnoreCase));
        if (part is null) return null;
        using var ms = new MemoryStream();
        await part.Content.DecodeToAsync(ms, ct);
        return (part.ContentType?.MimeType ?? "application/octet-stream", ms.ToArray(), part.FileName ?? attachmentName);
    }

    /// <summary>Raw .eml of a DIRECT outbound message (its own MIME).</summary>
    public async Task<byte[]?> GetDirectRawEmlAsync(string watchedUpn, string messageId, CancellationToken ct = default)
    {
        if (!_states.TryGetValue(watchedUpn, out var s)) return null;
        await using var mime = await GetGraph().Users[s.Config.DestinationUpn].Messages[messageId].Content.GetAsync(cancellationToken: ct);
        if (mime is null) return null;
        using var ms = new MemoryStream();
        await mime.CopyToAsync(ms, ct);
        return ms.ToArray();
    }

    public async Task<bool> SetReadAsync(string watchedUpn, string messageId, bool read, CancellationToken ct)
    {
        if (!_states.TryGetValue(watchedUpn, out var s)) return false;
        await GetGraph().Users[s.Config.DestinationUpn].Messages[messageId]
            .PatchAsync(new Message { IsRead = read }, cancellationToken: ct);
        lock (s.Lock)
        {
            if (s.ById.TryGetValue(messageId, out var c)) { c.Header.IsRead = read; c.Body.IsRead = read; }
        }
        return true;
    }

    public bool IsWatched(string watchedUpn) => _states.ContainsKey(watchedUpn);
    public int MailboxCount => _states.Count;

    // ── Helpers ──────────────────────────────────────────────────────────────────

    /// <summary>Resolves a "/"-delimited folder path (e.g. "Inbox/Hackensack-Mirror") to a Graph folder id.</summary>
    private async Task<string?> ResolveFolderIdAsync(string upn, string path, CancellationToken ct)
    {
        var segments = (path ?? "").Split('/', StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries);
        if (segments.Length == 0) return null;

        // First segment: prefer the well-known name for "Inbox"; otherwise match by display name at root.
        string? currentId;
        if (segments[0].Equals("Inbox", StringComparison.OrdinalIgnoreCase))
        {
            var inbox = await GetGraph().Users[upn].MailFolders["inbox"].GetAsync(r => r.QueryParameters.Select = ["id"], ct);
            currentId = inbox?.Id;
        }
        else
        {
            currentId = await FindChildFolderIdAsync(upn, null, segments[0], ct);
        }
        if (currentId is null) return null;

        for (int i = 1; i < segments.Length && currentId is not null; i++)
            currentId = await FindChildFolderIdAsync(upn, currentId, segments[i], ct);

        return currentId;
    }

    private async Task<string?> FindChildFolderIdAsync(string upn, string? parentId, string displayName, CancellationToken ct)
    {
        var page = await (parentId is null
            ? GetGraph().Users[upn].MailFolders.GetAsync(r => {
                r.QueryParameters.Filter = $"displayName eq '{displayName.Replace("'", "''")}'";
                r.QueryParameters.Select = ["id", "displayName"];
            }, ct)
            : GetGraph().Users[upn].MailFolders[parentId].ChildFolders.GetAsync(r => {
                r.QueryParameters.Filter = $"displayName eq '{displayName.Replace("'", "''")}'";
                r.QueryParameters.Select = ["id", "displayName"];
            }, ct));

        return page?.Value?.FirstOrDefault()?.Id;
    }

    /// <summary>Depth-first collection of ALL embedded message/rfc822 parts (forwarded originals).</summary>
    private static List<MimeMessage> FindAllEmbeddedMessages(MimeEntity? entity)
    {
        var found = new List<MimeMessage>();
        void Walk(MimeEntity? e)
        {
            switch (e)
            {
                case MessagePart mp: if (mp.Message is not null) found.Add(mp.Message); break;
                case Multipart multi: foreach (var child in multi) Walk(child); break;
            }
        }
        Walk(entity);
        return found;
    }

    /// <summary>Splits a cached item id into (wrapperGraphId, embeddedIndex). Bare ids ⇒ index 0.</summary>
    private static (string WrapperId, int Index) SplitId(string id)
    {
        var h = id.LastIndexOf('#');
        if (h < 0) return (id, 0);
        return (id[..h], int.TryParse(id[(h + 1)..], out var n) ? n : 0);
    }

    private static readonly Regex _htmlTag = new(@"<[^>]+>", RegexOptions.Compiled);
    private static readonly Regex _ws      = new(@"[ \t]{2,}", RegexOptions.Compiled);

    private static string ExtractPlainBody(MimeMessage msg)
    {
        if (!string.IsNullOrWhiteSpace(msg.TextBody)) return msg.TextBody.Trim();
        if (!string.IsNullOrWhiteSpace(msg.HtmlBody))
        {
            var text = _htmlTag.Replace(msg.HtmlBody, " ");
            text = System.Net.WebUtility.HtmlDecode(text);
            return _ws.Replace(text, " ").Trim();
        }
        return "";
    }

    private static string Truncate(string s, int max) =>
        string.IsNullOrEmpty(s) || s.Length <= max ? s : s[..max] + "…";

    private static void PruneCache(MailboxState state)
    {
        if (state.ById.Count <= CacheCap) return;
        var drop = state.ById.Values
            .OrderBy(c => c.Header.ReceivedAt, StringComparer.Ordinal)
            .Take(state.ById.Count - CacheCap)
            .Select(c => c.Header.Id)
            .ToList();
        foreach (var id in drop) state.ById.Remove(id);
    }

    private static void PruneOutbound(MailboxState state)
    {
        if (state.OutboundById.Count <= CacheCap) return;
        var drop = state.OutboundById.Values
            .OrderBy(c => c.Header.ReceivedAt, StringComparer.Ordinal)
            .Take(state.OutboundById.Count - CacheCap)
            .Select(c => c.Header.Id)
            .ToList();
        foreach (var id in drop) state.OutboundById.Remove(id);
    }

    private static MailboxStatus CloneStatus(MailboxStatus s) => new()
    {
        WatchedUpn = s.WatchedUpn, DisplayName = s.DisplayName, PollSucceeded = s.PollSucceeded,
        LastError = s.LastError, LastPollAt = s.LastPollAt, MessageCount = s.MessageCount, UnreadCount = s.UnreadCount,
    };
}
