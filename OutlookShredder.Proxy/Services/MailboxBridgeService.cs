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
    }

    private sealed class CachedMessage
    {
        public required MailboxMessageHeader Header;
        public required MailboxMessageBody   Body;
    }

    private const int CacheCap = 250;

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

        var page = await GetGraph().Users[cfg.DestinationUpn].MailFolders[state.FolderId].Messages
            .GetAsync(req =>
            {
                req.QueryParameters.Select  = ["id", "subject", "from", "receivedDateTime", "isRead",
                                               "hasAttachments", "bodyPreview"];
                req.QueryParameters.Top     = 50;
                req.QueryParameters.Orderby = ["receivedDateTime desc"];
            }, ct);

        var msgs = page?.Value ?? [];

        foreach (var m in msgs)
        {
            if (m.Id is null) continue;

            bool known, skipped;
            lock (state.Lock) { known = state.ById.ContainsKey(m.Id); skipped = state.Skipped.Contains(m.Id); }

            if (known)
            {
                // Keep read-state fresh without re-parsing the MIME.
                lock (state.Lock)
                {
                    if (state.ById.TryGetValue(m.Id, out var c))
                    {
                        c.Header.IsRead = m.IsRead == true;
                        c.Body.IsRead   = m.IsRead == true;
                    }
                }
                continue;
            }
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
                var parsed = await ParseMirrorMessageAsync(cfg.DestinationUpn, m, ct);
                lock (state.Lock)
                {
                    if (parsed is null)
                    {
                        // Has attachments but no embedded message/rfc822 — an inline forward
                        // whose original carried files. Not a relayed message; ignore it.
                        state.Skipped.Add(m.Id);
                    }
                    else
                    {
                        state.ById[m.Id] = parsed;
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

        lock (state.Lock)
        {
            state.Status.PollSucceeded = true;
            state.Status.LastError = null;
            state.Status.LastPollAt = DateTimeOffset.UtcNow;
            state.Status.MessageCount = state.ById.Count;
            state.Status.UnreadCount = state.ById.Values.Count(c => !c.Header.IsRead);
        }
    }

    /// <summary>
    /// Fetches the wrapper message's raw MIME and parses the embedded original (message/rfc822)
    /// produced by "Forward as attachment". Returns null when there is no embedded message —
    /// the folder also receives inline forwards, which serve a different purpose and are ignored.
    /// </summary>
    private async Task<CachedMessage?> ParseMirrorMessageAsync(string destUpn, Message wrapper, CancellationToken ct)
    {
        await using var mime = await GetGraph().Users[destUpn].Messages[wrapper.Id!].Content.GetAsync(cancellationToken: ct)
            ?? throw new InvalidOperationException("Graph returned no MIME content");
        var wrapperMsg = await MimeMessage.LoadAsync(mime, ct);

        // Only true forward-as-attachment messages carry the original as an embedded
        // message/rfc822 part. No embedded message ⇒ inline forward ⇒ ignore.
        var src = FindEmbeddedMessage(wrapperMsg.Body);
        if (src is null) return null;

        var fromMb = src.From?.Mailboxes?.FirstOrDefault();
        var receivedIso = (wrapper.ReceivedDateTime ?? src.Date).ToString("o");
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
        var header = new MailboxMessageHeader
        {
            Id             = wrapper.Id!,
            Subject        = subject,
            FromAddress    = fromMb?.Address ?? "",
            FromName       = fromMb?.Name ?? "",
            ReceivedAt     = receivedIso,
            IsRead         = wrapper.IsRead == true,
            HasAttachments = attachments.Count > 0,
            Preview        = Truncate(bodyText, 200),
        };
        var body = new MailboxMessageBody
        {
            Id          = wrapper.Id!,
            Subject     = subject,
            FromAddress = fromMb?.Address ?? "",
            FromName    = fromMb?.Name ?? "",
            ToLine      = src.To?.ToString() ?? "",
            CcLine      = src.Cc?.ToString() ?? "",
            ReceivedAt  = receivedIso,
            IsRead      = wrapper.IsRead == true,
            BodyText    = bodyText,
            Attachments = attachments,
        };
        return new CachedMessage { Header = header, Body = body };
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
        lock (s.Lock) return s.ById.TryGetValue(id, out var c) ? c.Body : null;
    }

    /// <summary>Re-fetches the wrapper MIME and returns the named attachment bytes from the embedded original.</summary>
    public async Task<(string ContentType, byte[] Bytes, string FileName)?> GetAttachmentAsync(
        string watchedUpn, string messageId, string attachmentName, CancellationToken ct)
    {
        if (!_states.TryGetValue(watchedUpn, out var s)) return null;
        var destUpn = s.Config.DestinationUpn;

        await using var mime = await GetGraph().Users[destUpn].Messages[messageId].Content.GetAsync(cancellationToken: ct);
        if (mime is null) return null;
        var wrapperMsg = await MimeMessage.LoadAsync(mime, ct);
        var src = FindEmbeddedMessage(wrapperMsg.Body) ?? wrapperMsg;

        var part = src.Attachments
            .OfType<MimePart>()
            .FirstOrDefault(p => string.Equals(p.FileName ?? p.ContentType?.Name, attachmentName, StringComparison.OrdinalIgnoreCase));
        if (part is null) return null;

        using var ms = new MemoryStream();
        await part.Content.DecodeToAsync(ms, ct);
        return (part.ContentType?.MimeType ?? "application/octet-stream", ms.ToArray(),
                part.FileName ?? attachmentName);
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

    /// <summary>Depth-first search for the first embedded message/rfc822 part (the forwarded original).</summary>
    private static MimeMessage? FindEmbeddedMessage(MimeEntity? entity)
    {
        switch (entity)
        {
            case null: return null;
            case MessagePart mp: return mp.Message;
            case Multipart multi:
                foreach (var child in multi)
                {
                    var found = FindEmbeddedMessage(child);
                    if (found is not null) return found;
                }
                return null;
            default: return null;
        }
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

    private static MailboxStatus CloneStatus(MailboxStatus s) => new()
    {
        WatchedUpn = s.WatchedUpn, DisplayName = s.DisplayName, PollSucceeded = s.PollSucceeded,
        LastError = s.LastError, LastPollAt = s.LastPollAt, MessageCount = s.MessageCount, UnreadCount = s.UnreadCount,
    };
}
