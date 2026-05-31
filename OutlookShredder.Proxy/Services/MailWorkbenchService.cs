using System.Collections.Concurrent;
using System.IO;
using System.Text.Json;
using OutlookShredder.Proxy.Models;

namespace OutlookShredder.Proxy.Services;

/// <summary>
/// Orchestrates the mail workbench (wip/mail-classification.md, Phase 1b): capture a bridge
/// message into the immutable MailItems store, run the AI classifier, and write a versioned
/// MailClassifications row. Also serves the classification tree + per-node item lists and the
/// non-destructive re-classify path. Attachment/raw-eml storage in a doc library is Phase 1b-ii.
/// </summary>
public sealed class MailWorkbenchService
{
    private readonly SharePointService _sp;
    private readonly MailClassifierService _classifier;
    private readonly MailboxBridgeService _bridge;
    private readonly MailTaxonomyService _taxonomy;
    private readonly MailCacheService _cache;
    private readonly RfqNotificationService _notify;
    private readonly ILogger<MailWorkbenchService> _log;
    private readonly SeedProgress _seed = new();
    private readonly string _archiveRoot;

    public MailWorkbenchService(SharePointService sp, MailClassifierService classifier,
        MailboxBridgeService bridge, MailTaxonomyService taxonomy, MailCacheService cache,
        RfqNotificationService notify, IConfiguration config, ILogger<MailWorkbenchService> log)
    {
        _sp = sp; _classifier = classifier; _bridge = bridge; _taxonomy = taxonomy;
        _cache = cache; _notify = notify; _log = log;

        // Storage root = the OneDrive-synced "Shredder" folder, sibling to the publish directory
        // (…\Metal Supermarkets Hackensack - Documents\Shredder). Files are written locally; OneDrive
        // syncs them to the document library. Override with MailArchive:RootPath.
        _archiveRoot = config["MailArchive:RootPath"] ?? Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.UserProfile),
            "Mithril Metals Corp", "Metal Supermarkets Hackensack - Documents", "Shredder");
        try
        {
            Directory.CreateDirectory(Path.Combine(_archiveRoot, "Inbox"));   // Phase-4 watched-dir drop point
            foreach (var leaf in MailTaxonomy.Leaves)
                Directory.CreateDirectory(Path.Combine(_archiveRoot, leaf.Path.Replace('/', Path.DirectorySeparatorChar)));
            _log.LogInformation("[MailWB] archive root: {Root}", _archiveRoot);
        }
        catch (Exception ex) { _log.LogWarning(ex, "[MailWB] could not create archive folder tree at {Root}", _archiveRoot); }
    }

    /// <summary>Human-readable folder for an item's files: {root}\{category}\{yyyy-MM-dd}_{subject}_{shortId}\</summary>
    private string ItemFolder(string category, string receivedIso, string subject, string mailItemId) =>
        Path.Combine(_archiveRoot, category.Replace('/', Path.DirectorySeparatorChar),
            FolderName(receivedIso, subject, mailItemId));

    private static string FolderName(string receivedIso, string subject, string mailItemId)
    {
        var date = DateTimeOffset.TryParse(receivedIso, out var d) ? d.ToString("yyyy-MM-dd") : "nodate";
        var sid  = mailItemId.Length >= 8 ? mailItemId[..8] : mailItemId;   // short uniqueness suffix
        return $"{date}_{SlugText(subject, 50)}_{sid}";
    }

    /// <summary>Turns a subject into a filesystem-safe, readable slug.</summary>
    private static string SlugText(string? s, int max)
    {
        var sb = new System.Text.StringBuilder();
        foreach (var c in (s ?? "").Trim())
        {
            if (char.IsLetterOrDigit(c)) sb.Append(c);
            else if (c is ' ' or '-' or '_' or '.' or ',') sb.Append('-');
        }
        var slug = sb.ToString().Trim('-');
        while (slug.Contains("--")) slug = slug.Replace("--", "-");
        if (slug.Length > max) slug = slug[..max].Trim('-');
        return slug.Length == 0 ? "item" : slug;
    }

    /// <summary>Dedup set seeded with existing Internet Message-IDs ∪ wrapper/composite ids (handles
    /// both new Message-ID-keyed items and legacy wrapper-id-keyed ones).</summary>
    private async Task<ConcurrentDictionary<string, byte>> BuildSeenAsync(CancellationToken ct)
    {
        var d = new ConcurrentDictionary<string, byte>(StringComparer.Ordinal);
        foreach (var k in await _sp.GetExistingMessageIdsAsync(ct)) d.TryAdd(k, 0);
        foreach (var k in await _sp.GetExistingWrapperIdsAsync(ct)) d.TryAdd(k, 0);
        return d;
    }

    /// <summary>Splits a composite item id "{wrapperId}#{index}" into its parts (bare id ⇒ index 0).</summary>
    private static (string WrapperId, int Index) SplitId(string id)
    {
        var h = id.LastIndexOf('#');
        return h < 0 ? (id, 0) : (id[..h], int.TryParse(id[(h + 1)..], out var n) ? n : 0);
    }

    /// <summary>Capture a cached bridge message → MailItems (dedup) → classify → MailClassifications.</summary>
    public async Task<CaptureResult> CaptureAndClassifyAsync(string watchedUpn, string messageId,
        ConcurrentDictionary<string, byte>? seen = null,
        ConcurrentDictionary<string, Lazy<Task<bool>>>? wrapperClaims = null, CancellationToken ct = default)
    {
        var body = _bridge.GetMessage(watchedUpn, messageId)
            ?? throw new InvalidOperationException("Message not in bridge cache (not polled yet).");
        return await CaptureBodyAsync(watchedUpn, body, seen, wrapperClaims, ct);
    }

    /// <summary>Capture+classify a parsed message body directly — used by both the cache path and the full-folder backfill.</summary>
    public async Task<CaptureResult> CaptureBodyAsync(string watchedUpn, MailboxMessageBody body,
        ConcurrentDictionary<string, byte>? seen = null,
        ConcurrentDictionary<string, Lazy<Task<bool>>>? wrapperClaims = null, CancellationToken ct = default)
    {
        seen ??= await BuildSeenAsync(ct);
        wrapperClaims ??= new(StringComparer.Ordinal);

        // Dedup on the embedded email's own Internet Message-ID (stable across re-forwards) so a
        // re-sent .eml is DISCARDED, not duplicated; fall back to the wrapper/composite id when an
        // email lacks a Message-ID.
        var dedupKey = !string.IsNullOrEmpty(body.InternetMessageId) ? body.InternetMessageId : body.Id;
        if (!seen.TryAdd(dedupKey, 0))
            return new CaptureResult { MailItemId = "", IsNew = false };

        // Cross-proxy claim on the WRAPPER message — one claim covers all its embedded .emls.
        // Lazy<Task> ensures TryClaim runs exactly once per wrapper even under concurrent siblings.
        var (wrapperId, _) = SplitId(body.Id);
        var claim = wrapperClaims.GetOrAdd(wrapperId, w => new Lazy<Task<bool>>(() => _bridge.TryClaimAsync(watchedUpn, w, ct)));
        if (!await claim.Value)
            return new CaptureResult { MailItemId = "", IsNew = false };

        var manifest = body.Attachments
            .Select(a => new MailAttManifest { Name = a.Name, ContentType = a.ContentType, Size = a.Size })
            .ToList();

        var mailItemId = await _sp.WriteMailItemAsync(new MailItemInput
        {
            SourceType        = "email",
            SourceMailbox     = watchedUpn,
            WrapperGraphId    = body.Id,
            InternetMessageId = body.InternetMessageId,
            FromAddress       = body.FromAddress,
            FromName        = body.FromName,
            ToLine          = body.ToLine,
            CcLine          = body.CcLine,
            Subject         = body.Subject,
            ReceivedAtIso   = body.ReceivedAt,
            BodyText        = body.BodyText,
            HasAttachments  = manifest.Count > 0,
            AttachmentsJson = JsonSerializer.Serialize(manifest),
        }, ct);

        var input = new MailClassifyInput
        {
            Subject = body.Subject, FromAddress = body.FromAddress, FromName = body.FromName,
            ToLine = body.ToLine, BodyText = body.BodyText,
            AttachmentNames = manifest.Select(a => a.Name).ToList(),
        };
        var result   = await _classifier.ClassifyAsync(input, ct);
        var category = result?.Category ?? "Unclassified";

        // Always archive (folder + raw .eml, plus any attachments) so the full HTML viewer can render
        // from the .eml later — not just when attachments are present.
        await StoreFilesAsync(watchedUpn, body, mailItemId, category, body.ReceivedAt, ct);

        var version = result is not null ? await _sp.WriteClassificationAsync(mailItemId, result, ct) : 0;

        // Fully landed (item + classification + files) → mark the WRAPPER processed so no proxy reprocesses it.
        await _bridge.MarkProcessedAsync(watchedUpn, wrapperId, ct);

        var row = new MailItemRow
        {
            MailItemId = mailItemId, WrapperGraphId = body.Id, SourceType = "email", SourceMailbox = watchedUpn,
            FromAddress = body.FromAddress, FromName = body.FromName, Subject = body.Subject,
            ReceivedAt = body.ReceivedAt, HasAttachments = manifest.Count > 0,
            AttachmentsJson = JsonSerializer.Serialize(manifest), Completed = false,
        };
        var cls = result is not null ? ToClassRow(mailItemId, result, version) : null;
        CacheAndPublish("Captured", row, cls);

        return new CaptureResult
        {
            MailItemId = mailItemId, IsNew = true, Classified = result is not null,
            Category = result?.Category, Confidence = result?.Confidence ?? 0, Version = version,
        };
    }

    // ── Cache + bus fan-out (keeps every machine's Inbox coherent without per-request SP reads) ──

    private static MailClassRow ToClassRow(string mailItemId, MailClassificationResult r, int version) => new()
    {
        MailItemId = mailItemId, Version = version, IsCurrent = true, CategoryPath = r.Category,
        OtherLabel = r.OtherLabel, Confidence = r.Confidence, KeywordTags = string.Join(", ", r.Keywords),
        PoNumber = r.PoNumber, SoNumber = r.SoNumber, Amount = r.Amount,
        AiProvider = r.AiProvider, AiModel = r.AiModel,
    };

    /// <summary>Upsert the item (+ its current classification) into the local cache and broadcast a "Mail" bus event.</summary>
    private void CacheAndPublish(string action, MailItemRow row, MailClassRow? cls)
    {
        try
        {
            _cache.UpsertItem(row);
            if (cls is not null) _cache.UpsertClass(cls);
            _notify.NotifyMailItem(action, row.MailItemId, _cache.ToBusItem(row, cls));
        }
        catch (Exception ex) { _log.LogWarning(ex, "[MailWB] cache/publish failed for {Id}", row.MailItemId); }
    }

    /// <summary>Update only the classification of an already-cached item, then broadcast.</summary>
    private void CacheAndPublishClass(string action, string mailItemId, MailClassificationResult r, int version)
    {
        try
        {
            var cls = ToClassRow(mailItemId, r, version);
            _cache.UpsertClass(cls);
            _notify.NotifyMailItem(action, mailItemId,
                _cache.TryGetItem(mailItemId, out var row) ? _cache.ToBusItem(row, cls) : null);
        }
        catch (Exception ex) { _log.LogWarning(ex, "[MailWB] cache/publish (class) failed for {Id}", mailItemId); }
    }

    /// <summary>
    /// Uploads an item's attachments + the raw .eml to the SP document library under
    /// MailArchive/{category}/{yyyy-MM}/{mailItemId}/, then patches the item's manifest with the
    /// resulting webUrls + the .eml pointer. Best-effort per file. (Reclassify does NOT relocate
    /// already-stored files; the webUrl pointer stays valid regardless of the folder.)
    /// </summary>
    private async Task StoreFilesAsync(string watchedUpn, MailboxMessageBody body, string mailItemId,
        string category, string receivedAtIso, CancellationToken ct)
    {
        var folder = ItemFolder(category, receivedAtIso, body.Subject, mailItemId);
        Directory.CreateDirectory(folder);

        var manifest = new List<MailAttManifest>();
        foreach (var att in body.Attachments)
        {
            try
            {
                var dl = await _bridge.GetAttachmentAsync(watchedUpn, body.Id, att.Name, ct);
                if (dl is null) { manifest.Add(new MailAttManifest { Name = att.Name, ContentType = att.ContentType, Size = att.Size }); continue; }
                var path = Path.Combine(folder, Sanitize(att.Name));
                await File.WriteAllBytesAsync(path, dl.Value.Bytes, ct);
                manifest.Add(new MailAttManifest { Name = att.Name, ContentType = dl.Value.ContentType, Size = dl.Value.Bytes.Length, WebUrl = path });
            }
            catch (Exception ex)
            {
                _log.LogWarning(ex, "[MailWB] attachment write failed: {Name}", att.Name);
                manifest.Add(new MailAttManifest { Name = att.Name, ContentType = att.ContentType, Size = att.Size });
            }
        }

        string? emlPath = null;
        try
        {
            var eml = await _bridge.GetRawEmlAsync(watchedUpn, body.Id, ct);
            if (eml is not null) { emlPath = Path.Combine(folder, $"{mailItemId}.eml"); await File.WriteAllBytesAsync(emlPath, eml, ct); }
        }
        catch (Exception ex) { _log.LogWarning(ex, "[MailWB] eml write failed for {Id}", mailItemId); }

        await _sp.UpdateMailItemFilesAsync(mailItemId, JsonSerializer.Serialize(manifest), emlPath, ct);
        _log.LogInformation("[MailWB] stored {N} file(s) + eml for {Id} under {Folder}",
            manifest.Count(m => m.WebUrl is not null), mailItemId, folder);
    }

    /// <summary>
    /// Moves an item's stored files from its old taxonomy folder to the new one (on reclassify/amend)
    /// and rewrites the manifest/eml paths. Best-effort — no-op if the source folder doesn't exist.
    /// </summary>
    private async Task MoveItemFilesAsync(string mailItemId, string oldCategory, string newCategory, CancellationToken ct)
    {
        if (string.IsNullOrEmpty(oldCategory) || string.Equals(oldCategory, newCategory, StringComparison.OrdinalIgnoreCase)) return;

        var item = (await _sp.ReadMailItemsAsync(ct)).FirstOrDefault(i => i.MailItemId == mailItemId);
        if (item is null) return;
        var manifest = string.IsNullOrWhiteSpace(item.AttachmentsJson)
            ? new List<MailAttManifest>()
            : (JsonSerializer.Deserialize<List<MailAttManifest>>(item.AttachmentsJson) ?? []);

        // Source = the item's ACTUAL stored folder (parent of a manifest file) — robust to the
        // naming scheme (legacy GUID folders or the new readable ones). Dest = the readable folder.
        var anyPath = manifest.FirstOrDefault(m => !string.IsNullOrEmpty(m.WebUrl))?.WebUrl;
        var src = !string.IsNullOrEmpty(anyPath) ? Path.GetDirectoryName(anyPath)
            : ItemFolder(oldCategory, item.ReceivedAt, item.Subject, mailItemId);
        var dst = ItemFolder(newCategory, item.ReceivedAt, item.Subject, mailItemId);
        if (string.IsNullOrEmpty(src) || !Directory.Exists(src) || string.Equals(src, dst, StringComparison.OrdinalIgnoreCase)) return;

        try
        {
            Directory.CreateDirectory(Path.GetDirectoryName(dst)!);
            if (Directory.Exists(dst)) Directory.Delete(dst, true);
            Directory.Move(src, dst);

            foreach (var m in manifest)
                if (!string.IsNullOrEmpty(m.WebUrl)) m.WebUrl = m.WebUrl!.Replace(src, dst, StringComparison.OrdinalIgnoreCase);
            var newEml = File.Exists(Path.Combine(dst, $"{mailItemId}.eml")) ? Path.Combine(dst, $"{mailItemId}.eml") : null;
            await _sp.UpdateMailItemFilesAsync(mailItemId, JsonSerializer.Serialize(manifest), newEml, ct);
            _log.LogInformation("[MailWB] moved files {Id}: {Old} -> {New}", mailItemId, oldCategory, newCategory);
        }
        catch (Exception ex) { _log.LogWarning(ex, "[MailWB] could not move files for {Id}", mailItemId); }
    }

    private static string Sanitize(string name)
    {
        foreach (var c in System.IO.Path.GetInvalidFileNameChars()) name = name.Replace(c, '_');
        return name.Replace('/', '_').Replace('\\', '_').Trim();
    }

    /// <summary>
    /// Backfills doc-library storage for already-captured items that have attachments but no stored
    /// webUrls yet (e.g. items captured before 1b-ii). Background; shares the seed-progress tracker.
    /// </summary>
    public SeedSnapshot StartStoreAttachments(string watchedUpn, int maxConcurrency = 3)
    {
        if (!_seed.TryBegin()) return _seed.SnapshotNow();
        _log.LogInformation("[MailWB] store-attachments backfill started");

        _ = Task.Run(async () =>
        {
            var items    = (await _sp.ReadMailItemsAsync()).Where(i => i.HasAttachments).ToList();
            var todo     = items.Where(i => !ManifestHasUrls(i.AttachmentsJson)).ToList();
            var currents = (await _sp.ReadCurrentClassificationsAsync())
                .ToDictionary(c => c.MailItemId, c => c.CategoryPath, StringComparer.Ordinal);
            _seed.SetTotal(todo.Count);

            using var sem = new SemaphoreSlim(maxConcurrency);
            var tasks = todo.Select(async i =>
            {
                await sem.WaitAsync();
                try
                {
                    var manifest = JsonSerializer.Deserialize<List<MailAttManifest>>(i.AttachmentsJson) ?? [];
                    var body = new MailboxMessageBody
                    {
                        Id          = i.WrapperGraphId,
                        ReceivedAt  = i.ReceivedAt,
                        Attachments = manifest.Select(m => new MailboxAttachmentMeta { Name = m.Name, ContentType = m.ContentType, Size = m.Size }).ToList(),
                    };
                    var cat = currents.TryGetValue(i.MailItemId, out var c) ? c : "Unclassified";
                    await StoreFilesAsync(watchedUpn, body, i.MailItemId, cat, i.ReceivedAt, CancellationToken.None);
                    _seed.RecordClassified(cat);
                }
                catch (Exception ex) { _seed.RecordFailed(ex.Message); _log.LogWarning(ex, "[MailWB] store-attachments item failed"); }
                finally { sem.Release(); }
            }).ToArray();
            await Task.WhenAll(tasks);
            _seed.End();
            _log.LogInformation("[MailWB] store-attachments done: {S}", _seed.SnapshotNow().Summary());
        });

        return _seed.SnapshotNow();
    }

    private static bool ManifestHasUrls(string attachmentsJson)
    {
        try
        {
            var m = JsonSerializer.Deserialize<List<MailAttManifest>>(attachmentsJson);
            return m is not null && m.Any(a => !string.IsNullOrEmpty(a.WebUrl));
        }
        catch { return false; }
    }

    /// <summary>Re-run classification on a stored item — writes a NEW version, never mutates the email.</summary>
    public async Task<CaptureResult> ReclassifyAsync(string mailItemId, CancellationToken ct = default)
    {
        var input = await _sp.GetClassifyInputAsync(mailItemId, ct)
            ?? throw new InvalidOperationException($"MailItem '{mailItemId}' not found.");
        var priorCat = (await _sp.ReadClassificationsForItemAsync(mailItemId, ct))
            .OrderByDescending(c => c.Version).FirstOrDefault()?.CategoryPath;
        var result = await _classifier.ClassifyAsync(input, ct)
            ?? throw new InvalidOperationException("Both AI providers unavailable.");
        var version = await _sp.WriteClassificationAsync(mailItemId, result, ct);
        await MoveItemFilesAsync(mailItemId, priorCat ?? "", result.Category, ct);
        CacheAndPublishClass("Classified", mailItemId, result, version);
        return new CaptureResult
        {
            MailItemId = mailItemId, IsNew = false, Classified = true,
            Category = result.Category, Confidence = result.Confidence, Version = version,
        };
    }

    /// <summary>
    /// Human classification correction (dev-only feature). Writes a corrected classification version
    /// (AiProvider="human", non-destructive) so the item moves in the tree, AND appends an AI-vs-human
    /// feedback record to a local JSONL for prompt-tuning analysis. correctedCategory may be a known
    /// taxonomy path (selected) or free text (proposed new category) — free text is coerced to Other
    /// with the raw label kept, while the feedback log preserves the verbatim correction.
    /// </summary>
    public async Task<CaptureResult> AmendAsync(string mailItemId, string correctedCategory, string? reason, CancellationToken ct = default)
    {
        var history = await _sp.ReadClassificationsForItemAsync(mailItemId, ct);
        var prior   = history.OrderByDescending(c => c.Version).FirstOrDefault();
        var input   = await _sp.GetClassifyInputAsync(mailItemId, ct);

        var requested = (correctedCategory ?? "").Trim().Trim('/');
        var leaves    = await _taxonomy.GetLeavesAsync(ct);
        var matched   = leaves.FirstOrDefault(l => string.Equals(l.Path, requested, StringComparison.OrdinalIgnoreCase));
        var carried   = (prior?.KeywordTags ?? "").Split(',').Select(t => t.Trim()).Where(t => t.Length > 0).ToList();

        string category;
        if (matched is not null)
        {
            category = matched.Path;                               // existing leaf (canonical casing)
        }
        else if (requested.Length == 0 || string.Equals(requested, "Other", StringComparison.OrdinalIgnoreCase))
        {
            category = "Other";
        }
        else
        {
            // A deliberately-typed NEW bucket → promote it to a real taxonomy leaf (SP hint, optimistic
            // in-memory add) so it appears in the tree immediately, then file the item under it.
            category = requested;
            await _taxonomy.AddLeafHintAsync(category, reason, "amend", ct);
            try { Directory.CreateDirectory(Path.Combine(_archiveRoot, category.Replace('/', Path.DirectorySeparatorChar))); } catch { }
        }

        var result = new MailClassificationResult
        {
            Category    = category,
            OtherLabel  = null,
            Confidence  = 1.0,
            Keywords    = carried,
            Reasoning   = reason,
            AiProvider  = "human",
            AiModel     = "manual-amend",
            RawResponse = JsonSerializer.Serialize(new { correctedCategory, reason }),
        };
        var version = await _sp.WriteClassificationAsync(mailItemId, result, ct);
        await MoveItemFilesAsync(mailItemId, prior?.CategoryPath ?? "", category, ct);
        CacheAndPublishClass("Amended", mailItemId, result, version);

        AppendFeedback(new Dictionary<string, object?>
        {
            ["ts"]                 = DateTimeOffset.UtcNow.ToString("o"),
            ["mailItemId"]         = mailItemId,
            ["subject"]            = input?.Subject,
            ["from"]               = input?.FromAddress,
            ["aiCategory"]         = prior?.CategoryPath,
            ["aiConfidence"]       = prior?.Confidence,
            ["aiModel"]            = prior?.AiModel,
            ["correctedCategory"]  = correctedCategory,
            ["storedCategory"]     = category,
            ["reason"]             = reason,
        });

        _log.LogInformation("[MailWB] AMEND {Id}: {From} -> {To} (reason: {Reason})",
            mailItemId, prior?.CategoryPath, category, reason);
        return new CaptureResult { MailItemId = mailItemId, Classified = true, Category = category, Confidence = 1.0, Version = version };
    }

    /// <summary>
    /// Promote an emergent "Other" suggestion to a first-class taxonomy leaf: writes a SP hint
    /// (custom leaf + optional guidance) so the classifier targets it from the next call onward —
    /// no deploy — then re-files the originating item onto the new leaf. <paramref name="categoryPath"/>
    /// is the confirmed full path (e.g. "Corporate/Insurance"); <paramref name="hint"/> is the
    /// classifier guidance (defaults to the AI's proposed label/reasoning when omitted).
    /// </summary>
    public async Task<CaptureResult> ConfirmLeafAsync(string mailItemId, string categoryPath, string? hint, CancellationToken ct = default)
    {
        categoryPath = (categoryPath ?? "").Trim().Trim('/');
        if (categoryPath.Length == 0) throw new ArgumentException("categoryPath is required.");
        await _taxonomy.AddLeafHintAsync(categoryPath, hint, "confirm-leaf", ct);
        Directory.CreateDirectory(Path.Combine(_archiveRoot, categoryPath.Replace('/', Path.DirectorySeparatorChar)));
        _log.LogInformation("[MailWB] CONFIRM-LEAF {Path} (hint: {Hint}) from item {Id}", categoryPath, hint, mailItemId);
        // Re-file the source item onto the now-valid leaf (reason records the promotion).
        return await AmendAsync(mailItemId, categoryPath, $"Confirmed emergent leaf '{categoryPath}'" + (string.IsNullOrWhiteSpace(hint) ? "" : $": {hint}"), ct);
    }

    private static readonly object _feedbackLock = new();

    /// <summary>Path to the local prompt-tuning feedback log (dev machine; survives reinstall).</summary>
    public static string FeedbackPath => System.IO.Path.Combine(
        Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
        "ShredderData", "mail-classification-feedback.jsonl");

    private void AppendFeedback(Dictionary<string, object?> entry)
    {
        try
        {
            var path = FeedbackPath;
            System.IO.Directory.CreateDirectory(System.IO.Path.GetDirectoryName(path)!);
            var line = JsonSerializer.Serialize(entry) + Environment.NewLine;
            lock (_feedbackLock) System.IO.File.AppendAllText(path, line);
        }
        catch (Exception ex) { _log.LogWarning(ex, "[MailWB] could not append classification feedback"); }
    }

    /// <summary>Reads the local feedback log (most-recent first) for review/analysis.</summary>
    public List<Dictionary<string, JsonElement>> ReadFeedback(int max = 500)
    {
        var path = FeedbackPath;
        if (!System.IO.File.Exists(path)) return [];
        var lines = System.IO.File.ReadAllLines(path);
        var result = new List<Dictionary<string, JsonElement>>();
        foreach (var l in lines.Reverse().Take(max))
        {
            if (string.IsNullOrWhiteSpace(l)) continue;
            try { var d = JsonSerializer.Deserialize<Dictionary<string, JsonElement>>(l); if (d is not null) result.Add(d); } catch { }
        }
        return result;
    }

    /// <summary>
    /// Reads the item + current-classification snapshot. Served from the in-memory MailCache (no SP
    /// roundtrip) once warmed; falls back to a direct SP read on a cold cache and kicks off a warm.
    /// </summary>
    private async Task<(List<MailItemRow> Items, List<MailClassRow> Currents)> GetSnapshotAsync(CancellationToken ct)
    {
        if (_cache.Warmed)
            return (_cache.GetItems().ToList(), _cache.GetCurrents().ToList());
        // Cold cache (no disk snapshot yet): one direct SP read, and warm in the background.
        _ = Task.Run(() => _cache.ForceRefreshAsync(CancellationToken.None));
        return (await _sp.ReadMailItemsAsync(ct), await _sp.ReadCurrentClassificationsAsync(ct));
    }

    /// <summary>Classification tree: every taxonomy leaf with total/open/completed counts.</summary>
    public async Task<List<TreeTop>> GetTreeAsync(CancellationToken ct = default)
    {
        var (items, currents) = await GetSnapshotAsync(ct);
        var completedById = items.ToDictionary(i => i.MailItemId, i => i.Completed, StringComparer.Ordinal);

        // count per CategoryPath
        var counts = new Dictionary<string, (int Total, int Completed)>(StringComparer.OrdinalIgnoreCase);
        foreach (var c in currents)
        {
            completedById.TryGetValue(c.MailItemId, out var done);
            counts.TryGetValue(c.CategoryPath, out var cur);
            counts[c.CategoryPath] = (cur.Total + 1, cur.Completed + (done ? 1 : 0));
        }

        // AI-suggested breakdown of the Other bucket, grouped by the emergent OtherLabel.
        var suggestions = currents
            .Where(c => string.Equals(c.CategoryPath, "Other", StringComparison.OrdinalIgnoreCase) && !string.IsNullOrWhiteSpace(c.OtherLabel))
            .GroupBy(c => c.OtherLabel!.Trim(), StringComparer.OrdinalIgnoreCase)
            .Select(g => (Label: g.Key,
                          Total: g.Count(),
                          Completed: g.Count(c => completedById.TryGetValue(c.MailItemId, out var d) && d)))
            .ToList();

        var leaves = await _taxonomy.GetLeavesAsync(ct);
        var tops = new List<TreeTop>();
        foreach (var top in leaves.Select(l => l.Top).Distinct())
        {
            var node = new TreeTop { Name = top };

            // A leaf with no sub means the top IS a bucket (e.g. "Other") — selectable, no redundant child.
            var selfLeaf = leaves.FirstOrDefault(l => l.Top == top && string.IsNullOrEmpty(l.Sub));
            if (selfLeaf is not null)
            {
                node.Path = selfLeaf.Path;
                counts.TryGetValue(selfLeaf.Path, out var sc);
                node.Total += sc.Total; node.Open += sc.Total - sc.Completed;
            }

            // Real sub-leaves (alphabetical).
            foreach (var leaf in leaves.Where(l => l.Top == top && !string.IsNullOrEmpty(l.Sub))
                                       .OrderBy(l => l.Sub, StringComparer.OrdinalIgnoreCase))
            {
                counts.TryGetValue(leaf.Path, out var c);
                node.Subs.Add(new TreeNode { Name = leaf.Sub, Path = leaf.Path, Total = c.Total, Completed = c.Completed, Open = c.Total - c.Completed });
                node.Total += c.Total; node.Open += c.Total - c.Completed;
            }

            // Virtual AI-suggested sub-leaves under Other (muted; subset of the Other bucket, so not added
            // to node.Total). Skipped once a real "Other/<label>" leaf exists or the label has no items.
            if (string.Equals(top, "Other", StringComparison.OrdinalIgnoreCase))
            {
                foreach (var s in suggestions.OrderBy(s => s.Label, StringComparer.OrdinalIgnoreCase))
                {
                    if (s.Total == 0) continue;
                    if (leaves.Any(l => string.Equals(l.Path, $"Other/{s.Label}", StringComparison.OrdinalIgnoreCase))) continue;
                    node.Subs.Add(new TreeNode
                    {
                        Name = s.Label, Path = $"Other:{s.Label}",
                        Total = s.Total, Completed = s.Completed, Open = s.Total - s.Completed,
                        Suggested = true,
                    });
                }
            }
            tops.Add(node);
        }
        return tops;
    }

    /// <summary>Items currently classified under a taxonomy path.</summary>
    public async Task<List<WorkbenchItem>> GetItemsAsync(string category, bool includeCompleted, CancellationToken ct = default)
    {
        var (items, allCurrents) = await GetSnapshotAsync(ct);

        // Virtual suggested bucket "Other:<label>" → Other items whose emergent AI label matches.
        var suggLabel = category.StartsWith("Other:", StringComparison.Ordinal) ? category["Other:".Length..] : null;
        var currents = allCurrents
            .Where(c => suggLabel is not null
                ? string.Equals(c.CategoryPath, "Other", StringComparison.OrdinalIgnoreCase)
                    && string.Equals(c.OtherLabel?.Trim(), suggLabel, StringComparison.OrdinalIgnoreCase)
                : string.Equals(c.CategoryPath, category, StringComparison.OrdinalIgnoreCase))
            .ToDictionary(c => c.MailItemId, c => c, StringComparer.Ordinal);

        var list = new List<WorkbenchItem>();
        foreach (var i in items)
        {
            if (!currents.TryGetValue(i.MailItemId, out var c)) continue;
            if (!includeCompleted && i.Completed) continue;
            list.Add(new WorkbenchItem
            {
                MailItemId = i.MailItemId, Subject = i.Subject, FromAddress = i.FromAddress, FromName = i.FromName,
                ReceivedAt = i.ReceivedAt, HasAttachments = i.HasAttachments, Completed = i.Completed,
                CategoryPath = c.CategoryPath, OtherLabel = c.OtherLabel, Confidence = c.Confidence,
                KeywordTags = c.KeywordTags, PoNumber = c.PoNumber, SoNumber = c.SoNumber, Version = c.Version,
                IsRead = i.IsRead, ClaimedBy = i.ClaimedBy, ClaimedAt = i.ClaimedAt, CompletedBy = i.CompletedBy,
            });
        }
        return list.OrderByDescending(x => x.ReceivedAt, StringComparer.Ordinal).ToList();
    }

    /// <summary>Marks an item complete/incomplete in SP, updates the cache, and broadcasts a "Completed" bus event.</summary>
    public async Task<bool> CompleteAsync(string mailItemId, bool completed, string? by, CancellationToken ct = default)
    {
        var ok = await _sp.SetMailCompletedAsync(mailItemId, completed, by, ct);
        if (!ok) return false;
        var nowIso = DateTimeOffset.UtcNow.ToString("o");
        _cache.SetCompleted(mailItemId, completed, nowIso, by);
        _notify.NotifyMailItem("Completed", mailItemId,
            _cache.TryGetItem(mailItemId, out var row) ? _cache.ToBusItem(row, _cache.GetClass(mailItemId)) : null);
        return true;
    }

    /// <summary>Sets the global read flag (read-by-anyone), updates the cache, and broadcasts a "Read" bus event.</summary>
    public async Task<bool> MarkReadAsync(string mailItemId, bool read, string? by, CancellationToken ct = default)
    {
        var ok = await _sp.SetMailReadAsync(mailItemId, read, by, ct);
        if (!ok) return false;
        _cache.SetRead(mailItemId, read, by);
        _notify.NotifyMailItem("Read", mailItemId,
            _cache.TryGetItem(mailItemId, out var row) ? _cache.ToBusItem(row, _cache.GetClass(mailItemId)) : null);
        return true;
    }

    /// <summary>Full item detail for the viewer: HTML body (from the archived .eml), headers, attachments, classification, state.</summary>
    public async Task<MailItemDetail?> GetItemDetailAsync(string mailItemId, CancellationToken ct = default)
    {
        var d = await _sp.ReadMailItemDetailAsync(mailItemId, ct);
        if (d is null) return null;

        var manifest = ParseManifest(d.AttachmentsJson);
        string? html = null; var isHtml = false; var text = d.BodyText;
        if (!string.IsNullOrEmpty(d.RawEmlUrl) && File.Exists(d.RawEmlUrl))
        {
            try
            {
                var msg = MimeKit.MimeMessage.Load(d.RawEmlUrl);
                if (!string.IsNullOrWhiteSpace(msg.HtmlBody)) { html = msg.HtmlBody; isHtml = true; }
                else if (!string.IsNullOrWhiteSpace(msg.TextBody)) { text = msg.TextBody; }
            }
            catch (Exception ex) { _log.LogWarning(ex, "[MailWB] could not parse .eml for {Id}", mailItemId); }
        }

        var cls = _cache.GetClass(mailItemId);
        return new MailItemDetail
        {
            MailItemId = d.MailItemId, Subject = d.Subject, FromAddress = d.FromAddress, FromName = d.FromName,
            ToLine = d.ToLine, CcLine = d.CcLine, ReceivedAt = d.ReceivedAt,
            Html = html, IsHtml = isHtml, BodyText = text,
            HasAttachments = d.HasAttachments, Completed = d.Completed, CompletedAt = d.CompletedAt, CompletedBy = d.CompletedBy,
            IsRead = d.IsRead, ReadAt = d.ReadAt, ReadBy = d.ReadBy, ClaimedBy = d.ClaimedBy, ClaimedAt = d.ClaimedAt,
            CategoryPath = cls?.CategoryPath ?? "Other", OtherLabel = cls?.OtherLabel, Confidence = cls?.Confidence ?? 0,
            KeywordTags = cls?.KeywordTags ?? "", PoNumber = cls?.PoNumber, SoNumber = cls?.SoNumber,
            Reasoning = null, AiProvider = cls?.AiProvider, AiModel = cls?.AiModel, Version = cls?.Version ?? 0,
            Attachments = manifest.Select(m => new MailAttachmentInfo
            {
                Name = m.Name, ContentType = m.ContentType, Size = m.Size, Stored = !string.IsNullOrEmpty(m.WebUrl),
            }).ToList(),
        };
    }

    /// <summary>Returns a stored attachment's bytes, re-rooted under this machine's archive root (OneDrive-synced).</summary>
    public async Task<(byte[] Bytes, string ContentType, string Name)?> GetItemAttachmentAsync(string mailItemId, string name, CancellationToken ct = default)
    {
        var d = await _sp.ReadMailItemDetailAsync(mailItemId, ct);
        if (d is null) return null;
        var m = ParseManifest(d.AttachmentsJson).FirstOrDefault(a => string.Equals(a.Name, name, StringComparison.OrdinalIgnoreCase));
        if (m is null || string.IsNullOrEmpty(m.WebUrl)) return null;
        var path = ReRootToLocalArchive(m.WebUrl);
        if (!File.Exists(path)) return null;
        var bytes = await File.ReadAllBytesAsync(path, ct);
        return (bytes, string.IsNullOrEmpty(m.ContentType) ? "application/octet-stream" : m.ContentType, m.Name);
    }

    private static List<MailAttManifest> ParseManifest(string? json)
    {
        if (string.IsNullOrWhiteSpace(json)) return [];
        try { return JsonSerializer.Deserialize<List<MailAttManifest>>(json) ?? []; } catch { return []; }
    }

    /// <summary>Re-roots a stored absolute path (possibly from another machine/user) under THIS machine's archive root.</summary>
    private string ReRootToLocalArchive(string storedPath)
    {
        var marker = $"{Path.DirectorySeparatorChar}Shredder{Path.DirectorySeparatorChar}";
        var idx = storedPath.IndexOf(marker, StringComparison.OrdinalIgnoreCase);
        if (idx < 0) return storedPath;   // unknown shape — try as-is
        var rel = storedPath[(idx + marker.Length)..];
        return Path.Combine(_archiveRoot, rel);
    }

    // ── Bulk seed (capture+classify every surfaced message) ──────────────────────────

    public SeedSnapshot GetSeedStatus() => _seed.SnapshotNow();

    /// <summary>
    /// Safety-net dedup sweep: removes MailItems that share a WrapperGraphId (rare cross-proxy claim
    /// race), keeping the earliest, deleting the others + their classifications. Returns count removed.
    /// </summary>
    public async Task<int> DedupMailItemsAsync(CancellationToken ct = default)
    {
        var items   = await _sp.ReadMailItemsAsync(ct);
        var removed = 0;
        foreach (var grp in items.Where(i => !string.IsNullOrEmpty(i.WrapperGraphId))
                                  .GroupBy(i => i.WrapperGraphId, StringComparer.Ordinal)
                                  .Where(g => g.Count() > 1))
        {
            var keep = grp.OrderBy(i => int.TryParse(i.SpId, out var n) ? n : int.MaxValue).First();
            foreach (var dup in grp.Where(i => i.SpId != keep.SpId))
            {
                try
                {
                    await _sp.DeleteClassificationsForItemAsync(dup.MailItemId, ct);
                    await _sp.DeleteMailItemAsync(dup.SpId, ct);
                    _cache.Remove(dup.MailItemId);
                    _notify.NotifyMailItem("Deleted", dup.MailItemId, null);
                    removed++;
                }
                catch (Exception ex) { _log.LogWarning(ex, "[MailWB] dedup delete failed for {Id}", dup.MailItemId); }
            }
        }
        if (removed > 0) _log.LogInformation("[MailWB] dedup removed {N} duplicate MailItem(s)", removed);
        return removed;
    }

    /// <summary>
    /// Full reset (dev): deletes all MailItems + classifications, clears the wrapper claim/processed
    /// categories (so everything re-surfaces), and removes the stored taxonomy folders (keeps Inbox).
    /// Follow with a backfill to re-import fresh (Message-IDs populated, readable folders, multi-.eml).
    /// </summary>
    public async Task<object> PurgeAsync(string watchedUpn, CancellationToken ct = default)
    {
        var items = await _sp.ReadMailItemsAsync(ct);
        foreach (var i in items)
        {
            try { await _sp.DeleteClassificationsForItemAsync(i.MailItemId, ct); await _sp.DeleteMailItemAsync(i.SpId, ct); }
            catch (Exception ex) { _log.LogWarning(ex, "[MailWB] purge delete failed {Id}", i.MailItemId); }
        }
        var cleared = await _bridge.ResetClaimCategoriesAsync(watchedUpn, ct);
        _cache.ClearAll();

        int dirs = 0;
        try
        {
            if (Directory.Exists(_archiveRoot))
                foreach (var dir in Directory.EnumerateDirectories(_archiveRoot))
                {
                    if (Path.GetFileName(dir).Equals("Inbox", StringComparison.OrdinalIgnoreCase)) continue;
                    try { Directory.Delete(dir, true); dirs++; } catch (Exception ex) { _log.LogWarning(ex, "[MailWB] purge folder delete failed {Dir}", dir); }
                }
        }
        catch (Exception ex) { _log.LogWarning(ex, "[MailWB] purge folder enumerate failed"); }

        _log.LogInformation("[MailWB] purge: {Items} items, {Cats} categories cleared, {Dirs} folders", items.Count, cleared, dirs);
        return new { itemsDeleted = items.Count, categoriesCleared = cleared, foldersDeleted = dirs };
    }

    /// <summary>
    /// One auto-capture pass (used by MailAutoCaptureService): capture+classify any new surfaced
    /// message not yet in SP. Quiet (no seed-progress tracker); dedup via the in-memory wrapper-id
    /// set + the cross-proxy claim.
    /// </summary>
    public async Task AutoCaptureCycleAsync(string watchedUpn, CancellationToken ct)
    {
        var headers = _bridge.GetMessages(watchedUpn, 250) ?? [];
        if (headers.Count == 0) return;
        var seen   = await BuildSeenAsync(ct);
        var claims = new ConcurrentDictionary<string, Lazy<Task<bool>>>(StringComparer.Ordinal);
        foreach (var h in headers)
        {
            if (ct.IsCancellationRequested) break;
            try { await CaptureAndClassifyAsync(watchedUpn, h.Id, seen, claims, ct); }
            catch (Exception ex) { _log.LogWarning(ex, "[MailWB] auto-capture item failed"); }
        }
    }

    /// <summary>
    /// Starts a background pass that captures+classifies every message currently surfaced by the
    /// bridge (idempotent — dedup skips already-captured items). Returns immediately; poll
    /// GetSeedStatus for progress. No-op if a pass is already running.
    /// </summary>
    public SeedSnapshot StartCaptureAll(string watchedUpn, int maxConcurrency = 4)
    {
        if (!_seed.TryBegin()) return _seed.SnapshotNow();

        var headers = _bridge.GetMessages(watchedUpn, 250) ?? [];
        _seed.SetTotal(headers.Count);
        _log.LogInformation("[MailWB] capture-all started: {N} surfaced messages", headers.Count);

        _ = Task.Run(async () =>
        {
            var seen   = await BuildSeenAsync(CancellationToken.None);
            var claims = new ConcurrentDictionary<string, Lazy<Task<bool>>>(StringComparer.Ordinal);
            using var sem = new SemaphoreSlim(maxConcurrency);
            var tasks = headers.Select(async h =>
            {
                await sem.WaitAsync();
                try
                {
                    var r = await CaptureAndClassifyAsync(watchedUpn, h.Id, seen, claims);
                    if (!r.IsNew) _seed.RecordExisting();
                    else if (r.Classified) _seed.RecordClassified(r.Category ?? "Other");
                    else _seed.RecordFailed("classify returned null");
                }
                catch (Exception ex) { _seed.RecordFailed(ex.Message); _log.LogWarning(ex, "[MailWB] seed item failed"); }
                finally { sem.Release(); }
            }).ToArray();
            await Task.WhenAll(tasks);
            _seed.End();
            _log.LogInformation("[MailWB] capture-all done: {S}", _seed.SnapshotNow().Summary());
        });

        return _seed.SnapshotNow();
    }

    /// <summary>
    /// Full-folder backfill: paginates the ENTIRE mirror folder (not just the polled top-50 window)
    /// and captures+classifies every forward-as-attachment item. Idempotent (in-memory dedup).
    /// Background; shares the seed-progress tracker. Use to ingest the complete dump history.
    /// </summary>
    public SeedSnapshot StartBackfill(string watchedUpn, int maxConcurrency = 4)
    {
        if (!_seed.TryBegin()) return _seed.SnapshotNow();
        _log.LogInformation("[MailWB] backfill started for {Upn}", watchedUpn);

        _ = Task.Run(async () =>
        {
            List<MailboxMessageBody> bodies;
            try { bodies = await _bridge.EnumerateAllForwardedAsync(watchedUpn); }
            catch (Exception ex) { _seed.RecordFailed(ex.Message); _seed.End(); _log.LogWarning(ex, "[MailWB] backfill enumerate failed"); return; }

            _seed.SetTotal(bodies.Count);
            var seen   = await BuildSeenAsync(CancellationToken.None);
            var claims = new ConcurrentDictionary<string, Lazy<Task<bool>>>(StringComparer.Ordinal);

            using var sem = new SemaphoreSlim(maxConcurrency);
            var tasks = bodies.Select(async b =>
            {
                await sem.WaitAsync();
                try
                {
                    var r = await CaptureBodyAsync(watchedUpn, b, seen, claims);
                    if (!r.IsNew) _seed.RecordExisting();
                    else if (r.Classified) _seed.RecordClassified(r.Category ?? "Other");
                    else _seed.RecordFailed("classify returned null");
                }
                catch (Exception ex) { _seed.RecordFailed(ex.Message); _log.LogWarning(ex, "[MailWB] backfill item failed"); }
                finally { sem.Release(); }
            }).ToArray();
            await Task.WhenAll(tasks);
            _seed.End();
            _log.LogInformation("[MailWB] backfill done: {S}", _seed.SnapshotNow().Summary());
        });

        return _seed.SnapshotNow();
    }

    /// <summary>
    /// Re-runs classification over EVERY stored MailItem with the current taxonomy (each writes a
    /// new version, non-destructive). Use after a taxonomy change so existing items resweep into new
    /// leaves. Background; shares the seed-progress tracker. No-op if a pass is already running.
    /// </summary>
    public SeedSnapshot StartReclassifyAll(int maxConcurrency = 4)
    {
        if (!_seed.TryBegin()) return _seed.SnapshotNow();
        _log.LogInformation("[MailWB] reclassify-all started");

        _ = Task.Run(async () =>
        {
            List<MailItemRow> items;
            try { items = await _sp.ReadMailItemsAsync(); }
            catch (Exception ex) { _seed.RecordFailed(ex.Message); _seed.End(); _log.LogWarning(ex, "[MailWB] reclassify-all: read items failed"); return; }

            _seed.SetTotal(items.Count);
            using var sem = new SemaphoreSlim(maxConcurrency);
            var tasks = items.Select(async it =>
            {
                await sem.WaitAsync();
                try
                {
                    var r = await ReclassifyAsync(it.MailItemId);
                    if (r.Classified) _seed.RecordClassified(r.Category ?? "Other");
                    else _seed.RecordFailed("classify returned null");
                }
                catch (Exception ex) { _seed.RecordFailed(ex.Message); _log.LogWarning(ex, "[MailWB] reclassify-all item failed"); }
                finally { sem.Release(); }
            }).ToArray();
            await Task.WhenAll(tasks);
            _seed.End();
            _log.LogInformation("[MailWB] reclassify-all done: {S}", _seed.SnapshotNow().Summary());
        });

        return _seed.SnapshotNow();
    }

    public sealed class SeedProgress
    {
        private readonly object _lock = new();
        private bool _running;
        private int _total, _classified, _existing, _failed;
        private readonly Dictionary<string, int> _byCat = new(StringComparer.OrdinalIgnoreCase);
        private readonly List<string> _errors = new();

        public bool TryBegin()
        {
            lock (_lock)
            {
                if (_running) return false;
                _running = true; _total = _classified = _existing = _failed = 0;
                _byCat.Clear(); _errors.Clear();
                return true;
            }
        }
        public void SetTotal(int t)            { lock (_lock) _total = t; }
        public void RecordClassified(string c) { lock (_lock) { _classified++; _byCat[c] = _byCat.TryGetValue(c, out var v) ? v + 1 : 1; } }
        public void RecordExisting()           { lock (_lock) _existing++; }
        public void RecordFailed(string e)     { lock (_lock) { _failed++; if (_errors.Count < 10) _errors.Add(e); } }
        public void End()                      { lock (_lock) _running = false; }

        public SeedSnapshot SnapshotNow()
        {
            lock (_lock) return new SeedSnapshot
            {
                Running = _running, Total = _total, Classified = _classified,
                Existing = _existing, Failed = _failed,
                Processed = _classified + _existing + _failed,
                ByCategory = new Dictionary<string, int>(_byCat), Errors = new List<string>(_errors),
            };
        }
    }

    public sealed class SeedSnapshot
    {
        public bool Running { get; set; }
        public int  Total { get; set; }
        public int  Processed { get; set; }
        public int  Classified { get; set; }
        public int  Existing { get; set; }
        public int  Failed { get; set; }
        public Dictionary<string, int> ByCategory { get; set; } = new();
        public List<string> Errors { get; set; } = new();
        public string Summary() => $"{Processed}/{Total} classified={Classified} existing={Existing} failed={Failed}";
    }

    public sealed class CaptureResult
    {
        public string MailItemId { get; set; } = "";
        public bool   IsNew      { get; set; }
        public bool   Classified { get; set; }
        public string? Category  { get; set; }
        public double Confidence { get; set; }
        public int    Version    { get; set; }
    }

    public sealed class TreeTop
    {
        public string Name { get; set; } = "";
        /// <summary>Set when the top is itself a selectable bucket (a leaf with no sub, e.g. "Other").</summary>
        public string? Path { get; set; }
        public int Total { get; set; }
        public int Open  { get; set; }
        public List<TreeNode> Subs { get; set; } = [];
    }

    public sealed class TreeNode
    {
        public string Name { get; set; } = "";
        public string Path { get; set; } = "";
        public int Total { get; set; }
        public int Open  { get; set; }
        public int Completed { get; set; }
        /// <summary>AI-proposed Other sub-label not yet confirmed as a real leaf — rendered muted in the UI.</summary>
        public bool Suggested { get; set; }
    }

    public sealed class WorkbenchItem
    {
        public string  MailItemId   { get; set; } = "";
        public string  Subject      { get; set; } = "";
        public string  FromAddress  { get; set; } = "";
        public string  FromName     { get; set; } = "";
        public string  ReceivedAt   { get; set; } = "";
        public bool    HasAttachments { get; set; }
        public bool    Completed    { get; set; }
        public string  CategoryPath { get; set; } = "";
        public string? OtherLabel   { get; set; }
        public double  Confidence   { get; set; }
        public string  KeywordTags  { get; set; } = "";
        public string? PoNumber     { get; set; }
        public string? SoNumber     { get; set; }
        public int     Version      { get; set; }
        public bool    IsRead       { get; set; }
        public string? ClaimedBy    { get; set; }
        public string? ClaimedAt    { get; set; }
        public string? CompletedBy  { get; set; }
    }

    public sealed class MailItemDetail
    {
        public string  MailItemId  { get; set; } = "";
        public string  Subject     { get; set; } = "";
        public string  FromAddress { get; set; } = "";
        public string  FromName    { get; set; } = "";
        public string  ToLine      { get; set; } = "";
        public string  CcLine      { get; set; } = "";
        public string  ReceivedAt  { get; set; } = "";
        public string? Html        { get; set; }
        public bool    IsHtml      { get; set; }
        public string  BodyText    { get; set; } = "";
        public bool    HasAttachments { get; set; }
        public bool    Completed   { get; set; }
        public string? CompletedAt { get; set; }
        public string? CompletedBy { get; set; }
        public bool    IsRead      { get; set; }
        public string? ReadAt      { get; set; }
        public string? ReadBy      { get; set; }
        public string? ClaimedBy   { get; set; }
        public string? ClaimedAt   { get; set; }
        public string  CategoryPath { get; set; } = "Other";
        public string? OtherLabel  { get; set; }
        public double  Confidence  { get; set; }
        public string  KeywordTags { get; set; } = "";
        public string? PoNumber    { get; set; }
        public string? SoNumber    { get; set; }
        public string? Reasoning   { get; set; }
        public string? AiProvider  { get; set; }
        public string? AiModel     { get; set; }
        public int     Version     { get; set; }
        public List<MailAttachmentInfo> Attachments { get; set; } = [];
    }

    public sealed class MailAttachmentInfo
    {
        public string Name        { get; set; } = "";
        public string ContentType { get; set; } = "";
        public long   Size        { get; set; }
        public bool   Stored      { get; set; }
    }
}
