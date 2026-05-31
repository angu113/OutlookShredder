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
    private readonly ILogger<MailWorkbenchService> _log;
    private readonly SeedProgress _seed = new();
    private readonly string _archiveRoot;

    public MailWorkbenchService(SharePointService sp, MailClassifierService classifier,
        MailboxBridgeService bridge, IConfiguration config, ILogger<MailWorkbenchService> log)
    {
        _sp = sp; _classifier = classifier; _bridge = bridge; _log = log;

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

    /// <summary>Local folder for an item's stored files: {root}\{category}\{mailItemId}\</summary>
    private string ItemFolder(string category, string mailItemId) =>
        Path.Combine(_archiveRoot, category.Replace('/', Path.DirectorySeparatorChar), mailItemId);

    /// <summary>Capture a cached bridge message → MailItems (dedup) → classify → MailClassifications.</summary>
    public async Task<CaptureResult> CaptureAndClassifyAsync(string watchedUpn, string wrapperId,
        ConcurrentDictionary<string, byte>? seen = null, CancellationToken ct = default)
    {
        var body = _bridge.GetMessage(watchedUpn, wrapperId)
            ?? throw new InvalidOperationException("Message not in bridge cache (not polled yet).");
        return await CaptureBodyAsync(watchedUpn, body, seen, ct);
    }

    /// <summary>Capture+classify a parsed message body directly — used by both the cache path and the full-folder backfill.</summary>
    public async Task<CaptureResult> CaptureBodyAsync(string watchedUpn, MailboxMessageBody body,
        ConcurrentDictionary<string, byte>? seen = null, CancellationToken ct = default)
    {
        // Reliable in-memory dedup on the wrapper id (the SP eq-filter false-matches these long,
        // near-identical Graph ids). Single-call path loads the existing set; bulk passes a shared one.
        if (seen is null)
        {
            var existingIds = await _sp.GetExistingWrapperIdsAsync(ct);
            seen = new ConcurrentDictionary<string, byte>(
                existingIds.Select(id => new KeyValuePair<string, byte>(id, (byte)0)), StringComparer.Ordinal);
        }
        if (!seen.TryAdd(body.Id, 0))
            return new CaptureResult { MailItemId = "", IsNew = false };

        var manifest = body.Attachments
            .Select(a => new MailAttManifest { Name = a.Name, ContentType = a.ContentType, Size = a.Size })
            .ToList();

        var mailItemId = await _sp.WriteMailItemAsync(new MailItemInput
        {
            SourceType      = "email",
            SourceMailbox   = watchedUpn,
            WrapperGraphId  = body.Id,
            FromAddress     = body.FromAddress,
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

        if (body.Attachments.Count > 0)
            await StoreFilesAsync(watchedUpn, body, mailItemId, category, body.ReceivedAt, ct);

        if (result is null)
            return new CaptureResult { MailItemId = mailItemId, IsNew = true, Classified = false };

        var version = await _sp.WriteClassificationAsync(mailItemId, result, ct);
        return new CaptureResult
        {
            MailItemId = mailItemId, IsNew = true, Classified = true,
            Category = result.Category, Confidence = result.Confidence, Version = version,
        };
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
        var folder = ItemFolder(category, mailItemId);
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
        var src = ItemFolder(oldCategory, mailItemId);
        var dst = ItemFolder(newCategory, mailItemId);
        if (!Directory.Exists(src)) return;

        try
        {
            Directory.CreateDirectory(Path.GetDirectoryName(dst)!);
            if (Directory.Exists(dst)) Directory.Delete(dst, true);
            Directory.Move(src, dst);

            // Rewrite stored paths (old folder -> new folder) in the item's manifest + eml pointer.
            var item = (await _sp.ReadMailItemsAsync(ct)).FirstOrDefault(i => i.MailItemId == mailItemId);
            if (item is not null && !string.IsNullOrWhiteSpace(item.AttachmentsJson))
            {
                var manifest = JsonSerializer.Deserialize<List<MailAttManifest>>(item.AttachmentsJson) ?? [];
                foreach (var m in manifest)
                    if (!string.IsNullOrEmpty(m.WebUrl)) m.WebUrl = m.WebUrl.Replace(src, dst, StringComparison.OrdinalIgnoreCase);
                var newEml = File.Exists(Path.Combine(dst, $"{mailItemId}.eml")) ? Path.Combine(dst, $"{mailItemId}.eml") : null;
                await _sp.UpdateMailItemFilesAsync(mailItemId, JsonSerializer.Serialize(manifest), newEml, ct);
            }
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

        var coerced  = MailTaxonomy.Coerce(correctedCategory);
        var isKnown  = MailTaxonomy.ValidPaths.Contains(correctedCategory);
        var carried  = (prior?.KeywordTags ?? "").Split(',').Select(t => t.Trim()).Where(t => t.Length > 0).ToList();

        var result = new MailClassificationResult
        {
            Category    = coerced,
            OtherLabel  = (!isKnown && !string.Equals(correctedCategory, "Other", StringComparison.OrdinalIgnoreCase)) ? correctedCategory.Trim() : null,
            Confidence  = 1.0,
            Keywords    = carried,
            Reasoning   = reason,
            AiProvider  = "human",
            AiModel     = "manual-amend",
            RawResponse = JsonSerializer.Serialize(new { correctedCategory, reason }),
        };
        var version = await _sp.WriteClassificationAsync(mailItemId, result, ct);
        await MoveItemFilesAsync(mailItemId, prior?.CategoryPath ?? "", coerced, ct);

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
            ["storedCategory"]     = coerced,
            ["reason"]             = reason,
        });

        _log.LogInformation("[MailWB] AMEND {Id}: {From} -> {To} (reason: {Reason})",
            mailItemId, prior?.CategoryPath, correctedCategory, reason);
        return new CaptureResult { MailItemId = mailItemId, Classified = true, Category = coerced, Confidence = 1.0, Version = version };
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

    /// <summary>Classification tree: every taxonomy leaf with total/open/completed counts.</summary>
    public async Task<List<TreeTop>> GetTreeAsync(CancellationToken ct = default)
    {
        var currents = await _sp.ReadCurrentClassificationsAsync(ct);
        var items    = await _sp.ReadMailItemsAsync(ct);
        var completedById = items.ToDictionary(i => i.MailItemId, i => i.Completed, StringComparer.Ordinal);

        // count per CategoryPath
        var counts = new Dictionary<string, (int Total, int Completed)>(StringComparer.OrdinalIgnoreCase);
        foreach (var c in currents)
        {
            completedById.TryGetValue(c.MailItemId, out var done);
            counts.TryGetValue(c.CategoryPath, out var cur);
            counts[c.CategoryPath] = (cur.Total + 1, cur.Completed + (done ? 1 : 0));
        }

        var tops = new List<TreeTop>();
        foreach (var top in MailTaxonomy.Leaves.Select(l => l.Top).Distinct())
        {
            var node = new TreeTop { Name = top };
            foreach (var leaf in MailTaxonomy.Leaves.Where(l => l.Top == top))
            {
                counts.TryGetValue(leaf.Path, out var c);
                node.Subs.Add(new TreeNode
                {
                    Name = string.IsNullOrEmpty(leaf.Sub) ? top : leaf.Sub,
                    Path = leaf.Path,
                    Total = c.Total, Completed = c.Completed, Open = c.Total - c.Completed,
                });
            }
            node.Total     = node.Subs.Sum(s => s.Total);
            node.Open      = node.Subs.Sum(s => s.Open);
            tops.Add(node);
        }
        return tops;
    }

    /// <summary>Items currently classified under a taxonomy path.</summary>
    public async Task<List<WorkbenchItem>> GetItemsAsync(string category, bool includeCompleted, CancellationToken ct = default)
    {
        var currents = (await _sp.ReadCurrentClassificationsAsync(ct))
            .Where(c => string.Equals(c.CategoryPath, category, StringComparison.OrdinalIgnoreCase))
            .ToDictionary(c => c.MailItemId, c => c, StringComparer.Ordinal);
        var items = await _sp.ReadMailItemsAsync(ct);

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
            });
        }
        return list.OrderByDescending(x => x.ReceivedAt, StringComparer.Ordinal).ToList();
    }

    // ── Bulk seed (capture+classify every surfaced message) ──────────────────────────

    public SeedSnapshot GetSeedStatus() => _seed.SnapshotNow();

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
            var existingIds = await _sp.GetExistingWrapperIdsAsync();
            var seen = new ConcurrentDictionary<string, byte>(
                existingIds.Select(id => new KeyValuePair<string, byte>(id, (byte)0)), StringComparer.Ordinal);
            using var sem = new SemaphoreSlim(maxConcurrency);
            var tasks = headers.Select(async h =>
            {
                await sem.WaitAsync();
                try
                {
                    var r = await CaptureAndClassifyAsync(watchedUpn, h.Id, seen);
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
            var existingIds = await _sp.GetExistingWrapperIdsAsync();
            var seen = new ConcurrentDictionary<string, byte>(
                existingIds.Select(id => new KeyValuePair<string, byte>(id, (byte)0)), StringComparer.Ordinal);

            using var sem = new SemaphoreSlim(maxConcurrency);
            var tasks = bodies.Select(async b =>
            {
                await sem.WaitAsync();
                try
                {
                    var r = await CaptureBodyAsync(watchedUpn, b, seen);
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
    }
}
