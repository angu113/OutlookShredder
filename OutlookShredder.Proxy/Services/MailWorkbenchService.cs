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

    public MailWorkbenchService(SharePointService sp, MailClassifierService classifier,
        MailboxBridgeService bridge, ILogger<MailWorkbenchService> log)
    {
        _sp = sp; _classifier = classifier; _bridge = bridge; _log = log;
    }

    /// <summary>Capture a cached bridge message → MailItems (dedup) → classify → MailClassifications.</summary>
    public async Task<CaptureResult> CaptureAndClassifyAsync(string watchedUpn, string wrapperId, CancellationToken ct = default)
    {
        var body = _bridge.GetMessage(watchedUpn, wrapperId)
            ?? throw new InvalidOperationException("Message not in bridge cache (not polled yet).");

        var manifest = body.Attachments
            .Select(a => new MailAttManifest { Name = a.Name, ContentType = a.ContentType, Size = a.Size })
            .ToList();

        var (mailItemId, isNew) = await _sp.WriteMailItemAsync(new MailItemInput
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

        if (!isNew)
            return new CaptureResult { MailItemId = mailItemId, IsNew = false };

        var input = new MailClassifyInput
        {
            Subject = body.Subject, FromAddress = body.FromAddress, FromName = body.FromName,
            ToLine = body.ToLine, BodyText = body.BodyText,
            AttachmentNames = manifest.Select(a => a.Name).ToList(),
        };
        var result = await _classifier.ClassifyAsync(input, ct);
        if (result is null)
            return new CaptureResult { MailItemId = mailItemId, IsNew = true, Classified = false };

        var version = await _sp.WriteClassificationAsync(mailItemId, result, ct);
        return new CaptureResult
        {
            MailItemId = mailItemId, IsNew = true, Classified = true,
            Category = result.Category, Confidence = result.Confidence, Version = version,
        };
    }

    /// <summary>Re-run classification on a stored item — writes a NEW version, never mutates the email.</summary>
    public async Task<CaptureResult> ReclassifyAsync(string mailItemId, CancellationToken ct = default)
    {
        var input = await _sp.GetClassifyInputAsync(mailItemId, ct)
            ?? throw new InvalidOperationException($"MailItem '{mailItemId}' not found.");
        var result = await _classifier.ClassifyAsync(input, ct)
            ?? throw new InvalidOperationException("Both AI providers unavailable.");
        var version = await _sp.WriteClassificationAsync(mailItemId, result, ct);
        return new CaptureResult
        {
            MailItemId = mailItemId, IsNew = false, Classified = true,
            Category = result.Category, Confidence = result.Confidence, Version = version,
        };
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
