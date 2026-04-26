using System.Text;
using System.Text.Json;
using OutlookShredder.Proxy.Models;

namespace OutlookShredder.Proxy.Services;

/// <summary>
/// Loads the product synonym dictionary from the SharePoint ProductSynonyms list,
/// caches it in memory, and builds the formatted prompt block injected into AI system
/// prompts. Write-through: additions/edits update the in-memory cache and publish a
/// Service Bus message so all Shredder clients update their local caches too.
/// </summary>
public class ProductSynonymService
{
    private readonly ILogger<ProductSynonymService> _log;
    private readonly SharePointService              _sp;
    private readonly RfqNotificationService         _notify;

    private List<SynonymGroup> _groups     = [];
    private string?            _cachedBlock;
    private readonly object    _lock       = new();

    public IReadOnlyList<SynonymGroup> Groups
    {
        get { lock (_lock) return [.. _groups]; }
    }

    public ProductSynonymService(
        ILogger<ProductSynonymService> log,
        SharePointService sp,
        RfqNotificationService notify)
    {
        _log    = log;
        _sp     = sp;
        _notify = notify;
    }

    /// <summary>
    /// Loads synonyms from SharePoint. Called at proxy startup from the prewarm hook.
    /// If SP returns an empty list and a product-synonyms.json seed file exists in the
    /// exe directory, seeds SP automatically from that file (one-time migration).
    /// </summary>
    public async Task LoadAsync(CancellationToken ct = default)
    {
        try
        {
            var groups = await _sp.ReadSynonymsAsync(ct);
            if (groups.Count == 0)
            {
                // One-time seed from JSON file if present
                var seedPath = Path.Combine(AppContext.BaseDirectory, "product-synonyms.json");
                if (File.Exists(seedPath))
                {
                    _log.LogInformation("[Synonyms] SP list empty — seeding from {Path}", seedPath);
                    groups = await SeedFromJsonAsync(seedPath, ct);
                }
            }
            lock (_lock) { _groups = groups; _cachedBlock = null; }
            _log.LogInformation("[Synonyms] Loaded {Count} synonym groups from SP", groups.Count);
        }
        catch (Exception ex)
        {
            _log.LogWarning(ex, "[Synonyms] Failed to load from SP — AI extraction continues without synonym block");
        }
    }

    /// <summary>Returns the formatted prompt block for injection into AI system prompts.</summary>
    public string BuildPromptBlock()
    {
        lock (_lock)
        {
            if (_cachedBlock is not null) return _cachedBlock;
            if (_groups.Count == 0) return _cachedBlock = string.Empty;

            var sb = new StringBuilder();
            sb.AppendLine();
            sb.AppendLine("── PRODUCT TERM SYNONYMS ────────────────────────────────────────────────");
            sb.AppendLine("When matching supplier descriptions to requested products, treat these as equivalent:");
            foreach (var g in _groups)
            {
                if (g.Variants.Length == 0) continue;
                sb.Append("• \"").Append(g.Canonical).Append("\" = ")
                  .AppendLine(string.Join(", ", g.Variants));
            }
            sb.AppendLine("Always use the canonical name when matching, regardless of which variant the supplier used.");
            return _cachedBlock = sb.ToString();
        }
    }

    /// <summary>Adds a new synonym group to SP, updates cache, publishes bus message.</summary>
    public async Task<SynonymGroup> AddSynonymAsync(SynonymGroup group, CancellationToken ct = default)
    {
        var saved = await _sp.WriteSynonymAsync(group, ct);
        AddOrUpdateLocal(saved);
        _notify.NotifySynonym(saved);
        return saved;
    }

    /// <summary>Updates an existing synonym group in SP, updates cache, publishes bus message.</summary>
    public async Task<SynonymGroup> UpdateSynonymAsync(string spItemId, SynonymGroup group, CancellationToken ct = default)
    {
        var saved = await _sp.UpdateSynonymAsync(spItemId, group, ct);
        AddOrUpdateLocal(saved);
        _notify.NotifySynonym(saved);
        return saved;
    }

    /// <summary>Updates only the local in-memory cache. Called when a "Synonym" bus message
    /// arrives on another proxy instance (avoids SP roundtrip — the writer already updated SP).</summary>
    public void AddOrUpdateLocal(SynonymGroup group)
    {
        lock (_lock)
        {
            var idx = _groups.FindIndex(g =>
                g.Canonical.Equals(group.Canonical, StringComparison.OrdinalIgnoreCase));
            if (idx >= 0) _groups[idx] = group;
            else _groups.Add(group);
            _cachedBlock = null;
        }
    }

    // ── Private helpers ───────────────────────────────────────────────────────

    private async Task<List<SynonymGroup>> SeedFromJsonAsync(string path, CancellationToken ct)
    {
        var json   = await File.ReadAllTextAsync(path, ct);
        var parsed = JsonSerializer.Deserialize<JsonSynonymGroup[]>(json,
            new JsonSerializerOptions { PropertyNameCaseInsensitive = true });
        if (parsed is null) return [];

        var results = new List<SynonymGroup>();
        foreach (var p in parsed)
        {
            if (string.IsNullOrWhiteSpace(p.Canonical)) continue;
            var group = new SynonymGroup
            {
                Canonical = p.Canonical,
                Category  = p.Category,
                Variants  = p.Variants ?? [],
            };
            try
            {
                var saved = await _sp.WriteSynonymAsync(group, ct);
                results.Add(saved);
            }
            catch (Exception ex)
            {
                _log.LogWarning(ex, "[Synonyms] Failed to seed '{Canonical}'", p.Canonical);
            }
        }
        _log.LogInformation("[Synonyms] Seeded {Count} groups into SP", results.Count);
        return results;
    }

    private sealed class JsonSynonymGroup
    {
        public string   Canonical { get; set; } = string.Empty;
        public string?  Category  { get; set; }
        public string[]? Variants { get; set; }
    }
}
