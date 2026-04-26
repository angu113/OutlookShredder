using System.Text;
using System.Text.Json;

namespace OutlookShredder.Proxy.Services;

/// <summary>
/// Loads product-synonyms.json and builds a formatted prompt block for injection
/// into AI system prompts so supplier-specific terminology resolves correctly.
/// </summary>
public class ProductSynonymService
{
    private readonly ILogger<ProductSynonymService> _log;
    private readonly string _jsonPath;

    private string? _cachedBlock;
    private DateTime _cachedAt = DateTime.MinValue;
    private static readonly TimeSpan CacheTtl = TimeSpan.FromMinutes(10);

    public ProductSynonymService(ILogger<ProductSynonymService> log, IWebHostEnvironment env)
    {
        _log      = log;
        _jsonPath = Path.Combine(env.ContentRootPath, "product-synonyms.json");
    }

    /// <summary>
    /// Returns the formatted synonym block for injection into a system prompt.
    /// Returns an empty string if the file is missing or unreadable.
    /// </summary>
    public string BuildPromptBlock()
    {
        var now = DateTime.UtcNow;
        if (_cachedBlock is not null && now - _cachedAt < CacheTtl)
            return _cachedBlock;

        _cachedBlock = LoadBlock();
        _cachedAt    = now;
        return _cachedBlock;
    }

    private string LoadBlock()
    {
        if (!File.Exists(_jsonPath))
        {
            _log.LogWarning("[Synonyms] product-synonyms.json not found at {Path}", _jsonPath);
            return string.Empty;
        }

        try
        {
            var json   = File.ReadAllText(_jsonPath);
            var groups = JsonSerializer.Deserialize<SynonymGroup[]>(json,
                new JsonSerializerOptions { PropertyNameCaseInsensitive = true });

            if (groups is null || groups.Length == 0)
                return string.Empty;

            var sb = new StringBuilder();
            sb.AppendLine();
            sb.AppendLine("── PRODUCT TERM SYNONYMS ────────────────────────────────────────────────");
            sb.AppendLine("When matching supplier descriptions to requested products, treat these as equivalent:");
            foreach (var g in groups)
            {
                if (g.Variants is null || g.Variants.Length == 0) continue;
                sb.Append("• \"").Append(g.Canonical).Append("\" = ")
                  .AppendLine(string.Join(", ", g.Variants));
            }
            sb.AppendLine("Always use the canonical name when matching, regardless of which variant the supplier used.");

            var block = sb.ToString();
            _log.LogDebug("[Synonyms] Loaded {Count} synonym groups from {Path}", groups.Length, _jsonPath);
            return block;
        }
        catch (Exception ex)
        {
            _log.LogError(ex, "[Synonyms] Failed to load product-synonyms.json");
            return string.Empty;
        }
    }

    private sealed class SynonymGroup
    {
        public string Canonical { get; set; } = string.Empty;
        public string? Category { get; set; }
        public string[]? Variants { get; set; }
    }
}
