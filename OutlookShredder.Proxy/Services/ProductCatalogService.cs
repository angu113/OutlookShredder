using System.Text.RegularExpressions;
using Azure.Identity;
using Microsoft.Graph;

namespace OutlookShredder.Proxy.Services;

/// <summary>
/// Singleton background service that fetches the canonical product list from the
/// "Product Catalog" SharePoint list and resolves incoming (vendor-extracted) product
/// names to their canonical catalog name and search key via fuzzy matching.
///
/// Matching strategy (in priority order):
///   1. Containment — all catalog tokens are present in the vendor description.
///      Requires ≥ 2 catalog tokens and numeric/grade tokens must agree.
///   2. Jaccard ≥ 0.25 on non-dimension tokens only, with grade-token agreement.
///
/// Config keys (optional — defaults shown):
///   ProductCatalog:SiteUrl   default: https://metalsupermarkets-my.sharepoint.com/personal/angus_mithrilmetals_com
///   ProductCatalog:ListName  default: Product Catalog
///
/// Uses the same app-only credentials as SharePointService.
/// </summary>
public class ProductCatalogService : BackgroundService
{
    private readonly IConfiguration _config;
    private readonly ILogger<ProductCatalogService> _log;
    private GraphServiceClient? _graph;

    private sealed record Entry(string Name, string? SearchKey, HashSet<string> Tokens);
    private volatile IReadOnlyList<Entry> _cache = [];

    public ProductCatalogService(IConfiguration config, ILogger<ProductCatalogService> log)
    {
        _config = config;
        _log    = log;
    }

    // ── Background loop ───────────────────────────────────────────────────────

    protected override async Task ExecuteAsync(CancellationToken stoppingToken)
    {
        while (!stoppingToken.IsCancellationRequested)
        {
            await RefreshAsync();
            try { await Task.Delay(TimeSpan.FromHours(1), stoppingToken); }
            catch (OperationCanceledException) { break; }
        }
    }

    private async Task RefreshAsync()
    {
        try
        {
            var entries = await FetchEntriesAsync();
            _cache = entries.AsReadOnly();
            _log.LogInformation("[ProductCatalog] Cache refreshed — {Count} product(s) loaded", entries.Count);
        }
        catch (Exception ex)
        {
            _log.LogWarning(ex, "[ProductCatalog] Cache refresh failed — stale cache will be used");
        }
    }

    // ── Graph fetch ───────────────────────────────────────────────────────────

    private async Task<List<Entry>> FetchEntriesAsync()
    {
        var graph = GetGraph();

        var siteUrl = _config["ProductCatalog:SiteUrl"]
            ?? "https://metalsupermarkets-my.sharepoint.com/personal/angus_mithrilmetals_com";
        var uri     = new Uri(siteUrl);
        var siteKey = $"{uri.Host}:{uri.AbsolutePath}";

        var site = await graph.Sites[siteKey].GetAsync();
        if (site?.Id is null)
        {
            _log.LogWarning("[ProductCatalog] Could not resolve site '{Key}'", siteKey);
            return [];
        }

        var listName = _config["ProductCatalog:ListName"] ?? "Product Catalog";
        var lists = await graph.Sites[site.Id].Lists
            .GetAsync(r => r.QueryParameters.Filter = $"displayName eq '{listName}'");

        var list = lists?.Value?.FirstOrDefault();
        if (list?.Id is null)
        {
            _log.LogWarning("[ProductCatalog] List '{Name}' not found in site '{Key}'", listName, siteKey);
            return [];
        }

        var items = await graph.Sites[site.Id].Lists[list.Id].Items
            .GetAsync(r =>
            {
                r.QueryParameters.Expand = ["fields($select=Title,SearchKey)"];
                r.QueryParameters.Top    = 2000;
            });

        return items?.Value?
            .Select(i =>
            {
                var data = i.Fields?.AdditionalData;
                if (data is null) return null;
                var name = data.TryGetValue("Title",     out var t) ? t?.ToString() : null;
                var key  = data.TryGetValue("SearchKey", out var k) ? k?.ToString() : null;
                if (string.IsNullOrWhiteSpace(name)) return null;
                return new Entry(name!, key, Tokenize(name!));
            })
            .Where(e => e is not null)
            .Select(e => e!)
            .ToList()
            ?? [];
    }

    private GraphServiceClient GetGraph()
    {
        if (_graph is not null) return _graph;
        var tenantId     = _config["SharePoint:TenantId"]     ?? throw new InvalidOperationException("SharePoint:TenantId not set");
        var clientId     = _config["SharePoint:ClientId"]     ?? throw new InvalidOperationException("SharePoint:ClientId not set");
        var clientSecret = _config["SharePoint:ClientSecret"] ?? throw new InvalidOperationException("SharePoint:ClientSecret not set");
        var credential   = new ClientSecretCredential(tenantId, clientId, clientSecret);
        _graph = new GraphServiceClient(credential, ["https://graph.microsoft.com/.default"]);
        return _graph;
    }

    // ── Public resolver ───────────────────────────────────────────────────────

    /// <summary>
    /// Returns the best-matching catalog entry for <paramref name="rawName"/>, or
    /// <see langword="null"/> when the cache is empty or no match is found.
    /// </summary>
    public (string Name, string? SearchKey)? ResolveProduct(string? rawName)
    {
        if (string.IsNullOrWhiteSpace(rawName)) return null;

        var cache = _cache;   // atomic snapshot
        if (cache.Count == 0) return null;

        var vendorTokens = Tokenize(rawName);

        // Strategy 1: all catalog tokens are contained in the vendor description.
        // Requires ≥ 2 catalog tokens so single-word entries don't match too broadly.
        // Grade/numeric tokens must agree (no matching 304 catalog to a 316 vendor name).
        Entry? bestContained = null;
        int    bestCount     = 0;
        foreach (var entry in cache)
        {
            if (entry.Tokens.Count < 2) continue;
            if (!entry.Tokens.IsSubsetOf(vendorTokens)) continue;
            if (!NumericTokensCompatible(entry.Tokens, vendorTokens)) continue;
            if (entry.Tokens.Count > bestCount)
            {
                bestContained = entry;
                bestCount     = entry.Tokens.Count;
            }
        }

        if (bestContained is not null)
        {
            _log.LogDebug("[ProductCatalog] '{Raw}' → '{Catalog}' (containment, {N} tokens)",
                rawName, bestContained.Name, bestCount);
            return (bestContained.Name, bestContained.SearchKey);
        }

        // Strategy 2: Jaccard ≥ 0.25 on non-dimension tokens, with grade agreement.
        var vendorNonDim = NonDimTokens(vendorTokens);

        var best = cache
            .Where(e =>
            {
                if (!NumericTokensCompatible(e.Tokens, vendorTokens)) return false;
                var catalogNonDim = NonDimTokens(e.Tokens);
                return Jaccard(catalogNonDim, vendorNonDim) >= 0.25;
            })
            .Select(e => (Entry: e, Score: Jaccard(NonDimTokens(e.Tokens), vendorNonDim)))
            .OrderByDescending(x => x.Score)
            .FirstOrDefault();

        if (best.Entry is not null)
        {
            _log.LogDebug("[ProductCatalog] '{Raw}' → '{Catalog}' (jaccard {Score:F2})",
                rawName, best.Entry.Name, best.Score);
            return (best.Entry.Name, best.Entry.SearchKey);
        }

        return null;
    }

    /// <summary>Cached product names, for API use.</summary>
    public IReadOnlyList<string> CachedNames => _cache.Select(e => e.Name).ToList().AsReadOnly();

    // ── Tokenisation ──────────────────────────────────────────────────────────

    // Common abbreviation expansions so vendor shorthand matches catalog full names.
    private static readonly (Regex Pattern, string Replacement)[] _abbrevExpansions =
    [
        (new Regex(@"\bss\b",   RegexOptions.IgnoreCase | RegexOptions.Compiled), "stainless"),
        (new Regex(@"\bh\.?r\.?s\.?\b", RegexOptions.IgnoreCase | RegexOptions.Compiled), "hotrolled"),
        (new Regex(@"\bc\.?r\.?s\.?\b", RegexOptions.IgnoreCase | RegexOptions.Compiled), "coldrolled"),
        (new Regex(@"\b(?:al|alum)\b",  RegexOptions.IgnoreCase | RegexOptions.Compiled), "aluminum"),
        (new Regex(@"\bcu\b",           RegexOptions.IgnoreCase | RegexOptions.Compiled), "copper"),
    ];

    // Reuse the same dimension-normalisation approach as SharePointService.
    private static readonly Regex _dimFraction  = new(@"(\d+)/(\d+)",                                      RegexOptions.Compiled);
    private static readonly Regex _dimDecimal   = new(@"(\d+)\.(\d+)",                                     RegexOptions.Compiled);
    private static readonly Regex _dimSeparator = new(@"(\d[a-z0-9]*)[""']?\s*[xX×]\s*[""']?(\d[a-z0-9]*)", RegexOptions.Compiled);
    private static readonly Regex _dimSplit     = new(@"[^a-z0-9]+",                                        RegexOptions.Compiled);
    private static readonly Regex _orLength     = new(@"\bor\s+\d+[a-z""']*\b",                             RegexOptions.IgnoreCase | RegexOptions.Compiled);

    private static HashSet<string> Tokenize(string s)
    {
        s = s.ToLowerInvariant();
        // Expand abbreviations before anything else.
        foreach (var (pattern, replacement) in _abbrevExpansions)
            s = pattern.Replace(s, replacement);
        // Dimension normalisation (mirrors SharePointService.PreprocessProduct).
        s = _orLength.Replace(s, "");
        s = Regex.Replace(s, @"\brandom\s+lengths?\b|\bmill\s+lengths?\b|\bfull\s+lengths?\b|\blengths?\b", "");
        s = _dimFraction.Replace(s, "$1f$2");
        s = _dimDecimal.Replace(s, "$1d$2");
        s = Regex.Replace(s, @"d(\d+)", m => { var stripped = m.Groups[1].Value.TrimEnd('0'); return "d" + (stripped.Length == 0 ? "0" : stripped); });
        s = _dimSeparator.Replace(s, "$1x$2");
        s = _dimSeparator.Replace(s, "$1x$2");
        s = Regex.Replace(s, @"[""']", "");
        return _dimSplit.Split(s)
            .Where(t => t.Length > 1 || (t.Length == 1 && char.IsDigit(t[0])))
            .ToHashSet();
    }

    // ── Helpers ───────────────────────────────────────────────────────────────

    // A "dimension" token contains a digit plus one of the compound separators (x, f, d).
    private static bool IsDimToken(string t) =>
        t.Any(char.IsDigit) && t.Any(c => c == 'x' || c == 'f' || c == 'd');

    private static HashSet<string> NonDimTokens(HashSet<string> tokens) =>
        tokens.Where(t => !IsDimToken(t)).ToHashSet();

    private static double Jaccard(HashSet<string> a, HashSet<string> b)
    {
        if (a.Count == 0 && b.Count == 0) return 1.0;
        var intersection = a.Count(t => b.Contains(t));
        var union        = a.Count + b.Count - intersection;
        return union == 0 ? 0 : (double)intersection / union;
    }

    /// <summary>
    /// Returns true when the numeric tokens in both sets are compatible:
    /// grade-only tokens must be subsets of each other; dimension tokens must agree when both sides have them.
    /// </summary>
    private static bool NumericTokensCompatible(HashSet<string> a, HashSet<string> b)
    {
        var numA = a.Where(t => t.Any(char.IsDigit)).ToHashSet();
        var numB = b.Where(t => t.Any(char.IsDigit)).ToHashSet();
        var dimA = numA.Where(IsDimToken).ToHashSet();
        var dimB = numB.Where(IsDimToken).ToHashSet();

        if (dimA.Count > 0 && dimB.Count > 0)
        {
            if (!dimA.SetEquals(dimB)) return false;
            var gradeA = numA.Where(t => !IsDimToken(t)).ToHashSet();
            var gradeB = numB.Where(t => !IsDimToken(t)).ToHashSet();
            return gradeA.IsSubsetOf(gradeB) || gradeB.IsSubsetOf(gradeA);
        }

        if (dimA.Count > 0) return false;   // A has dims, B doesn't — block

        var gA = numA.Where(t => !IsDimToken(t)).ToHashSet();
        var gB = numB.Where(t => !IsDimToken(t)).ToHashSet();
        return gA.IsSubsetOf(gB) || gB.IsSubsetOf(gA);
    }
}
