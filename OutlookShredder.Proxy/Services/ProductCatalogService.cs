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
    public string? LastError    { get; private set; }
    public string? LastDiag     { get; private set; }
    public DateTime? LastRefreshAt { get; private set; }

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

    public async Task RefreshAsync()
    {
        try
        {
            var (entries, diag) = await FetchEntriesAsync();
            _cache = entries.AsReadOnly();
            LastError     = null;
            LastDiag      = diag;
            LastRefreshAt = DateTime.UtcNow;
            _log.LogInformation("[ProductCatalog] Cache refreshed — {Count} product(s) loaded", entries.Count);
        }
        catch (Exception ex)
        {
            LastError = ex.Message;
            LastDiag  = ex.ToString();
            _log.LogWarning(ex, "[ProductCatalog] Cache refresh failed — stale cache will be used");
        }
    }

    // ── Graph fetch ───────────────────────────────────────────────────────────

    private async Task<(List<Entry> Entries, string Diag)> FetchEntriesAsync()
    {
        var graph = GetGraph();

        var siteUrl = _config["ProductCatalog:SiteUrl"]
            ?? "https://metalsupermarkets-my.sharepoint.com/personal/angus_mithrilmetals_com";
        var uri     = new Uri(siteUrl);
        var siteKey = $"{uri.Host}:{uri.AbsolutePath}";

        var site = await graph.Sites[siteKey].GetAsync();
        if (site?.Id is null)
            throw new InvalidOperationException($"Site not found: '{siteKey}'");

        var listName = _config["ProductCatalog:ListName"] ?? "Product Catalog";

        var lists = await graph.Sites[site.Id].Lists
            .GetAsync(r => r.QueryParameters.Filter = $"displayName eq '{listName}'");

        var list = lists?.Value?.FirstOrDefault();
        if (list?.Id is null)
        {
            var allLists = await graph.Sites[site.Id].Lists
                .GetAsync(r => r.QueryParameters.Select = ["displayName", "id"]);
            var names = string.Join(", ", allLists?.Value?.Select(l => l.DisplayName ?? "?") ?? []);
            throw new InvalidOperationException(
                $"List '{listName}' not found. Available lists: [{names}]");
        }

        var items = await graph.Sites[site.Id].Lists[list.Id].Items
            .GetAsync(r =>
            {
                r.QueryParameters.Expand = ["fields"];
                r.QueryParameters.Top    = 2000;
            });

        var raw = items?.Value ?? [];
        var firstFields = raw.Count > 0
            ? string.Join(", ", raw[0].Fields?.AdditionalData?.Keys ?? [])
            : "(no items)";

        var diag = $"site={site.Id} list={list.Id} rawItems={raw.Count} firstFields=[{firstFields}]";

        var entries = raw
            .Select(i =>
            {
                var data = i.Fields?.AdditionalData;
                if (data is null) return null;
                // Internal field names: "Product" (display: Product), "Product_x0020_SearchKey" (display: Product SearchKey).
                // Fallback to "Title" in case the list is restructured.
                var name = (data.TryGetValue("Product",                  out var p)  ? p  :
                            data.TryGetValue("Title",                    out var tt) ? tt : null)
                           ?.ToString();
                var key  = (data.TryGetValue("Product_x0020_SearchKey",  out var sk) ? sk :
                            data.TryGetValue("SearchKey",                out var k2) ? k2 : null)
                           ?.ToString();
                if (string.IsNullOrWhiteSpace(name)) return null;
                return new Entry(name!, key, Tokenize(name!));
            })
            .Where(e => e is not null)
            .Select(e => e!)
            .ToList();

        return (entries, diag);
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
    ///
    /// Matching is done on non-dimension tokens only. Catalog entries use decimal
    /// cross-section dimensions (0.375 X 5.000) while vendor descriptions use fractions
    /// (3/8 x 5 x 144), so dimension tokens will never match across sources. Only grade
    /// tokens (3+ digit pure numbers, e.g. 304, 6061) are used to block false matches.
    /// </summary>
    public (string Name, string? SearchKey)? ResolveProduct(string? rawName)
    {
        if (string.IsNullOrWhiteSpace(rawName)) return null;

        var cache = _cache;   // atomic snapshot
        if (cache.Count == 0) return null;

        var vendorTokens = Tokenize(rawName);
        var vendorNonDim = NonDimTokens(vendorTokens);

        // Strategy 1: containment with dim-overlap composite score.
        //
        // For each catalog entry:
        //   • fraction = (# catalog non-dim tokens found in vendor non-dim tokens)
        //                / (total catalog non-dim tokens)   — require ≥ 0.55 and ≥ 2 overlap tokens
        //   • dimScore = fraction of catalog's individual dim-components (e.g. "0d375", "5d0")
        //                that appear in vendor's dim-components (rewards correct size matches)
        //   • composite = overlap_count × (1 + dimScore)
        //
        // Sorting by composite rather than fraction alone means a catalog entry with 4/7 tokens
        // matching AND perfect dims beats one with 4/4 tokens but zero dim overlap.
        // This is critical when catalog entries include surface-finish words (Mill Finish, Ornamental)
        // that vendors omit, but share the same dimensions.
        Entry? bestContained = null;
        double bestComposite = -1;
        foreach (var entry in cache)
        {
            var catalogNonDim = NonDimTokens(entry.Tokens);
            if (catalogNonDim.Count < 2) continue;
            if (!GradeTokensCompatible(catalogNonDim, vendorNonDim)) continue;
            var overlap  = catalogNonDim.Count(t => vendorNonDim.Contains(t));
            if (overlap < 2) continue;
            var fraction = (double)overlap / catalogNonDim.Count;
            if (fraction < 0.55) continue;
            var dimScore  = DimOverlapFraction(entry.Tokens, vendorTokens);
            var composite = overlap * (1.0 + dimScore);
            if (composite > bestComposite)
            {
                bestContained = entry;
                bestComposite = composite;
            }
        }

        if (bestContained is not null)
        {
            _log.LogDebug("[ProductCatalog] '{Raw}' → '{Catalog}' (containment composite {Score:F2})",
                rawName, bestContained.Name, bestComposite);
            return (bestContained.Name, bestContained.SearchKey);
        }

        // Strategy 2: Jaccard ≥ 0.30 on non-dimension tokens with grade agreement.
        // Tiebreaker: dim-component overlap.
        var best = cache
            .Select(e =>
            {
                var catalogNonDim = NonDimTokens(e.Tokens);
                if (!GradeTokensCompatible(catalogNonDim, vendorNonDim)) return default;
                var jac = Jaccard(catalogNonDim, vendorNonDim);
                if (jac < 0.30) return default;
                var dim = DimOverlapFraction(e.Tokens, vendorTokens);
                return (Entry: e, Jac: jac, Dim: dim);
            })
            .Where(x => x.Entry is not null)
            .OrderByDescending(x => x.Jac)
            .ThenByDescending(x => x.Dim)
            .FirstOrDefault();

        if (best.Entry is not null)
        {
            _log.LogDebug("[ProductCatalog] '{Raw}' → '{Catalog}' (jaccard {Jac:F2}, dim {Dim:P0})",
                rawName, best.Entry.Name, best.Jac, best.Dim);
            return (best.Entry.Name, best.Entry.SearchKey);
        }

        // Log best candidate even on miss to aid diagnosis.
        var top = cache
            .Select(e => (Entry: e, Score: Jaccard(NonDimTokens(e.Tokens), vendorNonDim)))
            .OrderByDescending(x => x.Score)
            .FirstOrDefault();
        if (top.Entry is not null)
            _log.LogDebug("[ProductCatalog] No match for '{Raw}' — best '{Name}' jaccard {Score:F2}",
                rawName, top.Entry.Name, top.Score);

        return null;
    }

    /// <summary>Cached product names, for API use.</summary>
    public IReadOnlyList<string> CachedNames => _cache.Select(e => e.Name).ToList().AsReadOnly();

    // ── Tokenisation ──────────────────────────────────────────────────────────

    // Common abbreviation expansions so vendor shorthand matches catalog full names.
    private static readonly (Regex Pattern, string Replacement)[] _abbrevExpansions =
    [
        (new Regex(@"\bss\b",              RegexOptions.IgnoreCase | RegexOptions.Compiled), "stainless"),
        (new Regex(@"\bh\.?r\.?s\.?\b",   RegexOptions.IgnoreCase | RegexOptions.Compiled), "hotrolled"),
        (new Regex(@"\bc\.?r\.?s\.?\b",   RegexOptions.IgnoreCase | RegexOptions.Compiled), "coldrolled"),
        (new Regex(@"\b(?:al|alum)\b",    RegexOptions.IgnoreCase | RegexOptions.Compiled), "aluminum"),
        (new Regex(@"\bcu\b",             RegexOptions.IgnoreCase | RegexOptions.Compiled), "copper"),
        (new Regex(@"\bsch(?:edule)?\b",  RegexOptions.IgnoreCase | RegexOptions.Compiled), "schedule"),
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
        // Convert fractions to their 3-decimal-place equivalents so they can be
        // compared against catalog entries that use decimal notation (3/8 → 0.375).
        s = _dimFraction.Replace(s, m =>
        {
            if (int.TryParse(m.Groups[1].Value, out var num) &&
                int.TryParse(m.Groups[2].Value, out var den) && den != 0)
                return ((double)num / den).ToString("F3");
            return m.Value;
        });
        s = _orLength.Replace(s, "");
        s = Regex.Replace(s, @"\brandom\s+lengths?\b|\bmill\s+lengths?\b|\bfull\s+lengths?\b|\blengths?\b", "");
        // Decimal → internal "XdY" form; strip trailing zeros (5.000 → 5d0, 0.375 → 0d375).
        s = _dimDecimal.Replace(s, "$1d$2");
        s = Regex.Replace(s, @"d(\d+)", m => { var stripped = m.Groups[1].Value.TrimEnd('0'); return "d" + (stripped.Length == 0 ? "0" : stripped); });
        // Combine cross-section dimensions; bare integers are normalised to "Nd0" so
        // a vendor's "5" and a catalog's "5.000" produce the same token (5d0).
        s = _dimSeparator.Replace(s, m =>
            $"{NormDimPart(m.Groups[1].Value)}x{NormDimPart(m.Groups[2].Value)}");
        s = _dimSeparator.Replace(s, m =>
            $"{NormDimPart(m.Groups[1].Value)}x{NormDimPart(m.Groups[2].Value)}");
        s = Regex.Replace(s, @"[""']", "");
        return _dimSplit.Split(s)
            .Where(t => t.Length > 1 || (t.Length == 1 && char.IsDigit(t[0])))
            .ToHashSet();
    }

    /// <summary>
    /// Normalises a single dimension group captured by <see cref="_dimSeparator"/>:
    /// bare integers (no d/f/x suffix) get a "d0" suffix so "5" matches "5.000" → "5d0".
    /// Already-processed tokens like "0d375" or "1d0x2d0" are returned unchanged.
    /// </summary>
    private static string NormDimPart(string t) =>
        t.Length > 0 && t.All(char.IsDigit) ? t + "d0" : t;

    // ── Helpers ──────────────────────────────────────────────────────────────

    /// <summary>
    /// Splits combined dimension tokens ("0d375x5d0x144d0") into their individual
    /// components ("0d375", "5d0", "144d0") for proximity comparison.
    /// </summary>
    private static HashSet<string> DimComponents(HashSet<string> tokens) =>
        tokens.Where(IsDimToken)
              .SelectMany(t => t.Split('x'))
              .Where(c => c.Length > 0)
              .ToHashSet();

    /// <summary>
    /// Fraction of catalog dim-components found inside the vendor dim-components.
    /// Used as a tiebreaker so "0.375 X 5.000" beats "0.125 X 0.500" when the
    /// vendor says "3/8 x 5".  Returns 0 when either side has no dim tokens.
    /// </summary>
    private static double DimOverlapFraction(HashSet<string> catalogTokens, HashSet<string> vendorTokens)
    {
        var catDims = DimComponents(catalogTokens);
        if (catDims.Count == 0) return 0;
        var vendDims = DimComponents(vendorTokens);
        if (vendDims.Count == 0) return 0;
        return (double)catDims.Count(c => vendDims.Contains(c)) / catDims.Count;
    }

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

    // Common metal stock lengths in inches (multiples of 12, 10 ft – 44 ft).
    // These look like grade numbers (3+ digits) but are measurements, not material grades.
    private static readonly HashSet<string> _commonLengths =
        Enumerable.Range(10, 35).Select(i => (i * 12).ToString()).ToHashSet();

    /// <summary>
    /// A "grade" token is a purely-numeric token with ≥ 3 digits that is NOT a common stock
    /// length (e.g. 304, 316, 1018, 6061 are grades; 120, 144, 240, 288 are lengths and excluded).
    /// </summary>
    private static bool IsGradeToken(string t) =>
        t.All(char.IsDigit) && t.Length >= 3 && !_commonLengths.Contains(t);

    /// <summary>
    /// Returns false only when both token sets contain grade tokens that don't overlap
    /// (e.g. 304 vs 316). If only one side specifies a grade, the match is allowed.
    /// Dimension tokens are intentionally not checked here — the catalog uses decimal
    /// dimensions and vendor descriptions use fractions, so they never compare equal.
    /// </summary>
    private static bool GradeTokensCompatible(HashSet<string> a, HashSet<string> b)
    {
        var gradeA = a.Where(IsGradeToken).ToHashSet();
        var gradeB = b.Where(IsGradeToken).ToHashSet();
        if (gradeA.Count > 0 && gradeB.Count > 0)
            return gradeA.Overlaps(gradeB);
        return true;
    }
}
