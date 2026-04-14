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
///   1. Containment — catalog non-dim tokens cover ≥ 55% of vendor tokens (≥ 2 overlap).
///      Composite score = overlap_count × (1 + dim_overlap_fraction) breaks ties.
///   2. Jaccard ≥ 0.30 on non-dimension tokens only, with all compatibility checks.
///
/// Compatibility gates (both strategies):
///   • Grade tokens (3+ pure-digit tokens, e.g. 304, 6061) must agree if both sides specify.
///   • Hot/Cold rolled are mutually exclusive (hotrolled vs coldrolled).
///   • Temper codes (H14, H22, T6, T6511, …) must agree if BOTH sides specify one;
///     if the catalog entry omits a temper, any vendor temper is accepted.
///   • [TRIAL] ASTM designation tokens (a500, a513, a36, a53, …) must agree if both
///     sides specify one; one-sided ASTM is allowed.
///
/// Pre-processing rules applied to both vendor and catalog strings:
///   • Parenthetical annotations stripped (removes H/W/FT/WT specs from beam catalog names).
///   • Mixed numbers normalised (1-1/2 → 1.500).
///   • Abbreviations expanded: HR/HRS → hotrolled, CR/CRS → coldrolled,
///     Wide Flange → wide beam, Carbon → steel, SS → stainless, Al/Alum → aluminum,
///     Cu → copper, Sch → schedule, Grade 60 → (removed), N GA/Gauge → Nga.
///   • Beam/channel designations normalised: "W8 X 24.0" and "W8x24" → "W8x24".
///   • HSS designations mapped to product types:
///       HSS NxMxt (N≠M) → rectangular tube, HSS NxNxt → square tube,
///       HSS N Sch t    → pipe (nominal).
///   • "A-" used as a cross-section separator (Claude extraction artifact for ×).
///
/// Config keys (optional — defaults shown):
///   ProductCatalog:SiteUrl   default: https://metalsupermarkets-my.sharepoint.com/personal/angus_mithrilmetals_com
///   ProductCatalog:ListName  default: Product Catalog
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
    /// <see langword="null"/> when no match is found.
    /// Returns null immediately for service/processing items (powder coating, etc.).
    /// </summary>
    public (string Name, string? SearchKey)? ResolveProduct(string? rawName)
    {
        if (string.IsNullOrWhiteSpace(rawName)) return null;

        // Service/processing items are never catalog products.
        if (_serviceKeywords.Any(kw => rawName.Contains(kw, StringComparison.OrdinalIgnoreCase)))
        {
            _log.LogDebug("[ProductCatalog] '{Raw}' → no match (service item)", rawName);
            return null;
        }

        var cache = _cache;   // atomic snapshot
        if (cache.Count == 0) return null;

        var vendorTokens = Tokenize(rawName);
        var vendorNonDim = NonDimTokens(vendorTokens);

        // Strategy 1: containment with dim-overlap composite score.
        Entry? bestContained = null;
        double bestComposite = -1;
        foreach (var entry in cache)
        {
            var catalogNonDim = NonDimTokens(entry.Tokens);
            if (catalogNonDim.Count < 2) continue;
            if (!GradeTokensCompatible(catalogNonDim, vendorNonDim)) continue;
            if (!HotColdCompatible(catalogNonDim, vendorNonDim)) continue;
            if (!TemperTokensCompatible(catalogNonDim, vendorNonDim)) continue;
            if (!AstmTokensCompatible(catalogNonDim, vendorNonDim)) continue;  // TRIAL
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

        // Strategy 2: Jaccard ≥ 0.30 on non-dimension tokens with all compatibility gates.
        var best = cache
            .Select(e =>
            {
                var catalogNonDim = NonDimTokens(e.Tokens);
                if (!GradeTokensCompatible(catalogNonDim, vendorNonDim)) return default;
                if (!HotColdCompatible(catalogNonDim, vendorNonDim)) return default;
                if (!TemperTokensCompatible(catalogNonDim, vendorNonDim)) return default;
                if (!AstmTokensCompatible(catalogNonDim, vendorNonDim)) return default;  // TRIAL
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

    /// <summary>
    /// Looks up a catalog entry by its SearchKey (MSPC).
    /// Used when Claude has already resolved the MSPC via RLI anchoring and we
    /// just need the canonical catalog name for CatalogProductName.
    /// Returns null when the key is not found in the current cache.
    /// </summary>
    public (string Name, string? SearchKey)? FindBySearchKey(string? searchKey)
    {
        if (string.IsNullOrWhiteSpace(searchKey)) return null;
        var entry = _cache.FirstOrDefault(e =>
            string.Equals(e.SearchKey, searchKey, StringComparison.OrdinalIgnoreCase));
        return entry is null ? null : (entry.Name, entry.SearchKey);
    }

    // ── Service item detection ────────────────────────────────────────────────

    // Products that are services/processing — never a catalog match.
    private static readonly string[] _serviceKeywords =
    [
        "powder coat", "powdercoat", "powder coating", "sandblast", "coating service",
    ];

    // ── Tokenisation ──────────────────────────────────────────────────────────

    // Strip parenthetical annotations from catalog names, e.g.:
    //   "HR Wide Beam W8 X 24.0 (H7.93 x W6.495 x FT 0.4 x WT 0.245)"
    //   → "HR Wide Beam W8 X 24.0"
    // Also removes "(Welded)", "(12 ft)", etc. from both vendor and catalog.
    private static readonly Regex _stripParens =
        new(@"\([^)]*\)", RegexOptions.Compiled);

    // Mixed number normalisation: "1-1/2" → 1.5, "3-3/4" → 3.75, etc.
    // Must run before _dimFraction so the fractional part is not processed in isolation.
    private static readonly Regex _mixedNumber =
        new(@"\b(\d+)-(\d+)/(\d+)\b", RegexOptions.Compiled);

    // Beam / channel designation normalisation.
    // Handles:  "W8 X 24.0", "W8x24", "W8 X 24", "W8A-20" (A- extraction artifact)
    //           "S3 X 5.7", "C3 X 4.1", "C3x4.1", "MC10 X 8.4"
    // Produces: "W8x24", "S3x5.7", "C3x4.1", "MC10x8.4"
    // The "x" in the result is intentional — IsDimToken now requires a leading digit,
    // so "w8x24" (starts with 'w') is classified as a non-dim token and participates
    // directly in the overlap/Jaccard calculation.
    // Also matches U+00C3 (Ã) + U+2014 (em dash) which Claude extraction sometimes
    // writes instead of ASCII "A-" in beam designations like "W8[U+00C3][U+2014]15".
    private static readonly Regex _beamDesig =
        new(@"\b(MC|[WSC])(\d+)\s*(?:[xX×]|A-|\u00C3\u2014)\s*(\d+(?:\.\d+)?)\b",
            RegexOptions.Compiled | RegexOptions.IgnoreCase);

    // HSS (Hollow Structural Section) designations used by suppliers.
    // Two dimensions + thickness → Rectangular or Square tube.
    // One dimension + schedule      → Pipe (nominal).
    private static readonly Regex _hssRect =
        new(@"\bHSS\s+(\d+(?:\.\d+)?)\s*[xX×]\s*(\d+(?:\.\d+)?)\s*[xX×]\s*(\d+(?:/\d+)?(?:\.\d+)?)\b",
            RegexOptions.Compiled | RegexOptions.IgnoreCase);
    private static readonly Regex _hssPipe =
        new(@"\bHSS\s+(\d+(?:\.\d+)?)\s+[Ss]ch(?:edule)?\s*(\d+)\b",
            RegexOptions.Compiled | RegexOptions.IgnoreCase);

    // "A-" used as a cross-section dimension separator (Claude extraction artifact for ×).
    // Applied AFTER decimal normalisation; requires a digit-token before "a-".
    // Example: "0d125" a- 0d75" a- 144" → "0d125 x 0d75 x 144"
    // Does NOT match "a-36" (ASTM grade) because there is no preceding digit token.
    // Negative lookbehind (?<![a-z0-9]) prevents matching digits that are part of a larger
    // alphanumeric token (e.g. "20" in "w6x20 a- 40'" must not trigger this rule).
    // (?<![a-z]) alone is insufficient: "0" in "w6x20" is preceded by "2" (a digit, not a letter).
    private static readonly Regex _aDimSep =
        new(@"(?<![a-z0-9])([0-9][0-9d]*)[""']?\s+a-\s+", RegexOptions.Compiled);

    // Abbreviation expansions applied in order after toLower.
    private static readonly (Regex Pattern, string Replacement)[] _abbrevExpansions =
    [
        // Aluminum alloy + temper fused-token split: "6063T52" -> "6063 T52", "6061T6511" -> "6061 T6511".
        // Must run first so the separated tokens trigger grade/temper compatibility gates.
        // Pattern: exactly 4 digits immediately followed by H or T + 1-6 digits (no space).
        (new Regex(@"\b(\d{4})([ht]\d{1,6})\b", RegexOptions.IgnoreCase | RegexOptions.Compiled), "$1 $2"),

        // Hot/cold rolled — unify all vendor and catalog variants to compound tokens.
        // "hot roll" (without "ed") is also used by some vendors and catalog entries.
        (new Regex(@"\bhot\s+roll(?:ed)?\b",  RegexOptions.IgnoreCase | RegexOptions.Compiled), "hotrolled"),
        (new Regex(@"\bcold\s+roll(?:ed)?\b", RegexOptions.IgnoreCase | RegexOptions.Compiled), "coldrolled"),
        (new Regex(@"\bh\.?r\.?s\.?\b",       RegexOptions.IgnoreCase | RegexOptions.Compiled), "hotrolled"),
        (new Regex(@"\bc\.?r\.?s\.?\b",       RegexOptions.IgnoreCase | RegexOptions.Compiled), "coldrolled"),
        // "H.R." abbreviation (two dotted letters) for hot rolled — must come AFTER h.r.s. rule.
        (new Regex(@"\bh\.r\.(?!\w)",         RegexOptions.IgnoreCase | RegexOptions.Compiled), "hotrolled "),
        (new Regex(@"\bc\.r\.(?!\w)",         RegexOptions.IgnoreCase | RegexOptions.Compiled), "coldrolled "),
        (new Regex(@"\bhr\b",                 RegexOptions.IgnoreCase | RegexOptions.Compiled), "hotrolled"),
        (new Regex(@"\bcr\b",                 RegexOptions.IgnoreCase | RegexOptions.Compiled), "coldrolled"),

        // Rectangular shorthand: "rect" and "rectangle" → "rectangular" (matches catalog names).
        (new Regex(@"\brect(?:angle)?\b",     RegexOptions.IgnoreCase | RegexOptions.Compiled), "rectangular"),

        // Material shorthands.
        (new Regex(@"\bss\b",             RegexOptions.IgnoreCase | RegexOptions.Compiled), "stainless"),
        (new Regex(@"\b(?:al|alum)\b",    RegexOptions.IgnoreCase | RegexOptions.Compiled), "aluminum"),
        (new Regex(@"\bcu\b",             RegexOptions.IgnoreCase | RegexOptions.Compiled), "copper"),

        // "Carbon" structural products map to "steel" (catalog uses "steel", not "carbon").
        (new Regex(@"\bcarbon\b",         RegexOptions.IgnoreCase | RegexOptions.Compiled), "steel"),

        // Structural section type synonyms.
        (new Regex(@"\bwide\s+flange\b",  RegexOptions.IgnoreCase | RegexOptions.Compiled), "wide beam"),

        // Schedule / pipe sizing.
        (new Regex(@"\bsch(?:edule)?\b",  RegexOptions.IgnoreCase | RegexOptions.Compiled), "schedule"),

        // Gauge normalisation: "12 gauge", "12 GA", "12GA" → "12ga".
        // Runs before decimal steps so e.g. "0.104" (the decimal thickness) is not affected.
        (new Regex(@"\b(\d{1,2})\s*(?:ga(?:uge)?)\b", RegexOptions.IgnoreCase | RegexOptions.Compiled), "$1ga"),

        // AR400 abrasion-resistant plate: normalise "AR 400" and "AR400" to "ar400".
        // "Abrasion Resistant" is a redundant English descriptor for AR400 — remove as noise.
        (new Regex(@"\bar\s*400\b",             RegexOptions.IgnoreCase | RegexOptions.Compiled), "ar400"),
        (new Regex(@"\babrasion\s+resistant\b", RegexOptions.IgnoreCase | RegexOptions.Compiled), ""),

        // Rebar noise: "Grade 60" is the universal rebar grade and adds no matching value.
        (new Regex(@"\bgrade\s+60\b",     RegexOptions.IgnoreCase | RegexOptions.Compiled), ""),
    ];

    // Reuse the same dimension-normalisation approach as SharePointService.
    private static readonly Regex _dimFraction  = new(@"(\d+)/(\d+)",                                      RegexOptions.Compiled);
    private static readonly Regex _dimDecimal   = new(@"(\d+)\.(\d+)",                                     RegexOptions.Compiled);
    private static readonly Regex _dimSeparator = new(@"(\d[a-z0-9]*)[""']?\s*[xX×]\s*[""']?(\d[a-z0-9]*)", RegexOptions.Compiled);
    private static readonly Regex _dimSplit     = new(@"[^a-z0-9]+",                                        RegexOptions.Compiled);
    private static readonly Regex _orLength     = new(@"\bor\s+\d+[a-z""']*\b",                             RegexOptions.IgnoreCase | RegexOptions.Compiled);

    private static HashSet<string> Tokenize(string s)
    {
        // Strip parenthetical annotations before any other processing.
        // Removes beam spec annotations like (H3 x W1.41 x FT 0.273 x WT 0.17),
        // processing notes like (Welded), length hints like (12 ft), etc.
        s = _stripParens.Replace(s, " ");

        // Normalise beam/channel designations BEFORE lowercasing so that the Unicode
        // "Ã—" variant (U+00C3 + U+2014) used by Claude extraction can be matched by
        // the regex directly. The replacement lowercases the prefix itself.
        // "W8 X 24.0" → "w8x24",  "W6A-20" → "w6x20",  "W8Ã—15" → "w8x15"
        s = _beamDesig.Replace(s, m =>
        {
            var prefix = m.Groups[1].Value.ToLowerInvariant();
            var height = m.Groups[2].Value;
            // Strip trailing ".0" cosmetic suffix on beam weights (e.g. "24.0" → "24").
            // Only strip when a decimal point is present — bare integers like "20" must
            // not be trimmed ("20".TrimEnd('0') → "2", which would break the token).
            var weight = m.Groups[3].Value;
            if (weight.Contains('.'))
                weight = weight.TrimEnd('0').TrimEnd('.');
            if (string.IsNullOrEmpty(weight)) weight = "0";
            return $"{prefix}{height}x{weight}";
        });

        s = s.ToLowerInvariant();

        // Normalise mixed numbers before fraction expansion.
        // "1-1/2" → 1 + 1/2 = 1.5 → "1.500"; "3-3/4" → "3.750"
        s = _mixedNumber.Replace(s, m =>
        {
            if (int.TryParse(m.Groups[1].Value, out var whole) &&
                int.TryParse(m.Groups[2].Value, out var num)   &&
                int.TryParse(m.Groups[3].Value, out var den)   && den != 0)
                return ((double)whole + (double)num / den).ToString("F3");
            return m.Value;
        });

        // Abbreviation expansions (order matters — hot/cold rolled before bare HR/CR).
        foreach (var (pattern, replacement) in _abbrevExpansions)
            s = pattern.Replace(s, replacement);

        // Map HSS supplier notation to the catalog product type.
        // HSS NxMxt where N≠M → "rectangular tube NxMxt"
        // HSS NxNxt          → "square tube NxNxt"
        // HSS N Sch t        → "pipe N schedule t"
        s = _hssRect.Replace(s, m =>
        {
            var d1 = m.Groups[1].Value;
            var d2 = m.Groups[2].Value;
            var t  = m.Groups[3].Value;
            var type = d1 == d2 ? "square tube" : "rectangular tube";
            return $"{type} {d1} x {d2} x {t}";
        });
        s = _hssPipe.Replace(s, m =>
            $"pipe {m.Groups[1].Value} schedule {m.Groups[2].Value}");

        // Convert fractions to 3-decimal equivalents (3/8 → 0.375).
        s = _dimFraction.Replace(s, m =>
        {
            if (int.TryParse(m.Groups[1].Value, out var num) &&
                int.TryParse(m.Groups[2].Value, out var den) && den != 0)
                return ((double)num / den).ToString("F3");
            return m.Value;
        });

        s = _orLength.Replace(s, "");
        s = Regex.Replace(s, @"\brandom\s+lengths?\b|\bmill\s+lengths?\b|\bfull\s+lengths?\b|\blengths?\b", "");

        // Decimal → internal "XdY" form; strip trailing zeros.
        s = _dimDecimal.Replace(s, "$1d$2");
        s = Regex.Replace(s, @"d(\d+)", m =>
        {
            var stripped = m.Groups[1].Value.TrimEnd('0');
            return "d" + (stripped.Length == 0 ? "0" : stripped);
        });

        // Convert "A-" cross-section separator (Claude extraction artifact) to "x".
        // Only fires when preceded by a digit-form token (never matches "a-36" grade).
        s = _aDimSep.Replace(s, "$1 x ");

        // Combine cross-section dimensions; bare integers normalised to "Nd0".
        s = _dimSeparator.Replace(s, m =>
            $"{NormDimPart(m.Groups[1].Value)}x{NormDimPart(m.Groups[2].Value)}");
        s = _dimSeparator.Replace(s, m =>
            $"{NormDimPart(m.Groups[1].Value)}x{NormDimPart(m.Groups[2].Value)}");

        s = Regex.Replace(s, @"[""']", "");
        return _dimSplit.Split(s)
            .Where(t => t.Length > 1 || (t.Length == 1 && char.IsDigit(t[0])))
            .ToHashSet();
    }

    private static string NormDimPart(string t) =>
        t.Length > 0 && t.All(char.IsDigit) ? t + "d0" : t;

    // ── Helpers ──────────────────────────────────────────────────────────────

    private static HashSet<string> DimComponents(HashSet<string> tokens) =>
        tokens.Where(IsDimToken)
              .SelectMany(t => t.Split('x'))
              .Where(c => c.Length > 0)
              .ToHashSet();

    private static double DimOverlapFraction(HashSet<string> catalogTokens, HashSet<string> vendorTokens)
    {
        var catDims = DimComponents(catalogTokens);
        if (catDims.Count == 0) return 0;
        var vendDims = DimComponents(vendorTokens);
        if (vendDims.Count == 0) return 0;
        return (double)catDims.Count(c => vendDims.Contains(c)) / catDims.Count;
    }

    /// <summary>
    /// A "dimension" token is one that starts with a digit and contains a dimension
    /// separator ('x' for cross-section, 'd' for decimal notation).
    /// Requiring a leading digit ensures beam/channel designations like "w8x24" or
    /// "c3x4d1" (which start with a letter) are classified as non-dim tokens and
    /// participate directly in the overlap/Jaccard calculation.
    /// </summary>
    private static bool IsDimToken(string t) =>
        t.Length > 0 && char.IsDigit(t[0]) && t.Any(c => c == 'x' || c == 'd');

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
    private static readonly HashSet<string> _commonLengths =
        Enumerable.Range(10, 35).Select(i => (i * 12).ToString()).ToHashSet();

    /// <summary>
    /// A "grade" token is a purely-numeric token with ≥ 3 digits that is NOT a common
    /// stock length (e.g. 304, 316, 1018, 6061 are grades; 120, 144, 240 are lengths).
    /// </summary>
    private static bool IsGradeToken(string t) =>
        t.All(char.IsDigit) && t.Length >= 3 && !_commonLengths.Contains(t);

    /// <summary>
    /// Returns false only when both token sets contain grade tokens that don't overlap.
    /// One-sided grade specification is always allowed.
    /// </summary>
    private static bool GradeTokensCompatible(HashSet<string> a, HashSet<string> b)
    {
        var gradeA = a.Where(IsGradeToken).ToHashSet();
        var gradeB = b.Where(IsGradeToken).ToHashSet();
        if (gradeA.Count > 0 && gradeB.Count > 0)
            return gradeA.Overlaps(gradeB);
        return true;
    }

    /// <summary>
    /// Hot Rolled and Cold Rolled are mutually exclusive.
    /// A vendor saying "hot rolled" will never match a catalog "cold rolled" entry.
    /// </summary>
    private static bool HotColdCompatible(HashSet<string> a, HashSet<string> b)
    {
        bool aHot  = a.Contains("hotrolled");
        bool aCold = a.Contains("coldrolled");
        bool bHot  = b.Contains("hotrolled");
        bool bCold = b.Contains("coldrolled");
        if (aHot  && bCold) return false;
        if (aCold && bHot)  return false;
        return true;
    }

    // Temper token pattern: aluminum H-temper (H14, H22, H32) and T-temper (T6, T6511).
    private static readonly Regex _temperPattern =
        new(@"^[ht]\d{1,4}$", RegexOptions.Compiled | RegexOptions.IgnoreCase);

    private static bool IsTemperToken(string t) => _temperPattern.IsMatch(t);

    /// <summary>
    /// Temper codes (H14, H22, T6, T6511, …) must agree when BOTH sides specify one.
    /// If the catalog entry omits a temper code, any vendor temper is accepted.
    /// This allows a generic catalog entry to match specific vendor temper callouts.
    /// </summary>
    private static bool TemperTokensCompatible(HashSet<string> catalogNonDim, HashSet<string> vendorNonDim)
    {
        var catTemper  = catalogNonDim.Where(IsTemperToken).ToHashSet();
        if (catTemper.Count == 0) return true;  // catalog doesn't specify — any vendor temper OK
        var vendTemper = vendorNonDim.Where(IsTemperToken).ToHashSet();
        if (vendTemper.Count == 0) return true;  // vendor doesn't specify — pass through
        return catTemper.Overlaps(vendTemper);
    }

    // ASTM designation token pattern (e.g. a500, a513, a36, a53b, a572).
    // [TRIAL] These are treated as one-sided tokens: only block if both sides specify
    // conflicting standards. A vendor omitting the ASTM designation is always allowed.
    private static readonly Regex _astmPattern =
        new(@"^a\d{2,3}[a-z]?$", RegexOptions.Compiled | RegexOptions.IgnoreCase);

    private static bool IsAstmToken(string t) => _astmPattern.IsMatch(t);

    /// <summary>
    /// [TRIAL] ASTM standard tokens (a500, a513, a36, a53, …) must agree only when
    /// BOTH the vendor and catalog descriptions specify a standard.
    /// One-sided ASTM callouts are always allowed (vendor may omit while catalog specifies).
    /// </summary>
    private static bool AstmTokensCompatible(HashSet<string> catalogNonDim, HashSet<string> vendorNonDim)
    {
        var catAstm  = catalogNonDim.Where(IsAstmToken).ToHashSet();
        var vendAstm = vendorNonDim.Where(IsAstmToken).ToHashSet();
        if (catAstm.Count == 0 || vendAstm.Count == 0) return true;  // one-sided — always OK
        return catAstm.Overlaps(vendAstm);
    }
}
