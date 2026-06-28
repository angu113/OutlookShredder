using System.Security.Cryptography;
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
///   1. Containment — vendor tokens cover ≥ 55% of catalog non-dim tokens (≥ 2 overlap).
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
///   • "A-" used as a cross-section separator (AI extraction artifact for ×).
///
/// Config keys (optional — defaults shown):
///   ProductCatalog:SiteUrl   default: https://metalsupermarkets-my.sharepoint.com/personal/angus_mithrilmetals_com
///   ProductCatalog:ListName  default: Product Catalog
/// </summary>
public class ProductCatalogService : BackgroundService, ICacheStatusProvider
{
    private readonly IConfiguration _config;
    private readonly ILogger<ProductCatalogService> _log;
    private GraphServiceClient? _graph;

    private sealed record Entry(string Name, string? SearchKey, HashSet<string> Tokens,
        double? WeightPerFoot = null, string? WeightUnit = null,
        string? SpItemId = null,
        string? TkMetal = null, string? TkShape = null,
        string? TkAlloy = null, string? TkTemper = null, string[]? TkConditions = null,
        bool IsCustom = false,
        string? OriginalTerm = null,
        string? SupersededBy = null);
    private volatile IReadOnlyList<Entry> _cache = [];
    private string? _cachedSiteId;
    private string? _cachedListId;
    public string? LastError    { get; private set; }
    public string? LastDiag     { get; private set; }
    public DateTime? LastRefreshAt { get; private set; }

    // ICacheStatusProvider
    public string    Name          => "catalog";
    public string    DisplayName   => "Product Catalog";
    public int       SchemaVersion => 1;
    public int       ItemCount     => _cache.Count;
    public DateTime? CacheBuiltUtc => LastRefreshAt;
    public DateTime? LastDeltaUtc  => LastRefreshAt;
    public bool      IsLoading     => false;

    public async Task ForceRebuildAsync(CancellationToken ct = default) => await RefreshAsync();
    public async Task ForceDeltaAsync(CancellationToken ct = default)   => await RefreshAsync();

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

        var page = await graph.Sites[site.Id].Lists[list.Id].Items
            .GetAsync(r =>
            {
                r.QueryParameters.Expand = ["fields"];
                r.QueryParameters.Top    = 500;
            });

        var raw = new List<Microsoft.Graph.Models.ListItem>();
        while (page is not null)
        {
            raw.AddRange(page.Value ?? []);
            if (page.OdataNextLink is null) break;
            page = await graph.Sites[site.Id].Lists[list.Id].Items
                .WithUrl(page.OdataNextLink).GetAsync();
        }

        var firstFields = raw.Count > 0
            ? string.Join(", ", raw[0].Fields?.AdditionalData?.Keys ?? [])
            : "(no items)";

        var diag = $"site={site.Id} list={list.Id} rawItems={raw.Count} firstFields=[{firstFields}]";

        // Cache for PatchCatalogTokensAsync (called after tokenization)
        _cachedSiteId = site.Id;
        _cachedListId = list.Id;

        static string? Str(IDictionary<string, object> d, string key)
            => (d.TryGetValue(key, out var v) ? v : null)?.ToString();

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
                var rawWt = data.TryGetValue("WeightPerFoot", out var wv) ? wv : null;
                double? wt = rawWt switch
                {
                    double d  => d,
                    int    iv => (double)iv,
                    System.Text.Json.JsonElement je when je.ValueKind == System.Text.Json.JsonValueKind.Number => je.GetDouble(),
                    not null  => double.TryParse(rawWt.ToString(), out var dp) ? dp : null,
                    _         => null
                };
                var wtUnit     = Str(data, "WeightUnit");
                var tkConds    = Str(data, "TkConditions");
                var tkCondArr  = string.IsNullOrEmpty(tkConds) ? null
                    : tkConds.Split(';', StringSplitOptions.RemoveEmptyEntries);
                var isCustom     = data.TryGetValue("IsCustom", out var ic) && ic is bool icb && icb;
                var supersededBy = Str(data, "SupersededBy");
                // Skip superseded custom entries — they've been promoted to a real MSPC.
                if (!string.IsNullOrEmpty(supersededBy)) return null;
                return new Entry(name!, key, Tokenize(name!), wt, wtUnit, i.Id,
                    Str(data, "TkMetal"), Str(data, "TkShape"),
                    Str(data, "TkAlloy"), Str(data, "TkTemper"), tkCondArr,
                    IsCustom: isCustom,
                    OriginalTerm: Str(data, "OriginalTerm"),
                    SupersededBy: null);
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

    // ── Public lookups ────────────────────────────────────────────────────────

    public double? FindWeightPerFoot(string searchKey)
    {
        var c = _cache;
        var e = c.FirstOrDefault(e => e.SearchKey != null &&
            e.SearchKey.Equals(searchKey, StringComparison.OrdinalIgnoreCase));
        if (e is null) return null;
        // Only return weight for linear products (lb/ft); lb/sqft is a different unit
        if (e.WeightUnit is not null && !e.WeightUnit.Equals("lb/ft", StringComparison.OrdinalIgnoreCase))
            return null;
        return e.WeightPerFoot;
    }

    public string? FindProductName(string searchKey)
    {
        var c = _cache;
        return c.FirstOrDefault(e => e.SearchKey != null &&
            e.SearchKey.Equals(searchKey, StringComparison.OrdinalIgnoreCase))?.Name;
    }

    /// <summary>
    /// Returns catalog tokens keyed by SearchKey, built directly from the in-memory SP cache.
    /// Returns an empty dict if no entries have been tokenized yet (SP token columns not yet populated).
    /// </summary>
    public Dictionary<string, OutlookShredder.Proxy.Models.ProductTokens> GetTokensByKey()
    {
        return _cache
            .Where(e => e.SearchKey != null && e.TkMetal != null && e.TkShape != null)
            .GroupBy(e => e.SearchKey!, StringComparer.OrdinalIgnoreCase)
            .ToDictionary(
                g => g.Key,
                g => { var e = g.First(); return new OutlookShredder.Proxy.Models.ProductTokens
                {
                    Name         = e.Name,
                    SearchKey    = e.SearchKey,
                    TkMetal      = e.TkMetal,
                    TkShape      = e.TkShape,
                    TkAlloy      = e.TkAlloy,
                    TkTemper     = e.TkTemper,
                    TkConditions = e.TkConditions ?? [],
                }; },
                StringComparer.OrdinalIgnoreCase);
    }

    /// <summary>Returns the SP item ID for a catalog entry by exact product name (case-insensitive).</summary>
    public string? GetSpItemId(string productName)
        => _cache.FirstOrDefault(e =>
            e.Name.Equals(productName, StringComparison.OrdinalIgnoreCase))?.SpItemId;

    // ── Weight computation ────────────────────────────────────────────────────

    public record CatalogWeightItem(
        string SpId, string SearchKey, string ProductName,
        double? WeightValue, string? WeightUnit, string? Formula, string? Note,
        bool Updated);

    public record CatalogWeightReport(
        bool DryRun, int Total, int Computed, int AlreadySet, int Skipped,
        IReadOnlyList<CatalogWeightItem> Items);

    /// <summary>
    /// Iterates all Product Catalog SP items, computes theoretical weight via WeightCalculator,
    /// and PATCHes WeightPerFoot + WeightUnit back to SharePoint.
    /// dryRun=true reports without writing. overwrite=false skips items that already have a value.
    /// </summary>
    public async Task<CatalogWeightReport> ComputeWeightsAsync(
        bool dryRun = true, bool overwrite = false, CancellationToken ct = default)
    {
        var graph   = GetGraph();
        var siteUrl = _config["ProductCatalog:SiteUrl"]
            ?? "https://metalsupermarkets-my.sharepoint.com/personal/angus_mithrilmetals_com";
        var uri     = new Uri(siteUrl);
        var siteKey = $"{uri.Host}:{uri.AbsolutePath}";
        var site    = await graph.Sites[siteKey].GetAsync(cancellationToken: ct);
        if (site?.Id is null) throw new InvalidOperationException($"Site not found: {siteKey}");

        var listName = _config["ProductCatalog:ListName"] ?? "Product Catalog";
        var lists    = await graph.Sites[site.Id].Lists
            .GetAsync(r => r.QueryParameters.Filter = $"displayName eq '{listName}'",
                      cancellationToken: ct);
        var list = lists?.Value?.FirstOrDefault();
        if (list?.Id is null) throw new InvalidOperationException($"Catalog list '{listName}' not found");

        // Load all items with IDs, product names, and current weight
        var page = await graph.Sites[site.Id].Lists[list.Id].Items
            .GetAsync(r =>
            {
                r.QueryParameters.Expand = ["fields($select=id,Title,Product,Product_x0020_SearchKey,SearchKey,WeightPerFoot,WeightUnit)"];
                r.QueryParameters.Top    = 5000;
            }, cancellationToken: ct);

        var all = new List<(string SpId, string Name, string SearchKey, double? ExistingWeight)>();
        while (page is not null)
        {
            foreach (var item in page.Value ?? [])
            {
                if (item.Id is null) continue;
                var d = item.Fields?.AdditionalData;
                if (d is null) continue;
                var name = (d.TryGetValue("Product",                 out var p)  ? p  :
                            d.TryGetValue("Title",                   out var tt) ? tt : null)?.ToString();
                var key  = (d.TryGetValue("Product_x0020_SearchKey", out var sk) ? sk :
                            d.TryGetValue("SearchKey",               out var k2) ? k2 : null)?.ToString();
                double? existWt = null;
                if (d.TryGetValue("WeightPerFoot", out var wv) && wv is not null)
                    existWt = wv switch
                    {
                        double dv => dv,
                        int iv    => (double)iv,
                        System.Text.Json.JsonElement je when je.ValueKind == System.Text.Json.JsonValueKind.Number => je.GetDouble(),
                        _ => double.TryParse(wv.ToString(), out var dp) ? dp : null
                    };
                if (!string.IsNullOrWhiteSpace(name))
                    all.Add((item.Id, name!, key ?? "", existWt));
            }
            if (page.OdataNextLink is null) break;
            page = await graph.Sites[site.Id].Lists[list.Id].Items
                .WithUrl(page.OdataNextLink).GetAsync(cancellationToken: ct);
        }

        var items   = new System.Collections.Concurrent.ConcurrentBag<CatalogWeightItem>();
        int computed = 0, alreadySet = 0, skipped = 0;

        await System.Threading.Tasks.Parallel.ForEachAsync(all,
            new ParallelOptions { MaxDegreeOfParallelism = 4, CancellationToken = ct },
            async (row, token) =>
            {
                var wr = WeightCalculator.Calculate(row.Name);
                bool hasResult = wr.LbPerFoot.HasValue || wr.LbPerSqFt.HasValue;

                if (!hasResult)
                {
                    System.Threading.Interlocked.Increment(ref skipped);
                    items.Add(new CatalogWeightItem(row.SpId, row.SearchKey, row.Name,
                        null, null, null, wr.Note, false));
                    return;
                }

                if (!overwrite && row.ExistingWeight.HasValue)
                {
                    System.Threading.Interlocked.Increment(ref alreadySet);
                    items.Add(new CatalogWeightItem(row.SpId, row.SearchKey, row.Name,
                        row.ExistingWeight, null, wr.Formula, null, false));
                    return;
                }

                double weightVal  = wr.LbPerFoot ?? wr.LbPerSqFt!.Value;
                string weightUnit = wr.LbPerFoot.HasValue ? "lb/ft" : "lb/sqft";

                if (!dryRun)
                {
                    try
                    {
                        await graph.Sites[site.Id].Lists[list.Id].Items[row.SpId].Fields
                            .PatchAsync(new Microsoft.Graph.Models.FieldValueSet
                            {
                                AdditionalData = new Dictionary<string, object?>
                                {
                                    ["WeightPerFoot"] = weightVal,
                                    ["WeightUnit"]    = weightUnit
                                }
                            }, cancellationToken: token);
                    }
                    catch (Exception ex)
                    {
                        _log.LogWarning(ex, "[Catalog] Weight PATCH failed for item {Id}", row.SpId);
                        System.Threading.Interlocked.Increment(ref skipped);
                        items.Add(new CatalogWeightItem(row.SpId, row.SearchKey, row.Name,
                            weightVal, weightUnit, wr.Formula, $"PATCH failed: {ex.Message}", false));
                        return;
                    }
                }

                System.Threading.Interlocked.Increment(ref computed);
                items.Add(new CatalogWeightItem(row.SpId, row.SearchKey, row.Name,
                    weightVal, weightUnit, wr.Formula, wr.Note, !dryRun));
            });

        _log.LogInformation("[Catalog] ComputeWeights dryRun={Dry}: computed={C}, alreadySet={A}, skipped={S}, total={T}",
            dryRun, computed, alreadySet, skipped, all.Count);

        return new CatalogWeightReport(dryRun, all.Count, computed, alreadySet, skipped,
            items.OrderBy(i => i.SearchKey).ToList());
    }

    // ── SP column setup ───────────────────────────────────────────────────────

    /// <summary>Idempotently creates all optional computed columns on the Product Catalog list:
    /// WeightPerFoot, WeightUnit, and the five token fields (TkMetal/Shape/Alloy/Temper/Conditions).
    /// Safe to call repeatedly — skips columns that already exist.</summary>
    public async Task EnsureCatalogWeightColumnAsync()
    {
        var graph = GetGraph();
        var siteUrl = _config["ProductCatalog:SiteUrl"]
            ?? "https://metalsupermarkets-my.sharepoint.com/personal/angus_mithrilmetals_com";
        var uri     = new Uri(siteUrl);
        var siteKey = $"{uri.Host}:{uri.AbsolutePath}";
        var site    = await graph.Sites[siteKey].GetAsync();
        if (site?.Id is null) throw new InvalidOperationException($"Site not found: {siteKey}");

        var listName = _config["ProductCatalog:ListName"] ?? "Product Catalog";
        var lists    = await graph.Sites[site.Id].Lists
            .GetAsync(r => r.QueryParameters.Filter = $"displayName eq '{listName}'");
        var list = lists?.Value?.FirstOrDefault();
        if (list?.Id is null) throw new InvalidOperationException($"Catalog list '{listName}' not found");

        var cols = await graph.Sites[site.Id].Lists[list.Id].Columns.GetAsync();
        var exists = (cols?.Value ?? []).Any(c =>
            c.DisplayName?.Equals("WeightPerFoot", StringComparison.OrdinalIgnoreCase) == true);
        if (!exists)
        {
            await graph.Sites[site.Id].Lists[list.Id].Columns
                .PostAsync(new Microsoft.Graph.Models.ColumnDefinition
                {
                    DisplayName = "WeightPerFoot",
                    Name        = "WeightPerFoot",
                    Number      = new Microsoft.Graph.Models.NumberColumn()
                });
            _log.LogInformation("[Catalog] Created 'WeightPerFoot' column");
        }
        else
        {
            _log.LogInformation("[Catalog] 'WeightPerFoot' column already exists");
        }

        var hasUnit = (cols?.Value ?? []).Any(c =>
            c.DisplayName?.Equals("WeightUnit", StringComparison.OrdinalIgnoreCase) == true);
        if (!hasUnit)
        {
            await graph.Sites[site.Id].Lists[list.Id].Columns
                .PostAsync(new Microsoft.Graph.Models.ColumnDefinition
                {
                    DisplayName = "WeightUnit",
                    Name        = "WeightUnit",
                    Text        = new Microsoft.Graph.Models.TextColumn { MaxLength = 20 }
                });
            _log.LogInformation("[Catalog] Created 'WeightUnit' column");
        }

        // Token columns — populated by catalog tokenization; read back on each SP cache refresh
        // so LQ calculations survive reinstall without needing a file cache.
        var tokenCols = new[]
        {
            ("TkMetal",      20),
            ("TkShape",      40),
            ("TkAlloy",      40),
            ("TkTemper",     20),
            ("TkConditions", 200),   // semicolon-delimited list, e.g. "hot_rolled;sch10"
        };
        foreach (var (colName, maxLen) in tokenCols)
        {
            var exists2 = (cols?.Value ?? []).Any(c =>
                c.DisplayName?.Equals(colName, StringComparison.OrdinalIgnoreCase) == true);
            if (!exists2)
            {
                await graph.Sites[site.Id].Lists[list.Id].Columns
                    .PostAsync(new Microsoft.Graph.Models.ColumnDefinition
                    {
                        DisplayName = colName,
                        Name        = colName,
                        Text        = new Microsoft.Graph.Models.TextColumn { MaxLength = maxLen }
                    });
                _log.LogInformation("[Catalog] Created '{Col}' column", colName);
            }
        }

        // Custom-entry metadata columns (idempotent).
        var hasIsCustom = (cols?.Value ?? []).Any(c =>
            c.DisplayName?.Equals("IsCustom", StringComparison.OrdinalIgnoreCase) == true);
        if (!hasIsCustom)
        {
            await graph.Sites[site.Id].Lists[list.Id].Columns
                .PostAsync(new Microsoft.Graph.Models.ColumnDefinition
                {
                    DisplayName = "IsCustom",
                    Name        = "IsCustom",
                    Boolean     = new Microsoft.Graph.Models.BooleanColumn()
                });
            _log.LogInformation("[Catalog] Created 'IsCustom' column");
        }
        foreach (var metaCol in new[] { ("OriginalTerm", 255), ("Source", 20), ("SupersededBy", 100) })
        {
            var hasMeta = (cols?.Value ?? []).Any(c =>
                c.DisplayName?.Equals(metaCol.Item1, StringComparison.OrdinalIgnoreCase) == true);
            if (!hasMeta)
            {
                await graph.Sites[site.Id].Lists[list.Id].Columns
                    .PostAsync(new Microsoft.Graph.Models.ColumnDefinition
                    {
                        DisplayName = metaCol.Item1,
                        Name        = metaCol.Item1,
                        Text        = new Microsoft.Graph.Models.TextColumn { MaxLength = metaCol.Item2 }
                    });
                _log.LogInformation("[Catalog] Created '{Col}' column", metaCol.Item1);
            }
        }
        _customColumnsEnsured = true;
    }

    // ── Custom entry creation ─────────────────────────────────────────────────

    private volatile bool _customColumnsEnsured;
    private readonly SemaphoreSlim _ensureColsLock = new(1, 1);

    private async Task EnsureCustomColumnsOnceAsync()
    {
        if (_customColumnsEnsured) return;
        await _ensureColsLock.WaitAsync();
        try
        {
            if (_customColumnsEnsured) return;
            // Reuse EnsureCatalogWeightColumnAsync — it now covers custom columns too.
            await EnsureCatalogWeightColumnAsync();
        }
        finally { _ensureColsLock.Release(); }
    }

    /// <summary>
    /// Returns an existing custom catalog entry matching <paramref name="customId"/>, or creates a
    /// new one in SP and adds it to the local cache.  Thread-safe within a single proxy process;
    /// cross-proxy races produce at most one extra SP row (cleaned by /api/catalog/dedup).
    /// </summary>
    public async Task<(string CustomId, string Name)> GetOrCreateCustomEntryAsync(
        string customId, string term, OutlookShredder.Proxy.Models.ProductTokens tokens, string source)
    {
        // Fast path: already in local cache (covers same-session creates + hourly SP refresh).
        var existing = _cache.FirstOrDefault(e =>
            string.Equals(e.SearchKey, customId, StringComparison.OrdinalIgnoreCase));
        if (existing is not null)
            return (customId, existing.Name);

        await EnsureCustomColumnsOnceAsync();

        var graph    = GetGraph();
        var siteUrl  = _config["ProductCatalog:SiteUrl"]
                       ?? "https://metalsupermarkets-my.sharepoint.com/personal/angus_mithrilmetals_com";
        var uri      = new Uri(siteUrl);
        var siteKey  = $"{uri.Host}:{uri.AbsolutePath}";
        var site     = await graph.Sites[siteKey].GetAsync();
        if (site?.Id is null) throw new InvalidOperationException($"Site not found: {siteKey}");

        var listName = _config["ProductCatalog:ListName"] ?? "Product Catalog";
        var lists    = await graph.Sites[site.Id].Lists
            .GetAsync(r => r.QueryParameters.Filter = $"displayName eq '{listName}'");
        var list     = lists?.Value?.FirstOrDefault();
        if (list?.Id is null) throw new InvalidOperationException($"Catalog list '{listName}' not found");

        var condStr  = tokens.TkConditions.Length > 0 ? string.Join(";", tokens.TkConditions) : null;
        var data     = new Dictionary<string, object?>
        {
            ["Title"]                   = term,
            ["Product"]                 = term,
            ["Product_x0020_SearchKey"] = customId,
            ["IsCustom"]                = (object)true,
            ["OriginalTerm"]            = term,
            ["Source"]                  = source,
        };
        if (tokens.TkMetal  is not null) data["TkMetal"]      = tokens.TkMetal;
        if (tokens.TkShape  is not null) data["TkShape"]      = tokens.TkShape;
        if (tokens.TkAlloy  is not null) data["TkAlloy"]      = tokens.TkAlloy;
        if (tokens.TkTemper is not null) data["TkTemper"]     = tokens.TkTemper;
        if (condStr          is not null) data["TkConditions"] = condStr;

        Microsoft.Graph.Models.ListItem? created = null;
        try
        {
            created = await graph.Sites[site.Id].Lists[list.Id].Items
                .PostAsync(new Microsoft.Graph.Models.ListItem
                {
                    Fields = new Microsoft.Graph.Models.FieldValueSet { AdditionalData = data }
                });
            _log.LogInformation("[Catalog] Created custom entry {Id} for '{Term}' (source={Src})",
                customId, term, source);
        }
        catch (Exception ex)
        {
            _log.LogWarning(ex, "[Catalog] Failed to write custom entry {Id} to SP — using in-memory only",
                customId);
        }

        // Add to local cache immediately — subsequent calls in this process find it without SP.
        var condArr = condStr?.Split(';');
        var entry   = new Entry(term, customId, Tokenize(term), null, null, created?.Id,
            tokens.TkMetal, tokens.TkShape, tokens.TkAlloy, tokens.TkTemper, condArr,
            IsCustom: true);
        _cache = _cache.Append(entry).ToList().AsReadOnly();

        return (customId, term);
    }

    public async Task<int> DeleteBySearchKeysAsync(IReadOnlyList<string> searchKeys, CancellationToken ct = default)
    {
        var siteId = _cachedSiteId;
        var listId = _cachedListId;
        if (siteId is null || listId is null)
            throw new InvalidOperationException("Catalog site/list IDs not cached — call GET /api/catalog first");

        var graph = GetGraph();
        int deleted = 0;
        foreach (var key in searchKeys)
        {
            var match = _cache.Where(e =>
                string.Equals(e.SearchKey, key, StringComparison.OrdinalIgnoreCase) && e.SpItemId is not null).ToList();
            foreach (var entry in match)
            {
                try
                {
                    await graph.Sites[siteId].Lists[listId].Items[entry.SpItemId!].DeleteAsync(cancellationToken: ct);
                    _log.LogInformation("[Catalog] Deleted {Key} '{Name}' (SpItemId={Id})", key, entry.Name, entry.SpItemId);
                    deleted++;
                }
                catch (Exception ex)
                {
                    _log.LogWarning(ex, "[Catalog] Failed to delete {Key} (SpItemId={Id})", key, entry.SpItemId);
                }
            }
        }
        _cache = _cache.Where(e => !searchKeys.Contains(e.SearchKey, StringComparer.OrdinalIgnoreCase)).ToList().AsReadOnly();
        return deleted;
    }

    /// <summary>
    /// PATCHes TkMetal/TkShape/TkAlloy/TkTemper/TkConditions onto SP catalog items.
    /// Called by CatalogAnalysisService after each tokenization batch.
    /// Requires _cachedSiteId/_cachedListId populated by a prior FetchEntriesAsync call.
    /// </summary>
    public async Task PatchCatalogTokensAsync(
        IReadOnlyList<(string SpItemId, OutlookShredder.Proxy.Models.ProductTokens Tokens)> patches,
        CancellationToken ct = default)
    {
        if (patches.Count == 0) return;
        var siteId = _cachedSiteId;
        var listId = _cachedListId;
        if (siteId is null || listId is null)
        {
            _log.LogWarning("[Catalog] PatchCatalogTokensAsync: site/list IDs not cached — skipping SP write");
            return;
        }

        var graph = GetGraph();
        var tasks = patches.Select(async p =>
        {
            try
            {
                await graph.Sites[siteId].Lists[listId].Items[p.SpItemId].Fields
                    .PatchAsync(new Microsoft.Graph.Models.FieldValueSet
                    {
                        AdditionalData = new Dictionary<string, object?>
                        {
                            ["TkMetal"]      = p.Tokens.TkMetal,
                            ["TkShape"]      = p.Tokens.TkShape,
                            ["TkAlloy"]      = p.Tokens.TkAlloy,
                            ["TkTemper"]     = p.Tokens.TkTemper,
                            ["TkConditions"] = p.Tokens.TkConditions?.Length > 0
                                ? string.Join(";", p.Tokens.TkConditions) : null,
                        }
                    }, cancellationToken: ct);
            }
            catch (Exception ex)
            {
                _log.LogWarning(ex, "[Catalog] Token PATCH failed for SP item {Id}", p.SpItemId);
            }
        });

        await System.Threading.Tasks.Task.WhenAll(tasks);
        _log.LogInformation("[Catalog] Patched tokens on {Count} catalog SP items", patches.Count);
    }

    // ── Dedup ─────────────────────────────────────────────────────────────────

    /// <summary>
    /// Finds duplicate catalog rows (same non-null SearchKey) and deletes the extras,
    /// keeping the row with the lowest SP item ID (first written).
    /// When <paramref name="dryRun"/> is true, reports what would be deleted without touching SP.
    /// Returns (duplicateGroups, totalDeleted, report[]).
    /// </summary>
    public async Task<(int Groups, int Deleted, List<DedupCatalogGroup> Report)> DedupAsync(
        bool dryRun, CancellationToken ct)
    {
        var graph = GetGraph();

        var siteUrl  = _config["ProductCatalog:SiteUrl"]
                       ?? "https://metalsupermarkets-my.sharepoint.com/personal/angus_mithrilmetals_com";
        var uri      = new Uri(siteUrl);
        var siteKey  = $"{uri.Host}:{uri.AbsolutePath}";
        var site     = await graph.Sites[siteKey].GetAsync(cancellationToken: ct);
        if (site?.Id is null) throw new InvalidOperationException($"Site not found: '{siteKey}'");

        var listName = _config["ProductCatalog:ListName"] ?? "Product Catalog";
        var lists    = await graph.Sites[site.Id].Lists
            .GetAsync(r => r.QueryParameters.Filter = $"displayName eq '{listName}'",
                      cancellationToken: ct);
        var list = lists?.Value?.FirstOrDefault();
        if (list?.Id is null) throw new InvalidOperationException($"Catalog list '{listName}' not found");

        // Read all items with their SP item IDs and SearchKey/Product fields.
        var page = await graph.Sites[site.Id].Lists[list.Id].Items
            .GetAsync(r =>
            {
                r.QueryParameters.Expand = ["fields($select=id,Title,Product,Product_x0020_SearchKey,SearchKey)"];
                r.QueryParameters.Top    = 500;
            }, cancellationToken: ct);

        var all = new List<(string SpId, string? Name, string? SearchKey)>();
        while (page is not null)
        {
            foreach (var item in page.Value ?? [])
            {
                if (item.Id is null) continue;
                var d    = item.Fields?.AdditionalData;
                var name = (d?.TryGetValue("Product",                 out var p)  == true ? p  :
                            d?.TryGetValue("Title",                   out var tt) == true ? tt : null)
                           ?.ToString();
                var key  = (d?.TryGetValue("Product_x0020_SearchKey", out var sk) == true ? sk :
                            d?.TryGetValue("SearchKey",               out var k2) == true ? k2 : null)
                           ?.ToString();
                all.Add((item.Id, name, key));
            }
            if (page.OdataNextLink is null) break;
            page = await graph.Sites[site.Id].Lists[list.Id].Items
                .WithUrl(page.OdataNextLink).GetAsync(cancellationToken: ct);
        }

        // Group by non-null SearchKey; keep the row with the lowest numeric SP item ID.
        var groups = all
            .Where(x => !string.IsNullOrWhiteSpace(x.SearchKey))
            .GroupBy(x => x.SearchKey!, StringComparer.OrdinalIgnoreCase)
            .Where(g => g.Count() > 1)
            .ToList();

        var report  = new List<DedupCatalogGroup>();
        int deleted = 0;

        foreach (var g in groups)
        {
            var ordered = g
                .OrderBy(x => int.TryParse(x.SpId, out var n) ? n : int.MaxValue)
                .ToList();
            var keeper  = ordered[0];
            var extras  = ordered.Skip(1).ToList();

            report.Add(new DedupCatalogGroup(
                g.Key, keeper.SpId, keeper.Name,
                extras.Select(x => new DedupCatalogExtra(x.SpId, x.Name)).ToList()));

            if (!dryRun)
            {
                foreach (var extra in extras)
                {
                    await graph.Sites[site.Id].Lists[list.Id].Items[extra.SpId]
                        .DeleteAsync(cancellationToken: ct);
                    deleted++;
                    _log.LogInformation(
                        "[Catalog Dedup] Deleted SP item {Id} ('{Name}', MSPC={Key}) — duplicate of {KeepId}",
                        extra.SpId, extra.Name, g.Key, keeper.SpId);
                }
            }
        }

        return (groups.Count, deleted, report);
    }

    // ── Public resolver ───────────────────────────────────────────────────────

    /// <summary>
    /// Returns the best-matching catalog entry for <paramref name="rawName"/>, or
    /// <see langword="null"/> when no candidates pass the compatibility gates.
    /// Returns null immediately for service/processing items (powder coating, etc.).
    ///
    /// Confidence is 0–1:
    ///   ≥ 0.875 → SearchKey is set (high confidence, use MSPC)
    ///   0.50–0.874 → SearchKey is null (family match, dims uncertain — amber circle)
    ///   0–0.499 → SearchKey is null (weak match — red circle; only strategy 2 can reach this range)
    ///   null return → no match (caller stores confidence = 0)
    /// </summary>
    public (string Name, string? SearchKey, double Confidence)? ResolveProduct(string? rawName)
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

        const double HighConf = 0.875;

        var vendorTokens = Tokenize(rawName);
        var vendorNonDim = NonDimTokens(vendorTokens);

        // Strategy 1: containment with dim-overlap composite score.
        Entry? bestContained = null;
        double bestComposite = -1;
        double bestDimScore  = 0;
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
                bestDimScore  = dimScore;
            }
        }

        if (bestContained is not null)
        {
            // confidence = (1 + dimScore) / 2: 1.0 when all dims match, 0.5 when none match.
            var confidence = (1.0 + bestDimScore) / 2.0;
            var searchKey  = confidence >= HighConf ? bestContained.SearchKey : null;
            _log.LogDebug("[ProductCatalog] '{Raw}' → '{Catalog}' (containment conf={Conf:P0}, dim={Dim:P0})",
                rawName, bestContained.Name, confidence, bestDimScore);
            return (bestContained.Name, searchKey, confidence);
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
            var confidence = best.Jac * (1.0 + best.Dim) / 2.0;
            var searchKey  = confidence >= HighConf ? best.Entry.SearchKey : null;
            _log.LogDebug("[ProductCatalog] '{Raw}' → '{Catalog}' (jaccard conf={Conf:P0}, jac={Jac:F2}, dim={Dim:P0})",
                rawName, best.Entry.Name, confidence, best.Jac, best.Dim);
            return (best.Entry.Name, searchKey, confidence);
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

    /// <summary>Cached entries with name + searchKey, for API use.</summary>
    public IReadOnlyList<(string Name, string? SearchKey)> CachedEntries =>
        _cache.Select(e => (e.Name, e.SearchKey)).ToList().AsReadOnly();

    /// <summary>Top catalog candidates for a free-text request (reuses the same tokenizer + composite score as
    /// <see cref="ResolveProduct"/>). Used by the SMS inquiry clarifier (Phase 6) to give the AI the real
    /// product families to compare against. Returns name + MSPC + score, best first; empty when nothing scores.</summary>
    public IReadOnlyList<(string Name, string? SearchKey, double Score)> TopCandidates(string? rawName, int topN = 6)
    {
        if (string.IsNullOrWhiteSpace(rawName)) return [];
        var vendorTokens = Tokenize(rawName);
        var vendorNonDim = NonDimTokens(vendorTokens);
        if (vendorNonDim.Count == 0) return [];
        return _cache.Select(e =>
            {
                var catNonDim = NonDimTokens(e.Tokens);
                var overlap   = catNonDim.Count(t => vendorNonDim.Contains(t));
                var composite = overlap * (1.0 + DimOverlapFraction(e.Tokens, vendorTokens));
                return (e.Name, e.SearchKey, Score: composite);
            })
            .Where(x => x.Score > 0)
            .OrderByDescending(x => x.Score)
            .Take(topN)
            .Select(x => (x.Name, x.SearchKey, Math.Round(x.Score, 2)))
            .ToList();
    }

    /// <summary>
    /// Diagnostic: returns the tokenised non-dim tokens for a raw string, and the top-N
    /// containment/Jaccard scores against the catalog. Useful for debugging mismatches.
    /// </summary>
    public object Diagnose(string rawName, int topN = 10)
    {
        var vendorTokens = Tokenize(rawName);
        var vendorNonDim = NonDimTokens(vendorTokens);
        var vendorDim    = vendorTokens.Where(IsDimToken).ToList();

        var scores = _cache.Select(e =>
        {
            var catNonDim = NonDimTokens(e.Tokens);
            var overlap   = catNonDim.Count(t => vendorNonDim.Contains(t));
            var fraction  = catNonDim.Count > 0 ? (double)overlap / catNonDim.Count : 0;
            var dimScore  = DimOverlapFraction(e.Tokens, vendorTokens);
            var composite = overlap * (1.0 + dimScore);
            var jac       = Jaccard(catNonDim, vendorNonDim);
            return new
            {
                Name       = e.Name,
                SearchKey  = e.SearchKey,
                CatTokens  = catNonDim.OrderBy(t => t).ToList(),
                Overlap    = overlap,
                Fraction   = Math.Round(fraction, 3),
                DimScore   = Math.Round(dimScore, 3),
                Composite  = Math.Round(composite, 3),
                Jaccard    = Math.Round(jac, 3),
                PassesContainment = overlap >= 2 && fraction >= 0.55,
            };
        })
        .OrderByDescending(x => x.Composite).ThenByDescending(x => x.Jaccard)
        .Take(topN)
        .ToList();

        return new
        {
            Raw          = rawName,
            VendorNonDim = vendorNonDim.OrderBy(t => t).ToList(),
            VendorDim    = vendorDim.OrderBy(t => t).ToList(),
            TopMatches   = scores,
        };
    }

    /// <summary>
    /// Looks up a catalog entry by its SearchKey (MSPC).
    /// Used when the AI has already resolved the MSPC via RLI anchoring and we
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

    /// <summary>
    /// Checks whether an RLI product name is still consistent with the catalog entry
    /// for the given SearchKey. Used to detect cases where the user selected a catalog
    /// product but then edited the product name on the RFQ — in that case we should
    /// send the RLI item without an MSPC so the AI uses name-only matching.
    ///
    /// Returns the Jaccard similarity between the two tokenised names, the catalog entry
    /// name, and whether they are considered consistent (jaccard ≥ <paramref name="threshold"/>).
    /// Returns (false, 0, null) when the SearchKey is not in the catalog.
    /// </summary>
    public (bool Consistent, double Jaccard, string? CatalogName) CheckRliConsistency(
        string? searchKey, string? productName, double threshold = 0.25)
    {
        if (string.IsNullOrWhiteSpace(searchKey) || string.IsNullOrWhiteSpace(productName))
            return (true, 1.0, null);   // nothing to check — treat as consistent

        var entry = _cache.FirstOrDefault(e =>
            string.Equals(e.SearchKey, searchKey, StringComparison.OrdinalIgnoreCase));
        if (entry is null)
            return (false, 0, null);    // MSPC not in catalog — can't validate

        var catalogTokens = entry.Tokens;                   // already tokenised when cached
        var rliTokens     = Tokenize(productName);

        int overlap = catalogTokens.Intersect(rliTokens, StringComparer.OrdinalIgnoreCase).Count();
        int union   = catalogTokens.Union(rliTokens,     StringComparer.OrdinalIgnoreCase).Count();
        double jac  = union > 0 ? (double)overlap / union : 0;

        return (jac >= threshold, jac, entry.Name);
    }

    // ── Custom entry promotion ────────────────────────────────────────────────

    /// <summary>
    /// Returns all CUST_* catalog entries currently in the in-memory cache.
    /// Called after a CSV load to find entries that might be promoted to real MSPCs.
    /// </summary>
    public IReadOnlyList<CustomCatalogEntry> GetCustomCatalogEntries()
        => _cache
            .Where(e => e.IsCustom && e.SearchKey is not null && e.SpItemId is not null)
            .Select(e => new CustomCatalogEntry(
                SpItemId:     e.SpItemId!,
                SearchKey:    e.SearchKey!,
                OriginalTerm: e.OriginalTerm,
                ProductName:  e.Name,
                TkMetal:      e.TkMetal,
                TkShape:      e.TkShape,
                TkAlloy:      e.TkAlloy,
                TkTemper:     e.TkTemper,
                TkConditions: e.TkConditions))
            .ToList()
            .AsReadOnly();

    /// <summary>
    /// PATCHes <c>SupersededBy = <paramref name="mspc"/></c> on the catalog SP item so it is
    /// skipped by <see cref="FetchEntriesAsync"/> on the next refresh. Removes from the local
    /// in-memory cache immediately so subsequent lookups in this process skip it.
    /// Requires <c>_cachedSiteId</c>/<c>_cachedListId</c> — trigger a catalog refresh first.
    /// </summary>
    public async Task MarkCatalogEntrySupersededAsync(
        string spItemId, string mspc, CancellationToken ct = default)
    {
        var siteId = _cachedSiteId;
        var listId = _cachedListId;
        if (siteId is null || listId is null)
            throw new InvalidOperationException(
                "Catalog site/list IDs not cached — trigger a catalog refresh first");

        var graph = GetGraph();
        await graph.Sites[siteId].Lists[listId].Items[spItemId].Fields
            .PatchAsync(new Microsoft.Graph.Models.FieldValueSet
            {
                AdditionalData = new Dictionary<string, object?> { ["SupersededBy"] = mspc }
            }, cancellationToken: ct);

        // Remove from local cache immediately so it no longer participates in lookups.
        _cache = _cache.Where(e => e.SpItemId != spItemId).ToList().AsReadOnly();
        _log.LogInformation("[Catalog] Marked custom entry {SpId} superseded by {Mspc}", spItemId, mspc);
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
    // Also matches U+00C3 (Ã) + U+2014 (em dash) which AI extraction sometimes
    // writes instead of ASCII "A-" in beam designations like "W8[U+00C3][U+2014]15".
    private static readonly Regex _beamDesig =
        new(@"\b(MC|[WSC])\s*(\d+)\s*(?:[xX×]|A-|\u00C3\u2014)\s*(\d+(?:\.\d+)?)\b",
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

    // "A-" used as a cross-section dimension separator (AI extraction artifact for ×).
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

        // AR400 abrasion-resistant plate: normalise "AR 400", "AR400", "A.R. 400", "A.R.400" to "ar400".
        // "Abrasion Resistant" is a redundant English descriptor for AR400 — remove as noise.
        (new Regex(@"\ba\.?r\.?\s*400\b",       RegexOptions.IgnoreCase | RegexOptions.Compiled), "ar400"),
        (new Regex(@"\babrasion\s+resistant\b", RegexOptions.IgnoreCase | RegexOptions.Compiled), ""),

        // Rebar noise: "Grade 60" is the universal rebar grade and adds no matching value.
        (new Regex(@"\bgrade\s+60\b",     RegexOptions.IgnoreCase | RegexOptions.Compiled), ""),
    ];

    // Reuse the same dimension-normalisation approach as SharePointService.
    // Trailing length specification: " x 20'" or " x 144\"" at end of string.
    // Vendor strings append the cut length after the structural designation; it must be
    // stripped before dimSeparator runs so it isn't absorbed into the beam token.
    private static readonly Regex _trailingLengthSpec =
        new(@"\s+x\s+\d+(?:['""]|(?:\s*(?:ft|foot|feet|in|inch|inches)\b))\s*$",
            RegexOptions.Compiled | RegexOptions.IgnoreCase);

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
        // "Ã—" variant (U+00C3 + U+2014) used by AI extraction can be matched by
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

        // Convert "A-" cross-section separator (AI extraction artifact) to "x".
        // Only fires when preceded by a digit-form token (never matches "a-36" grade).
        s = _aDimSep.Replace(s, "$1 x ");

        // Strip trailing length specification before dimSeparator runs.
        // Vendor strings often end with " x 20'" or " x 144\"" for the cut length.
        // Without this, the dimSeparator absorbs the length into the beam designation
        // token (e.g. "mc6x16d3 x 20'" → "mc6x16d3x20d0") causing catalog mismatches.
        // Only strips when a foot/inch mark is present so bare cross-section dims are kept.
        s = _trailingLengthSpec.Replace(s, "");

        // Combine cross-section dimensions; bare integers normalised to "Nd0".
        s = _dimSeparator.Replace(s, m =>
            $"{NormDimPart(m.Groups[1].Value)}x{NormDimPart(m.Groups[2].Value)}");
        s = _dimSeparator.Replace(s, m =>
            $"{NormDimPart(m.Groups[1].Value)}x{NormDimPart(m.Groups[2].Value)}");

        // Normalise remaining 1-2 digit standalone integers to "Nd0" dim form so they
        // participate in dimension overlap. This handles nominal pipe/tube sizes (e.g.
        // "Pipe 1 SCH 40" → "1d0" to match catalog "1.000" → "1d0") and standalone
        // schedule numbers ("40" → "40d0"). Integers with 3+ digits are skipped to
        // preserve grade tokens (6061, 304, etc.) as non-dim tokens.
        s = Regex.Replace(s, @"\b(\d{1,2})\b", m => m.Groups[1].Value + "d0");

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

public record DedupCatalogExtra(string SpId, string? Name);
public record DedupCatalogGroup(string SearchKey, string KeeperSpId, string? KeeperName,
    List<DedupCatalogExtra> Extras);

/// <summary>A CUST_* catalog entry snapshot used for promotion candidate matching.</summary>
public record CustomCatalogEntry(
    string    SpItemId,
    string    SearchKey,
    string?   OriginalTerm,
    string?   ProductName,
    string?   TkMetal,
    string?   TkShape,
    string?   TkAlloy,
    string?   TkTemper,
    string[]? TkConditions);
