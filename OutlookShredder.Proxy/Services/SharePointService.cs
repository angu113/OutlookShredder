using System.IO.Compression;
using System.Text.Json;
using System.Text.RegularExpressions;
using Azure.Identity;
using Microsoft.Graph;
using Microsoft.Graph.Models;
using Microsoft.Kiota.Abstractions.Serialization;
using OutlookShredder.Proxy.Models;

namespace OutlookShredder.Proxy.Services;

/// <summary>
/// Manages all SharePoint list operations via Microsoft Graph (app-only auth).
///
/// Lists:
///   SupplierResponses   -- one row per supplier email; holds email metadata + body + attachment
///   SupplierLineItems   -- one row per extracted product line; child of SupplierResponses
///   RFQ References      -- source RFQs written by ShredderXL; holds Notes field
///
/// Azure AD app requires:  Sites.FullControl.All  (Application, admin consented)
/// </summary>
public class SharePointService
{
    private readonly IConfiguration          _config;
    private readonly ILogger<SharePointService> _log;
    private readonly SupplierCacheService  _suppliers;
    private readonly ProductCatalogService _catalog;

    private GraphServiceClient?     _graph;
    private ClientSecretCredential? _spCredential;
    private string? _siteId;

    // Cached list IDs (lazy-resolved)
    private string? _srListId;            // SupplierResponses
    private string? _sliListId;           // SupplierLineItems
    private string? _rfqRefListId;        // RFQ References
    private string? _listId;              // RFQ Line Items (legacy  -- kept for EnsureColumnsAsync)
    private string? _qcSiteId;            // QC SP site
    private string? _qcListId;            // QC list
    private string? _shredderConfigListId; // ShredderConfig
    private string? _poListId;            // PurchaseOrders

    // SR row cache  --  SupplierResponses rows change only on email writes.
    // Caching for 60 s eliminates repeated full-table fetches during paginated loads
    // and concurrent startup requests. Write paths call InvalidateSrCache().
    private List<Microsoft.Graph.Models.ListItem>? _srRowCache;
    private DateTime                               _srRowCacheExpiry = DateTime.MinValue;
    private readonly SemaphoreSlim                 _srRowCacheLock   = new(1, 1);

    private static readonly string[] _regretPhrases =
        ["regret", "no stock", "unable to supply", "cannot supply", "not available", "out of stock",
         "do not offer", "no quote", "not able to offer", "unable to offer"];

    // ??"?????"??? Shared SP read helpers ??"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"???

    /// <summary>
    /// Coerces a SharePoint AdditionalData value to string.
    /// Graph SDK deserialises field values as JsonElement; this handles that plus plain strings.
    /// </summary>
    private static string? Str(object? v) => v switch
    {
        string s                                                       => s,
        JsonElement je when je.ValueKind == JsonValueKind.String       => je.GetString(),
        JsonElement je                                                 => je.ToString(),
        null                                                           => null,
        _                                                              => v.ToString()
    };

    private static string? GetStr(IDictionary<string, object?> d, string key) =>
        d.TryGetValue(key, out var v) ? Str(v) : null;

    private static string? GetStrRaw(IDictionary<string, object> d, string key) =>
        d.TryGetValue(key, out var v) ? Str(v) : null;

    /// <summary>
    /// Reads the RFQ ID from a field dictionary, handling both the logical name
    /// "RFQ_ID" and SharePoint's URL-encoded internal name "RFQ_x005F_ID".
    /// </summary>
    private static string? RfqId(IDictionary<string, object?> d) =>
        GetStr(d, "RFQ_ID") ?? GetStr(d, "RFQ_x005F_ID");

    private static string? RfqIdRaw(IDictionary<string, object> d) =>
        GetStrRaw(d, "RFQ_ID") ?? GetStrRaw(d, "RFQ_x005F_ID");

    /// <summary>Fields lifted from SupplierResponses into each SupplierLineItems read row.</summary>
    private static readonly string[] ParentFields =
    [
        "EmailFrom", "ContactEmail", "ReceivedAt", "ProcessedAt", "ProcessingSource",
        "SourceFile", "DateOfQuote", "EstimatedDeliveryDate",
        "QuoteReference", "FreightTerms", "EmailBody", "EmailSubject"
    ];

    private static bool IsAppField(string key) =>
        !key.StartsWith('@') &&
        !key.StartsWith('_') &&
        key is not ("LinkTitle" or "LinkTitleNoMenu" or "ContentType"
                 or "Edit" or "Attachments" or "ItemChildCount" or "FolderChildCount"
                 or "AuthorLookupId" or "EditorLookupId"
                 or "AppAuthorLookupId" or "AppEditorLookupId");

    public SharePointService(IConfiguration config, ILogger<SharePointService> log,
        SupplierCacheService suppliers, ProductCatalogService catalog)
    {
        _config    = config;
        _log       = log;
        _suppliers = suppliers;
        _catalog   = catalog;
    }

    // ??"?????"??? Graph client (lazy init) ??"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"???
    private GraphServiceClient GetGraph()
    {
        if (_graph is not null) return _graph;

        var tenantId     = _config["SharePoint:TenantId"]     ?? throw new InvalidOperationException("SharePoint:TenantId not set");
        var clientId     = _config["SharePoint:ClientId"]     ?? throw new InvalidOperationException("SharePoint:ClientId not set");
        var clientSecret = _config["SharePoint:ClientSecret"] ?? throw new InvalidOperationException("SharePoint:ClientSecret not set");

        var credential = new ClientSecretCredential(tenantId, clientId, clientSecret);
        _graph = new GraphServiceClient(credential, ["https://graph.microsoft.com/.default"]);
        return _graph;
    }

    // ??"?????"??? SharePoint REST credential (separate audience from Graph) ??"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"???
    private ClientSecretCredential GetSpCredential()
    {
        if (_spCredential is not null) return _spCredential;

        var tenantId     = _config["SharePoint:TenantId"]     ?? throw new InvalidOperationException("SharePoint:TenantId not set");
        var clientId     = _config["SharePoint:ClientId"]     ?? throw new InvalidOperationException("SharePoint:ClientId not set");
        var clientSecret = _config["SharePoint:ClientSecret"] ?? throw new InvalidOperationException("SharePoint:ClientSecret not set");

        _spCredential = new ClientSecretCredential(tenantId, clientId, clientSecret);
        return _spCredential;
    }

    /// <summary>
    /// Resolves the site ID and all frequently-used list IDs in parallel, and issues one
    /// throwaway Graph query to establish the HTTP/2 connection + cache the OAuth token.
    /// Called once from Program.cs at app startup  --  subsequent user requests skip ~500ms
    /// of one-time warmup cost.
    /// </summary>
    public async Task PrewarmAsync(CancellationToken ct = default)
    {
        try
        {
            var siteId = await GetSiteIdAsync();

            // Resolve list IDs in parallel. Missing lists are OK here  --  they'll fail on
            // first real use with a clearer error if setup-supplier-lists wasn't run yet.
            var tasks = new List<Task>
            {
                TrySwallowAsync(() => GetSupplierResponsesListIdAsync()),
                TrySwallowAsync(() => GetConversationsListIdAsync()),
                TrySwallowAsync(() => GetGraph().Sites[siteId].Lists
                    .GetAsync(r => r.QueryParameters.Top = 1, ct)),
            };
            await Task.WhenAll(tasks);

            _log.LogInformation("[SP] Pre-warm complete");
        }
        catch (Exception ex)
        {
            _log.LogWarning(ex, "[SP] Pre-warm failed (non-fatal  --  first request will warm cold)");
        }
    }

    private static async Task TrySwallowAsync(Func<Task> f)
    {
        try { await f(); } catch { /* first-request will retry with proper error */ }
    }

    // ??"?????"??? Site ID (cached) ??"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"???
    private async Task<string> GetSiteIdAsync()
    {
        if (_siteId is not null) return _siteId;

        var siteUrl = _config["SharePoint:SiteUrl"]
            ?? "https://metalsupermarkets-my.sharepoint.com/personal/angus_mithrilmetals_com";

        var uri  = new Uri(siteUrl);
        var host = uri.Host;
        var path = uri.AbsolutePath;

        _log.LogInformation("[SP] Resolving site: {Host}{Path}", host, path);

        var site = await GetGraph().Sites[$"{host}:{path}"].GetAsync();
        _siteId  = site!.Id ?? throw new Exception("Could not resolve SharePoint site ID");

        _log.LogInformation("[SP] Site ID: {Id}", _siteId);
        return _siteId;
    }

    // ??"?????"??? List ID getters (cached) ??"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"???

    private async Task<string> GetSupplierResponsesListIdAsync()
    {
        if (_srListId is not null) return _srListId;
        _srListId = await ResolveListIdAsync("SupplierResponses");
        return _srListId;
    }

    private async Task<string> GetSupplierLineItemsListIdAsync()
    {
        if (_sliListId is not null) return _sliListId;
        _sliListId = await ResolveListIdAsync("SupplierLineItems");
        return _sliListId;
    }

    /// <summary>
    /// Returns all SupplierResponses rows, serving from an in-memory cache for up to 60 s.
    /// Concurrent callers wait on the semaphore; only the first fetches from Graph.
    /// Call <see cref="InvalidateSrCache"/> after any write to SupplierResponses.
    /// </summary>
    private async Task<List<Microsoft.Graph.Models.ListItem>> GetCachedSrItemsAsync(
        string siteId, string srListId)
    {
        if (_srRowCache is not null && DateTime.UtcNow < _srRowCacheExpiry)
            return _srRowCache;

        await _srRowCacheLock.WaitAsync();
        try
        {
            if (_srRowCache is not null && DateTime.UtcNow < _srRowCacheExpiry)
                return _srRowCache;

            _log.LogDebug("[SP] SR cache miss  --  fetching all SupplierResponses from Graph");
            var resp = await GetGraph().Sites[siteId].Lists[srListId].Items
                .GetAsync(req => { req.QueryParameters.Expand = ["fields"]; req.QueryParameters.Top = 5000; });
            _srRowCache    = resp?.Value ?? [];
            _srRowCacheExpiry = DateTime.UtcNow.AddSeconds(60);
            _log.LogDebug("[SP] SR cache populated: {Count} rows", _srRowCache.Count);
            return _srRowCache;
        }
        finally
        {
            _srRowCacheLock.Release();
        }
    }

    private void InvalidateSrCache()
    {
        _srRowCache       = null;
        _srRowCacheExpiry = DateTime.MinValue;
    }

    private async Task<string> GetRfqReferencesListIdAsync()
    {
        if (_rfqRefListId is not null) return _rfqRefListId;
        _rfqRefListId = await ResolveListIdAsync("RFQ References");
        return _rfqRefListId;
    }

    // Legacy  -- used by EnsureColumnsAsync only
    private async Task<string> GetListIdAsync()
    {
        if (_listId is not null) return _listId;
        _listId = await ResolveListIdAsync(_config["SharePoint:ListName"] ?? "RFQ Line Items");
        return _listId;
    }

    private async Task<string> ResolveListIdAsync(string listName)
    {
        var siteId = await GetSiteIdAsync();
        var lists  = await GetGraph().Sites[siteId].Lists
            .GetAsync(req => req.QueryParameters.Filter = $"displayName eq '{listName}'");
        var list = lists?.Value?.FirstOrDefault()
            ?? throw new Exception($"SharePoint list '{listName}' not found. Run POST /api/setup-supplier-lists.");
        var id = list.Id ?? throw new Exception($"List '{listName}' ID was null");
        _log.LogInformation("[SP] List '{Name}' -> id: {Id}", listName, id);
        return id;
    }

    // ??"?????"??? Read: SupplierLineItems joined with SupplierResponses ??"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"???

    /// <summary>
    /// Returns up to <paramref name="top"/> SupplierLineItems merged with their parent
    /// SupplierResponse fields as flat field dictionaries (matches the shape expected
    /// by the Shredder dashboard).
    /// </summary>
    /// <summary>
    /// Fetches one page of SupplierLineItems (joined with SupplierResponses) and returns
    /// the processed rows together with the Graph <c>@odata.nextLink</c> URL for the next page
    /// (or <c>null</c> when all data has been returned).
    ///
    /// Pass <paramref name="sliNextLink"/> = <c>null</c> for the first page; on subsequent calls
    /// pass back the value returned in the previous response.  Graph SharePoint list items do not
    /// support <c>$skip</c>  -- cursor-based pagination via <c>@odata.nextLink</c> is the only
    /// supported approach.
    /// </summary>
    public async Task<(List<Dictionary<string, object?>> Items, string? NextLink)>
        ReadSupplierItemsAsync(int top = 5000, string? sliNextLink = null, bool skipDedup = false)
    {
        var siteId    = await GetSiteIdAsync();
        var srListId  = await GetSupplierResponsesListIdAsync();
        var sliListId = await GetSupplierLineItemsListIdAsync();

        // SR rows served from cache  --  all pages of SLI share one cached fetch.
        var srTask = GetCachedSrItemsAsync(siteId, srListId);

        // SLI: first page uses standard request; subsequent pages follow the @odata.nextLink.
        // Graph SharePoint list items don't support $skip  -- cursor pagination only.
        Task<Microsoft.Graph.Models.ListItemCollectionResponse?> sliTask;
        if (sliNextLink is null)
        {
            sliTask = GetGraph().Sites[siteId].Lists[sliListId].Items
                .GetAsync(req =>
                {
                    req.QueryParameters.Expand = ["fields"];
                    req.QueryParameters.Top    = top;
                });
        }
        else
        {
            // Construct a request builder from the raw nextLink URL  -- the SDK injects auth
            // automatically and the URL already carries all required query parameters.
            var nextBuilder = new Microsoft.Graph.Sites.Item.Lists.Item.Items.ItemsRequestBuilder(
                sliNextLink, GetGraph().RequestAdapter);
            sliTask = nextBuilder.GetAsync();
        }

        await Task.WhenAll(srTask, sliTask);

        var srItems     = srTask.Result;
        var sliResponse = sliTask.Result;
        var nextLink    = sliResponse?.OdataNextLink;   // null on last page

        _log.LogDebug("[SP] ReadSupplierItems: {SrCount} SR rows, {SliCount} SLI rows",
            srItems.Count, sliResponse?.Value?.Count ?? 0);

        var result = JoinSliToSr(sliResponse?.Value ?? [], srItems);

        if (skipDedup) return (result, nextLink);

        // ??"?????"??? Deduplicate by (SupplierResponseId, normalised ProductName) ??"?????"?????"?????"?????"?????"???
        // Pass 1: exact normalised-name dedup (whitespace/case/decimal variants).
        // Pass 2: fuzzy dedup within each SrId group  -- catches abbreviation variants
        //         like "HR Flat Bar" vs "Hot Rolled Flat Bar" that share the same
        //         numeric tokens and have Jaccard ??????? 0.5.  Keeps the longer name.
        static string NormProd(string? s)
        {
            if (string.IsNullOrWhiteSpace(s)) return "";
            s = s.Trim().ToLowerInvariant();
            s = System.Text.RegularExpressions.Regex.Replace(s, @"(?<!\d)\.(\d)", "0.$1");
            return System.Text.RegularExpressions.Regex.Replace(s, @"\s+", " ");
        }

        result = result
            .GroupBy(r => (
                SrId: r.TryGetValue("SupplierResponseId", out var sid) ? sid?.ToString() ?? "" : "",
                Prod: NormProd(r.TryGetValue("ProductName", out var pn) ? pn?.ToString() : null)
            ))
            .Select(g =>
            {
                if (g.Count() == 1) return g.First();
                _log.LogWarning("[SP] Dedup: {Count} SLI rows with SrId={SrId} product='{Prod}'  -- keeping most-populated",
                    g.Count(), g.Key.SrId, g.Key.Prod);
                return g.OrderByDescending(r => r.Count(kv => kv.Value is not null)).First();
            })
            .ToList();

        // Pass 2: fuzzy dedup within each SupplierResponseId group
        static string GetProd(Dictionary<string, object?> r) =>
            r.TryGetValue("ProductName", out var p) ? p?.ToString() ?? "" : "";

        var fuzzyResult = new List<Dictionary<string, object?>>();
        foreach (var srGroup in result.GroupBy(r =>
            r.TryGetValue("SupplierResponseId", out var sid) ? sid?.ToString() ?? "" : ""))
        {
            var rows = srGroup.ToList();
            if (srGroup.Key.Length == 0 || rows.Count == 1) { fuzzyResult.AddRange(rows); continue; }

            // Greedy cluster: for each row, merge into the first compatible accepted row,
            // or add as a new representative. Keep the longer product name.
            var accepted = new List<Dictionary<string, object?>>();
            foreach (var row in rows)
            {
                var rowTok = ProductTokens(GetProd(row));
                bool merged = false;
                for (int i = 0; i < accepted.Count; i++)
                {
                    var accTok = ProductTokens(GetProd(accepted[i]));
                    if (NumericTokensCompatible(rowTok, accTok) && ProductJaccard(rowTok, accTok) >= 0.5)
                    {
                        _log.LogWarning("[SP] Fuzzy-dedup: merging '{Row}' into '{Acc}' (SrId={SrId})",
                            GetProd(row), GetProd(accepted[i]), srGroup.Key);
                        // Keep the row with the longer product name as it is more descriptive
                        if (GetProd(row).Length > GetProd(accepted[i]).Length)
                            accepted[i] = row;
                        merged = true;
                        break;
                    }
                }
                if (!merged) accepted.Add(row);
            }
            fuzzyResult.AddRange(accepted);
        }
        result = fuzzyResult;

        return (result, nextLink);
    }

    // ??"?????"??? Read: all supplier items for one RFQ ID (targeted refresh) ??"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"???

    /// <summary>
    /// Fetches all SupplierLineItems for a specific RFQ ID, joined with their parent
    /// SupplierResponses.  Returns the same flat dict shape as <see cref="ReadSupplierItemsAsync"/>
    /// but scoped to one job  -- used for targeted UI refresh after a Service Bus notification.
    /// </summary>
    public async Task<List<Dictionary<string, object?>>> ReadSupplierItemsByRfqIdAsync(string rfqId)
    {
        var siteId    = await GetSiteIdAsync();
        var srListId  = await GetSupplierResponsesListIdAsync();
        var sliListId = await GetSupplierLineItemsListIdAsync();
        var sliCol    = await ResolveRfqIdColumnAsync(siteId, sliListId);

        // SR rows served from cache  --  no full-table fetch for targeted refreshes.
        var srTask = GetCachedSrItemsAsync(siteId, srListId);

        // Fetch only SLI rows for this rfqId via OData $filter.
        var sliTask = GetGraph().Sites[siteId].Lists[sliListId].Items
            .GetAsync(req =>
            {
                req.QueryParameters.Expand = ["fields"];
                req.QueryParameters.Filter = $"fields/{sliCol} eq '{rfqId}'";
                req.QueryParameters.Top    = 500;
            });

        await Task.WhenAll(srTask, sliTask);

        var result = JoinSliToSr(sliTask.Result?.Value ?? [], srTask.Result);

        _log.LogInformation("[SP] ReadSupplierItemsByRfqId({RfqId}): {Count} items", rfqId, result.Count);
        return result;
    }

    // -- Shared SLI join helper --

    /// <summary>
    /// Builds lookup tables from a set of SR rows and joins each SLI item against them,
    /// returning flat merged dictionaries ready for serialisation.
    /// Used by both ReadSupplierItemsAsync and ReadSupplierItemsByRfqIdAsync.
    /// </summary>
    private List<Dictionary<string, object?>> JoinSliToSr(
        IEnumerable<Microsoft.Graph.Models.ListItem> sliItems,
        List<Microsoft.Graph.Models.ListItem>        srItems)
    {
        var srById = srItems
            .Where(i => i.Id is not null && i.Fields?.AdditionalData is not null)
            .ToDictionary(i => i.Id!, i => i.Fields!.AdditionalData!);

        var srCreatedAt = srItems
            .Where(i => i.Id is not null && i.CreatedDateTime.HasValue)
            .ToDictionary(i => i.Id!, i => i.CreatedDateTime!.Value.UtcDateTime);

        var srBySupplierRfq = new Dictionary<string, (string SrId, IDictionary<string, object> Fields)>(StringComparer.OrdinalIgnoreCase);
        foreach (var (srItemId, srRaw) in srById)
        {
            var rfq = RfqIdRaw(srRaw);
            var sn  = GetStrRaw(srRaw, "SupplierName");
            if (rfq is not null && sn is not null)
                srBySupplierRfq.TryAdd($"{rfq}|{sn}", (srItemId, srRaw));
        }

        var result = new List<Dictionary<string, object?>>();
        foreach (var sli in sliItems)
        {
            if (sli.Fields?.AdditionalData is null) continue;

            var row = sli.Fields.AdditionalData
                .Where(kv => IsAppField(kv.Key))
                .ToDictionary(kv => kv.Key, kv => (object?)kv.Value);

            if (sli.Id is not null) row["SpItemId"] = sli.Id;

            IDictionary<string, object>? srMatch = null;

            // Primary join: SupplierResponseId -> SR item's SP integer ID
            var srId = GetStr(row, "SupplierResponseId");
            if (srId is not null)
            {
                srById.TryGetValue(srId, out srMatch);
                if (srMatch is null)
                    _log.LogDebug("[SP] SLI {SliId}: SupplierResponseId={SrId} not found in srById",
                        sli.Id, srId);
                else if (srCreatedAt.TryGetValue(srId, out var srCa))
                    row["SrCreatedAt"] = srCa;
            }

            // Fallback join: RFQ_ID + SupplierName (handles stale/missing SupplierResponseId)
            if (srMatch is null)
            {
                var sliRfq = RfqId(row);
                var sliSn  = GetStr(row, "SupplierName");
                if (sliRfq is not null && sliSn is not null &&
                    srBySupplierRfq.TryGetValue($"{sliRfq}|{sliSn}", out var fb))
                {
                    srMatch = fb.Fields;
                    row["SupplierResponseId"] = fb.SrId;
                    if (srCreatedAt.TryGetValue(fb.SrId, out var srCa))
                        row["SrCreatedAt"] = srCa;
                    _log.LogDebug("[SP] SLI {SliId} [{Rfq}/{Supplier}]: joined via fallback, corrected SrId {OldId}->{NewId}",
                        sli.Id, sliRfq, sliSn, srId ?? "null", fb.SrId);
                }
            }

            if (srMatch is not null)
            {
                // Prefer the SLI's own RFQ_ID when it was individually reparented
                var sliOwnRfqId = RfqId(row);
                var rfqIdVal = (!string.IsNullOrEmpty(sliOwnRfqId) && sliOwnRfqId != "000000")
                    ? sliOwnRfqId
                    : RfqIdRaw(srMatch);
                if (rfqIdVal is not null) row["JobReference"] = rfqIdVal;

                foreach (var f in ParentFields)
                    if (!row.ContainsKey(f) && srMatch.TryGetValue(f, out var v))
                        row[f] = v;

                // If SR is blanket-regret and line item has no pricing, inherit regret flag
                if (srMatch.TryGetValue("IsRegret", out var srRegret) &&
                    srRegret is true or JsonElement { ValueKind: JsonValueKind.True })
                {
                    if (!row.ContainsKey("PricePerPound") || row["PricePerPound"] is null)
                        if (!row.ContainsKey("PricePerFoot") || row["PricePerFoot"] is null)
                            row["IsRegret"] = true;
                }
            }

            result.Add(row);
        }
        return result;
    }

    // -- Read: new supplier activity since a timestamp --

    /// <summary>
    /// Returns SupplierLineItems created after <paramref name="since"/>, grouped into
    /// per-supplier activities.  Used by the 5-second UI poll to detect new quotes.
    /// </summary>
    public async Task<ChangesResult> GetNewResponsesSinceAsync(DateTime since)
    {
        var siteId    = await GetSiteIdAsync();
        var sliListId = await GetSupplierLineItemsListIdAsync();
        var sinceStr  = since.ToUniversalTime().ToString("yyyy-MM-ddTHH:mm:ssZ");
        var serverTime = DateTime.UtcNow;

        // SP OData $filter on createdDateTime is unreliable when combined with $expand.
        // Instead: fetch the 200 most-recently-created items (by SP id desc) and
        // filter client-side using the listItem.CreatedDateTime property.
        var page = await GetGraph().Sites[siteId].Lists[sliListId].Items
            .GetAsync(req =>
            {
                req.QueryParameters.Expand  = ["fields"];
                req.QueryParameters.Orderby = ["id desc"];
                req.QueryParameters.Top     = 200;
            });

        var activities = (page?.Value ?? [])
            .Where(item => item.CreatedDateTime.HasValue &&
                           item.CreatedDateTime.Value.UtcDateTime > since)
            .Select(item =>
            {
                var d = item.Fields?.AdditionalData;
                if (d is null) return ((string rfqId, string supplier, string product, decimal price)?)null;
                var rfqId    = RfqId(d!) ?? "";
                var supplier = GetStr(d!, "SupplierName") ?? "";
                var product  = GetStr(d!, "ProductName")  ?? "";
                var priceStr = GetStr(d!, "TotalPrice");
                decimal.TryParse(priceStr, System.Globalization.NumberStyles.Any,
                                 System.Globalization.CultureInfo.InvariantCulture, out var price);
                return ((string rfqId, string supplier, string product, decimal price)?)
                       (rfqId, supplier, product, price);
            })
            .Where(x => x != null && x!.Value.rfqId.Length > 0 && !string.IsNullOrEmpty(x!.Value.supplier))
            .GroupBy(x => (x!.Value.rfqId, x!.Value.supplier))
            .Select(g => new SupplierActivity
            {
                SupplierName = g.Key.supplier,
                RfqId        = g.Key.rfqId,
                Products     = g.Select(x => new ActivityProduct
                {
                    Name       = x!.Value.product,
                    TotalPrice = x!.Value.price
                }).ToList()
            })
            .ToList();

        _log.LogInformation("[SP] GetNewResponsesSince({Since}): scanned {Total} SLI rows, {New} new ?????' {Groups} activities",
            sinceStr, page?.Value?.Count ?? 0, activities.Sum(a => a.Products.Count), activities.Count);

        return new ChangesResult { Activities = activities, ServerTime = serverTime };
    }

    // ??"?????"??? Read: new RFQ References since a timestamp ??"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"???

    /// <summary>
    /// Returns RFQ References created after <paramref name="since"/>, with their
    /// associated RFQ Line Items (also created after that timestamp).
    /// Used by the 5-second UI poll to show "Name requested [X] on Date" toasts.
    /// </summary>
    public async Task<List<NewRfqActivity>> GetNewRfqReferencesSinceAsync(DateTime since)
    {
        var siteId    = await GetSiteIdAsync();
        var refListId = await GetRfqReferencesListIdAsync();
        var lirListId = await GetRfqLineItemsListIdAsync();
        var refCol    = await ResolveRfqIdColumnAsync(siteId, refListId);
        var lirCol    = await ResolveRfqIdColumnAsync(siteId, lirListId);
        var sinceStr  = since.ToUniversalTime().ToString("yyyy-MM-ddTHH:mm:ssZ");

        // Same pattern as GetNewResponsesSinceAsync: $orderby=id desc + $top + client-side date filter.
        var refsTask = GetGraph().Sites[siteId].Lists[refListId].Items
            .GetAsync(r => { r.QueryParameters.Expand = ["fields"]; r.QueryParameters.Orderby = ["id desc"]; r.QueryParameters.Top = 50; });
        var lirTask  = GetGraph().Sites[siteId].Lists[lirListId].Items
            .GetAsync(r => { r.QueryParameters.Expand = ["fields"]; r.QueryParameters.Orderby = ["id desc"]; r.QueryParameters.Top = 200; });

        await Task.WhenAll(refsTask, lirTask);

        // Parse new RFQ References  -- client-side date filter
        var newRefs = (refsTask.Result?.Value ?? [])
            .Where(i => i.Fields?.AdditionalData is not null &&
                        i.CreatedDateTime.HasValue &&
                        i.CreatedDateTime.Value.UtcDateTime > since)
            .Select(i =>
            {
                var d     = i.Fields!.AdditionalData!;
                var rfqId = GetStr(d, refCol) ?? RfqId(d);
                if (string.IsNullOrEmpty(rfqId)) return null;
                var requester = GetStr(d, "Requester") ?? "";
                DateTime? dateSent = null;
                var dcStr = GetStr(d, "DateCreated");
                if (dcStr is not null && DateTime.TryParse(dcStr, null,
                        System.Globalization.DateTimeStyles.RoundtripKind, out var dt))
                    dateSent = dt;
                return ((string RfqId, string Requester, DateTime? DateSent)?)(rfqId, requester, dateSent);
            })
            .Where(x => x is not null)
            .ToList();

        if (newRefs.Count == 0) return [];

        // Group new line items by RFQ_ID  -- client-side date filter
        var lirByRfq = (lirTask.Result?.Value ?? [])
            .Where(i => i.Fields?.AdditionalData is not null &&
                        i.CreatedDateTime.HasValue &&
                        i.CreatedDateTime.Value.UtcDateTime > since)
            .Select(i =>
            {
                var d     = i.Fields!.AdditionalData!;
                var rfqId = GetStr(d, lirCol) ?? RfqId(d);
                if (string.IsNullOrEmpty(rfqId)) return null;
                return ((string RfqId, RfqLineItemSummary Item)?)(rfqId, new RfqLineItemSummary
                {
                    Mspc        = GetStr(d, "MSPC"),
                    Product     = GetStr(d, "Product"),
                    Units       = GetStr(d, "Units"),
                    SizeOfUnits = GetStr(d, "SizeOfUnits"),
                });
            })
            .Where(x => x is not null)
            .GroupBy(x => x!.Value.RfqId, StringComparer.OrdinalIgnoreCase)
            .ToDictionary(g => g.Key, g => g.Select(x => x!.Value.Item).ToList(),
                          StringComparer.OrdinalIgnoreCase);

        var result = newRefs
            .Select(r => new NewRfqActivity
            {
                RfqId     = r!.Value.RfqId,
                Requester = r!.Value.Requester,
                DateSent  = r!.Value.DateSent,
                LineItems = lirByRfq.TryGetValue(r!.Value.RfqId, out var items) ? items : [],
            })
            .ToList();

        _log.LogDebug("[SP] GetNewRfqReferencesSince({Since}): {Count} new RFQ(s)", sinceStr, result.Count);
        return result;
    }

    // ??"?????"??? Read: RFQ References (for Notes) ??"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"???

    public async Task<List<Dictionary<string, object?>>> ReadRfqReferencesAsync()
    {
        var siteId = await GetSiteIdAsync();
        var listId = await GetRfqReferencesListIdAsync();
        var col    = await ResolveRfqIdColumnAsync(siteId, listId);

        var result = await GetGraph().Sites[siteId].Lists[listId].Items
            .GetAsync(req =>
            {
                req.QueryParameters.Expand = [$"fields($select={col},Notes,Requester,DateCreated,EmailRecipients,Complete,Flagged)"];
                req.QueryParameters.Top    = 500;
            });

        static string? FieldStr(IDictionary<string, object?> d, string key) =>
            d.TryGetValue(key, out var v) ? v?.ToString() : null;

        return result?.Value?
            .Where(i => i.Fields?.AdditionalData is not null)
            .Select(i =>
            {
                var d = i.Fields!.AdditionalData!;
                var rfqId = FieldStr(d, col)
                         ?? FieldStr(d, "RFQ_x005F_ID")
                         ?? FieldStr(d, "RFQ_ID");
                return new Dictionary<string, object?>
                {
                    ["Id"]               = i.Id,
                    ["RFQ_ID"]           = rfqId,
                    ["Notes"]            = FieldStr(d, "Notes"),
                    ["Requester"]        = FieldStr(d, "Requester"),
                    ["DateCreated"]      = d.TryGetValue("DateCreated", out var dc) ? dc : null,
                    ["Created"]          = i.CreatedDateTime?.UtcDateTime.ToString("o"),
                    ["EmailRecipients"]  = FieldStr(d, "EmailRecipients"),
                    ["Complete"]         = d.TryGetValue("Complete",    out var co) ? co : null,
                    ["Flagged"]          = d.TryGetValue("Flagged",     out var fl) ? fl : null,
                };
            })
            .Where(d => d["RFQ_ID"] is not null)
            .ToList()
            ?? [];
    }

    // ??"?????"??? Write: update Notes on an RFQ Reference ??"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"???

    public async Task UpdateRfqNotesAsync(string rfqId, string notes)
    {
        var siteId = await GetSiteIdAsync();
        var listId = await GetRfqReferencesListIdAsync();
        var col    = await ResolveRfqIdColumnAsync(siteId, listId);

        // Fetch all refs client-side  -- OData filter on unindexed columns is unreliable
        // and was causing duplicate rows to be created on every note save.
        var allItems = await GetGraph().Sites[siteId].Lists[listId].Items
            .GetAsync(req =>
            {
                req.QueryParameters.Expand = [$"fields($select=id,{col})"];
                req.QueryParameters.Top    = 500;
            });

        var matches = (allItems?.Value ?? [])
            .Where(i => i.Fields?.AdditionalData is { } d &&
                        string.Equals(
                            d.TryGetValue(col, out var v) ? v?.ToString() : null,
                            rfqId, StringComparison.OrdinalIgnoreCase))
            .ToList();

        if (matches.Count == 0)
        {
            _log.LogInformation("[SP] RFQ Reference '{Id}' not found  -- creating it", rfqId);
            await GetGraph().Sites[siteId].Lists[listId].Items
                .PostAsync(new ListItem
                {
                    Fields = new FieldValueSet
                    {
                        AdditionalData = new Dictionary<string, object?>
                        {
                            [col]     = rfqId,
                            ["Notes"] = notes,
                        }
                    }
                });
            _log.LogInformation("[SP] Created RFQ Reference '{Id}' with Notes", rfqId);
            return;
        }

        // Update the primary entry.
        var primary = matches[0];
        await GetGraph().Sites[siteId].Lists[listId].Items[primary.Id!].Fields
            .PatchAsync(new FieldValueSet
            {
                AdditionalData = new Dictionary<string, object?> { ["Notes"] = notes }
            });
        _log.LogInformation("[SP] Updated Notes for RFQ '{Id}'", rfqId);

        // Delete any duplicate entries found alongside the primary.
        foreach (var dupe in matches.Skip(1))
        {
            await GetGraph().Sites[siteId].Lists[listId].Items[dupe.Id!].DeleteAsync();
            _log.LogWarning("[SP] Deleted duplicate RFQ Reference '{Id}' (item {ItemId})", rfqId, dupe.Id);
        }
    }

    // ??"?????"??? Write: update Requester on an RFQ Reference ??"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"???

    public async Task UpdateRfqRequesterAsync(string rfqId, string requester)
    {
        var siteId = await GetSiteIdAsync();
        var listId = await GetRfqReferencesListIdAsync();
        var col    = await ResolveRfqIdColumnAsync(siteId, listId);

        var allItems = await GetGraph().Sites[siteId].Lists[listId].Items
            .GetAsync(req =>
            {
                req.QueryParameters.Expand = [$"fields($select=id,{col})"];
                req.QueryParameters.Top    = 500;
            });

        var matches = (allItems?.Value ?? [])
            .Where(i => i.Fields?.AdditionalData is { } d &&
                        string.Equals(
                            d.TryGetValue(col, out var v) ? v?.ToString() : null,
                            rfqId, StringComparison.OrdinalIgnoreCase))
            .ToList();

        if (matches.Count == 0)
        {
            _log.LogWarning("[SP] RFQ Reference '{Id}' not found  -- cannot update Requester", rfqId);
            return;
        }

        await GetGraph().Sites[siteId].Lists[listId].Items[matches[0].Id!].Fields
            .PatchAsync(new FieldValueSet
            {
                AdditionalData = new Dictionary<string, object?> { ["Requester"] = requester }
            });
        _log.LogInformation("[SP] Updated Requester for RFQ '{Id}' ?????' '{Requester}'", rfqId, requester);
    }

    // ??"?????"??? Write: update Complete flag on an RFQ Reference ??"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"???

    public async Task SetRfqCompleteAsync(string rfqId, bool complete)
    {
        var siteId = await GetSiteIdAsync();
        var listId = await GetRfqReferencesListIdAsync();
        var col    = await ResolveRfqIdColumnAsync(siteId, listId);

        var allItems = await GetGraph().Sites[siteId].Lists[listId].Items
            .GetAsync(req =>
            {
                req.QueryParameters.Expand = [$"fields($select=id,{col})"];
                req.QueryParameters.Top    = 500;
            });

        var matches = (allItems?.Value ?? [])
            .Where(i => i.Fields?.AdditionalData is { } d &&
                        string.Equals(
                            d.TryGetValue(col, out var v) ? v?.ToString() : null,
                            rfqId, StringComparison.OrdinalIgnoreCase))
            .ToList();

        if (matches.Count == 0)
        {
            _log.LogWarning("[SP] SetRfqCompleteAsync: RFQ Reference '{Id}' not found", rfqId);
            return;
        }

        var primary = matches[0];
        await GetGraph().Sites[siteId].Lists[listId].Items[primary.Id!].Fields
            .PatchAsync(new FieldValueSet
            {
                AdditionalData = new Dictionary<string, object?> { ["Complete"] = complete }
            });
        _log.LogInformation("[SP] Set Complete={Complete} for RFQ '{Id}'", complete, rfqId);
    }

    /// <summary>
    /// Patches arbitrary fields on an RFQ Reference row by its SP item ID — skips
    /// the full-list re-scan that the rfqId-based writers do. Used by background
    /// jobs that already hold the item ID from a prior read.
    /// </summary>
    public async Task PatchRfqReferenceByItemIdAsync(
        string itemId, IDictionary<string, object?> fields)
    {
        var siteId = await GetSiteIdAsync();
        var listId = await GetRfqReferencesListIdAsync();
        await GetGraph().Sites[siteId].Lists[listId].Items[itemId].Fields
            .PatchAsync(new FieldValueSet { AdditionalData = new Dictionary<string, object?>(fields) });
    }

    // ??"?????"??? Write: update Flagged flag on an RFQ Reference ??"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"???

    public async Task SetRfqFlaggedAsync(string rfqId, bool flagged)
    {
        var siteId = await GetSiteIdAsync();
        var listId = await GetRfqReferencesListIdAsync();
        var col    = await ResolveRfqIdColumnAsync(siteId, listId);

        var allItems = await GetGraph().Sites[siteId].Lists[listId].Items
            .GetAsync(req =>
            {
                req.QueryParameters.Expand = [$"fields($select=id,{col})"];
                req.QueryParameters.Top    = 500;
            });

        var matches = (allItems?.Value ?? [])
            .Where(i => i.Fields?.AdditionalData is { } d &&
                        string.Equals(
                            d.TryGetValue(col, out var v) ? v?.ToString() : null,
                            rfqId, StringComparison.OrdinalIgnoreCase))
            .ToList();

        if (matches.Count == 0)
        {
            _log.LogWarning("[SP] SetRfqFlaggedAsync: RFQ Reference '{Id}' not found", rfqId);
            return;
        }

        await GetGraph().Sites[siteId].Lists[listId].Items[matches[0].Id!].Fields
            .PatchAsync(new FieldValueSet
            {
                AdditionalData = new Dictionary<string, object?> { ["Flagged"] = flagged }
            });
        _log.LogInformation("[SP] Set Flagged={Flagged} for RFQ '{Id}'", flagged, rfqId);
    }

    /// <summary>
    /// Removes duplicate RFQ Reference rows (same RFQ_ID).
    /// Keeps the entry with Notes (if any), otherwise the oldest item.
    /// Returns the number of rows deleted.
    /// </summary>
    public record SrRawRow(string SpId, string? RfqId, string? EmailFrom, string? ReceivedAt);

    public async Task<List<SrRawRow>> ReadSupplierResponsesRawAsync()
    {
        var siteId = await GetSiteIdAsync();
        var listId = await GetSupplierResponsesListIdAsync();
        var rows   = new List<SrRawRow>();

        var page = await GetGraph().Sites[siteId].Lists[listId].Items
            .GetAsync(req =>
            {
                req.QueryParameters.Expand = ["fields($select=id,RFQ_ID,RFQ_x005F_ID,EmailFrom,ReceivedAt)"];
                req.QueryParameters.Top    = 5000;
            });

        while (page?.Value is not null)
        {
            foreach (var item in page.Value)
            {
                if (item.Id is null) continue;
                var d      = item.Fields?.AdditionalData;
                var rfqId  = d is not null ? (GetStr(d, "RFQ_ID") ?? GetStr(d, "RFQ_x005F_ID")) : null;
                var from   = d is not null ? GetStr(d, "EmailFrom")  : null;
                var recAt  = d is not null ? GetStr(d, "ReceivedAt") : null;
                rows.Add(new SrRawRow(item.Id, rfqId, from, recAt));
            }
            if (page.OdataNextLink is null) break;
            page = await GetGraph().Sites[siteId].Lists[listId].Items
                .WithUrl(page.OdataNextLink).GetAsync();
        }

        return rows;
    }

    /// <summary>
    /// Looks up the supplier name from the SR cache for a given rfqId, optionally
    /// preferring the row whose EmailFrom domain matches the caller's address.
    /// Used as a fallback when the sender's domain is not in the SupplierCacheService
    /// but we have a valid SHR tracking token.
    /// Returns null if no SR exists for the rfqId.
    /// </summary>
    public async Task<string?> ResolveSupplierNameFromSrAsync(string rfqId, string fromAddr)
    {
        var siteId = await GetSiteIdAsync();
        var listId = await GetSupplierResponsesListIdAsync();
        var cached = await GetCachedSrItemsAsync(siteId, listId);

        var fromDomain = fromAddr.Contains('@')
            ? fromAddr[(fromAddr.IndexOf('@') + 1)..].ToLowerInvariant()
            : null;

        string? domainMatch = null;
        string? anyMatch    = null;

        foreach (var item in cached)
        {
            var d = item.Fields?.AdditionalData;
            if (d is null) continue;

            var itemRfq = RfqIdRaw(d);
            if (!string.Equals(itemRfq, rfqId, StringComparison.OrdinalIgnoreCase)) continue;

            var supplierName = GetStrRaw(d, "SupplierName");
            if (string.IsNullOrWhiteSpace(supplierName)) continue;

            anyMatch = supplierName;

            if (fromDomain is not null)
            {
                var emailFrom = GetStrRaw(d, "EmailFrom") ?? "";
                var srDomain  = emailFrom.Contains('@')
                    ? emailFrom[(emailFrom.IndexOf('@') + 1)..].ToLowerInvariant()
                    : null;
                if (srDomain is not null && srDomain == fromDomain)
                {
                    domainMatch = supplierName;
                    break;
                }
            }
        }

        return domainMatch ?? anyMatch;
    }

    public async Task<int> DedupeRfqReferencesAsync()
    {
        var siteId = await GetSiteIdAsync();
        var listId = await GetRfqReferencesListIdAsync();
        var col    = await ResolveRfqIdColumnAsync(siteId, listId);

        var allItems = await GetGraph().Sites[siteId].Lists[listId].Items
            .GetAsync(req =>
            {
                req.QueryParameters.Expand = [$"fields($select=id,{col},Notes)"];
                req.QueryParameters.Top    = 1000;
            });

        var groups = (allItems?.Value ?? [])
            .Where(i => i.Fields?.AdditionalData is not null)
            .Select(i => new
            {
                Item  = i,
                RfqId = i.Fields!.AdditionalData!.TryGetValue(col,      out var v) ? v?.ToString() : null,
                Notes = i.Fields!.AdditionalData!.TryGetValue("Notes",  out var n) ? n?.ToString() : null,
            })
            .Where(x => x.RfqId is not null)
            .GroupBy(x => x.RfqId!, StringComparer.OrdinalIgnoreCase)
            .Where(g => g.Count() > 1)
            .ToList();

        int deleted = 0;
        foreach (var group in groups)
        {
            // Keep: prefer entry with notes, then smallest numeric item ID (oldest).
            var ordered = group
                .OrderByDescending(x => x.Notes?.Length > 0 ? 1 : 0)
                .ThenBy(x => int.TryParse(x.Item.Id, out var n) ? n : int.MaxValue)
                .ToList();

            foreach (var dupe in ordered.Skip(1))
            {
                await GetGraph().Sites[siteId].Lists[listId].Items[dupe.Item.Id!].DeleteAsync();
                _log.LogWarning("[SP] Dedupe: deleted duplicate RFQ Reference '{Id}' (item {ItemId})",
                    group.Key, dupe.Item.Id);
                deleted++;
            }
        }
        return deleted;
    }

    // ??"?????"??? RFQ Import: read / write RFQ References + RFQ Line Items ??"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"???

    private string? _rfqLineItemsListId;

    private async Task<string> GetRfqLineItemsListIdAsync()
    {
        if (_rfqLineItemsListId is not null) return _rfqLineItemsListId;
        _rfqLineItemsListId = await ResolveListIdAsync("RFQ Line Items");
        return _rfqLineItemsListId;
    }

    /// <summary>Returns all RFQ_ID values that exist in the RFQ References list.</summary>
    public async Task<HashSet<string>> GetExistingRfqIdsAsync()
    {
        var siteId = await GetSiteIdAsync();
        var listId = await GetRfqReferencesListIdAsync();
        var col    = await ResolveRfqIdColumnAsync(siteId, listId);

        var items = await GetGraph().Sites[siteId].Lists[listId].Items
            .GetAsync(req =>
            {
                req.QueryParameters.Expand = [$"fields($select={col})"];
                req.QueryParameters.Top    = 5000;
            });

        var ids = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        foreach (var item in items?.Value ?? [])
        {
            if (item.Fields?.AdditionalData is null) continue;
            var d  = item.Fields.AdditionalData;
            // Try the resolved name first, then both fallback variants
            string? id = d.TryGetValue(col,            out var v0) ? v0?.ToString()
                       : d.TryGetValue("RFQ_x005F_ID", out var v1) ? v1?.ToString()
                       : d.TryGetValue("RFQ_ID",        out var v2) ? v2?.ToString()
                       : null;
            if (!string.IsNullOrEmpty(id)) ids.Add(id);
        }
        return ids;
    }

    /// <summary>
    /// Upserts one row in the RFQ References list.
    /// If the row does not exist it is created with all fields.
    /// If it already exists, any blank Requester / DateCreated / EmailRecipients fields
    /// are filled in from <paramref name="req"/>  -- populated fields are left untouched.
    /// </summary>
    public async Task CreateRfqReferenceAsync(RfqReferenceRequest req)
    {
        var siteId = await GetSiteIdAsync();
        var listId = await GetRfqReferencesListIdAsync();
        var col    = await ResolveRfqIdColumnAsync(siteId, listId);

        // Fetch existing items for this RFQ_ID (same client-side filter approach as UpdateRfqNotesAsync
        //  -- OData filter on unindexed columns is unreliable).
        var allItems = await GetGraph().Sites[siteId].Lists[listId].Items
            .GetAsync(req2 =>
            {
                req2.QueryParameters.Expand = [$"fields($select=id,{col},Requester,DateCreated,EmailRecipients)"];
                req2.QueryParameters.Top    = 500;
            });

        var matches = (allItems?.Value ?? [])
            .Where(i => i.Fields?.AdditionalData is { } d &&
                        string.Equals(
                            d.TryGetValue(col, out var v) ? v?.ToString() : null,
                            req.RfqId, StringComparison.OrdinalIgnoreCase))
            .ToList();

        if (matches.Count == 0)
        {
            // New  -- create a full row.
            await GetGraph().Sites[siteId].Lists[listId].Items
                .PostAsync(new ListItem
                {
                    Fields = new FieldValueSet
                    {
                        AdditionalData = new Dictionary<string, object?>
                        {
                            [col]               = req.RfqId,
                            ["Requester"]       = req.Requester,
                            ["DateCreated"]     = (req.DateSent == default ? DateTime.UtcNow : req.DateSent.ToUniversalTime()).ToString("o"),
                            ["EmailRecipients"] = req.EmailRecipients,
                        }
                    }
                });
            _log.LogInformation("[SP] Created RFQ Reference '{Id}'", req.RfqId);
            return;
        }

        // Existing  -- patch only fields that are currently blank.
        var primary = matches[0];
        var data    = primary.Fields?.AdditionalData ?? new Dictionary<string, object?>();

        static bool IsBlank(IDictionary<string, object?> d, string key) =>
            !d.TryGetValue(key, out var v) || v is null || string.IsNullOrWhiteSpace(v.ToString());

        var patch = new Dictionary<string, object?>();

        if (IsBlank(data, "Requester") && !string.IsNullOrWhiteSpace(req.Requester))
            patch["Requester"] = req.Requester;

        if (IsBlank(data, "DateCreated") && req.DateSent != default)
            patch["DateCreated"] = req.DateSent.ToUniversalTime().ToString("o");

        if (IsBlank(data, "EmailRecipients") && !string.IsNullOrWhiteSpace(req.EmailRecipients))
            patch["EmailRecipients"] = req.EmailRecipients;

        if (patch.Count > 0)
        {
            await GetGraph().Sites[siteId].Lists[listId].Items[primary.Id!].Fields
                .PatchAsync(new FieldValueSet { AdditionalData = patch });
            _log.LogInformation("[SP] Updated missing fields for RFQ Reference '{Id}': {Fields}",
                req.RfqId, string.Join(", ", patch.Keys));
        }
        else
        {
            _log.LogInformation("[SP] RFQ Reference '{Id}' already complete  -- no update needed", req.RfqId);
        }
    }

    /// <summary>Returns the set of RFQ_ID values that already have at least one row in the RFQ Line Items list.</summary>
    public async Task<HashSet<string>> GetRfqIdsWithLineItemsAsync()
    {
        var siteId = await GetSiteIdAsync();
        var listId = await GetRfqLineItemsListIdAsync();
        var col    = await ResolveRfqIdColumnAsync(siteId, listId);

        var items = await GetGraph().Sites[siteId].Lists[listId].Items
            .GetAsync(req =>
            {
                req.QueryParameters.Expand = [$"fields($select={col})"];
                req.QueryParameters.Top    = 5000;
            });

        var ids = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        foreach (var item in items?.Value ?? [])
        {
            if (item.Fields?.AdditionalData is null) continue;
            var d = item.Fields.AdditionalData;
            string? id = d.TryGetValue(col,            out var v0) ? v0?.ToString()
                       : d.TryGetValue("RFQ_x005F_ID", out var v1) ? v1?.ToString()
                       : d.TryGetValue("RFQ_ID",        out var v2) ? v2?.ToString()
                       : null;
            if (!string.IsNullOrEmpty(id)) ids.Add(id);
        }
        return ids;
    }

    /// <summary>
    /// Returns all rows from the RFQ Line Items list.
    /// Used by the Shredder dashboard to display requested items under each RFQ group header.
    /// </summary>
    public async Task<List<(string RfqId, string? Mspc, string? Product, string? Units, string? SizeOfUnits, bool IsPurchased, string? PoNumber)>> ReadAllRfqLineItemsAsync()
    {
        var siteId = await GetSiteIdAsync();
        var listId = await GetRfqLineItemsListIdAsync();
        var col    = await ResolveRfqIdColumnAsync(siteId, listId);

        var items = await GetGraph().Sites[siteId].Lists[listId].Items
            .GetAsync(req =>
            {
                req.QueryParameters.Expand = [$"fields($select={col},MSPC,Product,Units,SizeOfUnits,IsPurchased,PoNumber)"];
                req.QueryParameters.Top    = 5000;
            });

        var result = new List<(string, string?, string?, string?, string?, bool, string?)>();
        foreach (var item in items?.Value ?? [])
        {
            if (item.Fields?.AdditionalData is null) continue;
            var d = item.Fields.AdditionalData;
            string? id = d.TryGetValue(col,            out var v0) ? v0?.ToString()
                       : d.TryGetValue("RFQ_x005F_ID", out var v1) ? v1?.ToString()
                       : d.TryGetValue("RFQ_ID",        out var v2) ? v2?.ToString()
                       : null;
            if (string.IsNullOrEmpty(id)) continue;
            var mspc    = d.TryGetValue("MSPC",        out var vm) ? vm?.ToString() : null;
            var product = d.TryGetValue("Product",     out var vp) ? vp?.ToString() : null;
            var size    = d.TryGetValue("SizeOfUnits", out var vs) ? vs?.ToString() : null;
            string? units = null;
            if (d.TryGetValue("Units", out var vu) && vu is not null)
                units = vu is System.Text.Json.JsonElement je ? je.ToString() : vu.ToString();
            var isPurchased = d.TryGetValue("IsPurchased", out var ip) &&
                              ip is true or JsonElement { ValueKind: JsonValueKind.True };
            var poNumber = d.TryGetValue("PoNumber", out var vpo) ? vpo?.ToString() : null;
            result.Add((id, mspc, product, units, size, isPurchased, poNumber));
        }
        return result;
    }

    /// <summary>
    /// Returns the RFQ Line Items for a specific RFQ ID  -- used to inject requested-item
    /// context into the AI extraction prompt (RLI anchoring).
    /// </summary>
    public async Task<List<RliContextItem>> ReadRfqLineItemsByRfqIdAsync(string rfqId)
    {
        var siteId = await GetSiteIdAsync();
        var listId = await GetRfqLineItemsListIdAsync();
        var col    = await ResolveRfqIdColumnAsync(siteId, listId);

        var items = await GetGraph().Sites[siteId].Lists[listId].Items
            .GetAsync(req =>
            {
                req.QueryParameters.Expand = [$"fields($select={col},MSPC,Product)"];
                req.QueryParameters.Filter = $"fields/{col} eq '{rfqId}'";
                req.QueryParameters.Top    = 100;
            });

        var result = new List<RliContextItem>();
        foreach (var item in items?.Value ?? [])
        {
            if (item.Fields?.AdditionalData is null) continue;
            var d = item.Fields.AdditionalData;
            string? id = d.TryGetValue(col,            out var v0) ? v0?.ToString()
                       : d.TryGetValue("RFQ_x005F_ID", out var v1) ? v1?.ToString()
                       : d.TryGetValue("RFQ_ID",        out var v2) ? v2?.ToString()
                       : null;
            if (!string.Equals(id, rfqId, StringComparison.OrdinalIgnoreCase)) continue;
            var mspc    = d.TryGetValue("MSPC",    out var vm) ? vm?.ToString() : null;
            var product = d.TryGetValue("Product", out var vp) ? vp?.ToString() : null;
            if (!string.IsNullOrEmpty(product) || !string.IsNullOrEmpty(mspc))
                result.Add(new RliContextItem { Mspc = mspc, ProductName = product });
        }
        return result;
    }

    // ??"?????"??? RLI anchoring dry-run helpers ??"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"???

    /// <summary>
    /// Returns SLI rows (product name + current matched MSPC) for a given RFQ ID.
    /// Used by the RLI anchoring dry-run endpoint to compare existing fuzzy matches
    /// against what RLI anchoring would produce.
    /// </summary>
    public async Task<List<SliCompactRow>> ReadSliCompactByRfqIdAsync(string rfqId)
    {
        var siteId = await GetSiteIdAsync();
        var listId = await GetSupplierLineItemsListIdAsync();
        var col    = await ResolveRfqIdColumnAsync(siteId, listId);

        var items = await GetGraph().Sites[siteId].Lists[listId].Items
            .GetAsync(req =>
            {
                req.QueryParameters.Expand = [$"fields($select={col},SupplierName,ProductName,SupplierProductName,ProductSearchKey,CatalogProductName)"];
                req.QueryParameters.Filter = $"fields/{col} eq '{rfqId}'";
                req.QueryParameters.Top    = 200;
            });

        var result = new List<SliCompactRow>();
        foreach (var item in items?.Value ?? [])
        {
            if (item.Fields?.AdditionalData is null) continue;
            var d = item.Fields.AdditionalData;
            var id = d.TryGetValue(col, out var v0) ? v0?.ToString()
                   : d.TryGetValue("RFQ_x005F_ID", out var v1) ? v1?.ToString()
                   : d.TryGetValue("RFQ_ID", out var v2) ? v2?.ToString() : null;
            if (!string.Equals(id, rfqId, StringComparison.OrdinalIgnoreCase)) continue;
            result.Add(new SliCompactRow
            {
                SupplierName     = d.TryGetValue("SupplierName",     out var sn)  ? sn?.ToString()  : null,
                ProductName      = d.TryGetValue("ProductName",      out var pn)  ? pn?.ToString()  : null,
                SupplierProductName = d.TryGetValue("SupplierProductName", out var spn) ? spn?.ToString() : null,
                ProductSearchKey = d.TryGetValue("ProductSearchKey", out var psk) ? psk?.ToString() : null,
                CatalogProductName = d.TryGetValue("CatalogProductName", out var cpn) ? cpn?.ToString() : null,
            });
        }
        return result;
    }

    /// <summary>
    /// Returns SR rows (email body + metadata) for a given RFQ ID.
    /// Used by the RLI anchoring dry-run endpoint to re-run AI extraction
    /// with RLI context and compare against existing data.
    /// </summary>
    public async Task<List<SrEmailRow>> ReadSrEmailsByRfqIdAsync(string rfqId)
    {
        var siteId = await GetSiteIdAsync();
        var listId = await GetSupplierResponsesListIdAsync();
        var col    = await ResolveRfqIdColumnAsync(siteId, listId);

        var items = await GetGraph().Sites[siteId].Lists[listId].Items
            .GetAsync(req =>
            {
                req.QueryParameters.Expand = [$"fields($select={col},SupplierName,EmailBody,EmailFrom,EmailSubject,MessageId)"];
                req.QueryParameters.Filter = $"fields/{col} eq '{rfqId}'";
                req.QueryParameters.Top    = 50;
            });

        var result = new List<SrEmailRow>();
        foreach (var item in items?.Value ?? [])
        {
            if (item.Fields?.AdditionalData is null) continue;
            var d = item.Fields.AdditionalData;
            var id = d.TryGetValue(col, out var v0) ? v0?.ToString()
                   : d.TryGetValue("RFQ_x005F_ID", out var v1) ? v1?.ToString()
                   : d.TryGetValue("RFQ_ID", out var v2) ? v2?.ToString() : null;
            if (!string.Equals(id, rfqId, StringComparison.OrdinalIgnoreCase)) continue;
            result.Add(new SrEmailRow
            {
                SupplierName  = d.TryGetValue("SupplierName",  out var sn)  ? sn?.ToString()  : null,
                EmailBody     = d.TryGetValue("EmailBody",     out var eb)  ? eb?.ToString()  : null,
                EmailFrom     = d.TryGetValue("EmailFrom",     out var ef)  ? ef?.ToString()  : null,
                EmailSubject  = d.TryGetValue("EmailSubject",  out var es)  ? es?.ToString()  : null,
                MessageId     = d.TryGetValue("MessageId",     out var mid) ? mid?.ToString() : null,
            });
        }
        return result;
    }

    /// <summary>Creates one row per entry in <paramref name="items"/> in the RFQ Line Items list.</summary>
    public async Task CreateRfqLineItemsAsync(IEnumerable<RfqLineItemRequest> items)
    {
        var siteId = await GetSiteIdAsync();
        var listId = await GetRfqLineItemsListIdAsync();
        var col    = await ResolveRfqIdColumnAsync(siteId, listId);

        foreach (var req in items)
        {
            var data = new Dictionary<string, object?> { [col] = req.RfqId };
            if (req.Mspc            is not null) data["MSPC"]             = req.Mspc;
            if (req.Product         is not null) data["Product"]          = req.Product;
            if (req.Units           is not null) data["Units"]            = req.Units;
            if (req.SizeOfUnits     is not null) data["SizeOfUnits"]      = req.SizeOfUnits;
            if (req.SupplierEmails  is not null) data["SupplierEmails"]   = req.SupplierEmails;
            if (req.ProductCategory is not null) data["ProductCategory"]  = req.ProductCategory;
            if (req.ProductShape    is not null) data["ProductShape"]     = req.ProductShape;
            if (req.JobReference    is not null) data["JobReference"]     = req.JobReference;
            if (req.ProcessingSource is not null) data["ProcessingSource"] = req.ProcessingSource;

            var fieldsSummary = string.Join(", ", data.Select(kv => $"{kv.Key}={kv.Value}"));
            _log.LogInformation("[SP] CreateRfqLineItem posting: {Fields}", fieldsSummary);

            try
            {
                var created = await GetGraph().Sites[siteId].Lists[listId].Items
                    .PostAsync(new ListItem { Fields = new FieldValueSet { AdditionalData = data } });
                _log.LogInformation("[SP] CreateRfqLineItem OK: [{RfqId}] {Product} ?????' SpId={SpId}",
                    req.RfqId, req.Product, created?.Id);
            }
            catch (Microsoft.Graph.Models.ODataErrors.ODataError e)
            {
                _log.LogError("[SP] CreateRfqLineItem ODataError for [{RfqId}] {Product}: {Code}  -- {Msg}",
                    req.RfqId, req.Product, e.Error?.Code, e.Error?.Message);
                throw;
            }
        }
    }

    // Cached: list ID ?????' internal column name for RFQ_ID
    private readonly Dictionary<string, string> _rfqColByList = new();

    private async Task<string> ResolveRfqIdColumnAsync(string siteId, string listId)
    {
        if (_rfqColByList.TryGetValue(listId, out var cached)) return cached;

        var cols = await GetGraph().Sites[siteId].Lists[listId].Columns.GetAsync();
        var all  = cols?.Value ?? [];

        // Match on internal name (Name) or display name (DisplayName)  -- whichever has RFQ+ID.
        var col = all.FirstOrDefault(c =>
            c.Name?.Equals("RFQ_x005F_ID", StringComparison.OrdinalIgnoreCase) == true ||
            c.Name?.Equals("RFQ_ID",        StringComparison.OrdinalIgnoreCase) == true ||
            c.DisplayName?.Equals("RFQ_ID", StringComparison.OrdinalIgnoreCase) == true ||
            c.DisplayName?.Equals("RFQ ID", StringComparison.OrdinalIgnoreCase) == true);

        if (col is null)
        {
            var dump = string.Join("\n", all.Select(c => $"  Name={c.Name}  DisplayName={c.DisplayName}"));
            throw new InvalidOperationException(
                $"RFQ_ID column not found in list (id={listId}).\nAvailable columns:\n{dump}");
        }

        // Use the internal name (Name) for writes  -- DisplayName is just for matching.
        var name = col.Name
            ?? throw new InvalidOperationException(
                $"Column '{col.DisplayName}' has a null internal Name. Cannot write to it.");

        _rfqColByList[listId] = name;
        _log.LogInformation("[SP] RFQ_ID column resolved: Name={Name}  DisplayName={Display}",
            name, col.DisplayName);
        return name;
    }

    // ??"?????"??? Write: upsert one supplier email + all its extracted lines ??"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"???

    /// <summary>
    /// Main write entry point.  For each extracted email:
    ///   1. Upsert a <b>SupplierResponses</b> row (one per supplier email).
    ///   2. Upload the PDF attachment to that row.
    ///   3. Upsert a <b>SupplierLineItems</b> row for this product line.
    ///
    /// Uniqueness key for SupplierResponses: (RFQ_ID, EmailFrom).
    /// Uniqueness key for SupplierLineItems: (SupplierResponseId, ProductName).
    /// </summary>
    public async Task<SpWriteResult> WriteProductRowAsync(
        RfqExtraction  header,
        ProductLine    product,
        ExtractRequest emailMeta,
        string         source,
        string?        sourceFile,
        int            rowIndex,
        string?        messageId = null)
    {
        var result = new SpWriteResult { ProductName = product.ProductName };
        try
        {
            var siteId = await GetSiteIdAsync();

            // ?????? Resolve job reference ????????????????????????????????????????????????????????????????????????????????????????????????????????????????????????
            // Priority: AI-extracted ref from the PDF document > email subject/body regex ref.
            // PDF content is per-attachment, so when a supplier batches quotes for multiple RFQs
            // into one email (e.g. two PDFs each referencing different job IDs), Claude's per-PDF
            // extraction routes each attachment to the correct RFQ. The email subject regex acts
            // as a fallback for PDFs that don't embed our job number (supplier quote sheets, etc.)
            // — those yield a supplier quote number that fails IsValidRfqId, so we fall through.
            var regexRef = emailMeta.JobRefs.FirstOrDefault()?.Trim('[', ']');
            var aiRef    = header.JobReference?.Trim();

            // RFQ IDs — two valid lengths:
            //   6 chars — legacy alphanumeric (e.g. BX9EWM) or new initials+Crockford4 (e.g. AW0001)
            //   8 chars — "HQ" + 6 alphanumeric (e.g. HQ4RCAPR)
            // Reject AI-extracted values that don't match either format.
            static bool IsValidRfqId(string? s)
            {
                if (string.IsNullOrEmpty(s)) return false;
                if (!s.All(char.IsLetterOrDigit)) return false;
                if (s.Length != 6 && !(s.Length == 8 && s.StartsWith("HQ", StringComparison.OrdinalIgnoreCase))) return false;
                // Reject common English words that AI may infer from prose
                var upper = s.ToUpperInvariant();
                return upper is not ("PLEASE" or "THANKS" or "QUOTES" or "URGENT" or "ATTACH" or
                                     "REVIEW" or "ORDERS" or "FOLLOW" or "PRICES" or "UPDATE" or
                                     "NOTICE" or "KINDLY" or "HEREIN" or "RETURN" or "SUBMIT" or
                                     "CANNOT" or "SHOULD" or "ALWAYS" or "BEFORE" or "WITHIN");
            }

            var rawJobRef = ((IsValidRfqId(aiRef) ? aiRef : null)
                ?? regexRef
                ?? string.Empty).ToUpperInvariant();
            var jobRef = string.IsNullOrEmpty(rawJobRef) ? "000000" : rawJobRef;

            // Detect mismatch: PDF has a valid RFQ ID that differs from the email subject ref.
            // This happens when a supplier batches quotes for multiple RFQs, or sends a quote PDF
            // for the wrong job. Flag every SLI with a visible warning so the user can verify.
            string? jobRefMismatchWarning = null;
            if (IsValidRfqId(aiRef) && !string.IsNullOrEmpty(regexRef) &&
                !string.Equals(aiRef, regexRef, StringComparison.OrdinalIgnoreCase))
            {
                jobRefMismatchWarning =
                    $"⚠ Quote job ref [{aiRef?.ToUpperInvariant()}] differs from email ref [{regexRef.ToUpperInvariant()}] — verify this quote is for the correct RFQ.";
                _log.LogWarning(
                    "[SP] Job ref mismatch: PDF extracted [{AiRef}] but email subject has [{RegexRef}]. " +
                    "Filing under [{JobRef}]. Subject='{Subject}'  From='{From}'",
                    aiRef?.ToUpperInvariant(), regexRef.ToUpperInvariant(), jobRef,
                    emailMeta.EmailSubject, emailMeta.EmailFrom);
            }

            if (!string.IsNullOrEmpty(aiRef) && !IsValidRfqId(aiRef) && string.IsNullOrEmpty(regexRef))
            {
                _log.LogInformation(
                    "[SP] AI returned non-RFQ ID '{AiRef}' as JobReference  --  " +
                    "treating as orphan [000000]. Subject='{Subject}'",
                    aiRef, emailMeta.EmailSubject);
                if (string.IsNullOrWhiteSpace(header.QuoteReference))
                    header.QuoteReference = aiRef;
            }

            if (jobRef == "000000")
                _log.LogWarning(
                    "[SP] No job reference resolved  -- writing under [000000]. " +
                    "Subject='{Subject}'  From='{From}'  RawRef='{RawRef}'",
                    emailMeta.EmailSubject, emailMeta.EmailFrom,
                    string.IsNullOrEmpty(rawJobRef) ? "(none)" : rawJobRef);

            // ??"?????"??? Resolve supplier name ??"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"???
            // Fast path: ShrConvInRouter already resolved the supplier authoritatively (domain map
            // or historical SR lookup). Use it directly so the SLI lands under the same name as the
            // SupplierConversations row and any existing SLI rows from the supplier's original quote.
            if (!string.IsNullOrEmpty(emailMeta.ResolvedSupplierName))
            {
                var resolvedSup = emailMeta.ResolvedSupplierName;
                result.SupplierName = resolvedSup;

                var srListId2                = await GetSupplierResponsesListIdAsync();
                var (srId2, srNew2, _)       = await EnsureSupplierResponseAsync(
                    siteId, srListId2, jobRef, resolvedSup, header, emailMeta, source, sourceFile, messageId);
                result.SpItemId = srId2;
                result.Updated  = !srNew2;
                result.RfqId    = jobRef;

                if (result.SpItemId is not null &&
                    emailMeta.SourceType == "attachment" &&
                    !string.IsNullOrEmpty(emailMeta.FileName) &&
                    !string.IsNullOrEmpty(emailMeta.Base64Data))
                {
                    try { await UpsertItemAttachmentAsync(srId2, srListId2, emailMeta.FileName, Convert.FromBase64String(emailMeta.Base64Data)); }
                    catch (Exception ex) { _log.LogError(ex, "[SP] Attachment upload FAILED for SR {Id} ('{File}')", srId2, emailMeta.FileName); }
                }

                var sliListId2 = await GetSupplierLineItemsListIdAsync();
                if (messageId is not null && rowIndex == 0)
                    await PurgeNoMessageIdSliForSrAsync(siteId, sliListId2, srId2);
                await WriteSupplierLineItemAsync(
                    siteId, sliListId2, srId2, jobRef, resolvedSup, product, rowIndex,
                    sourceFile, emailMeta.EmailFrom, messageId, header.QuoteReference,
                    jobRefMismatchWarning);
                result.Success = true;
                return result;
            }

            // ── Supplier resolution: email domain is always the primary source ──────
            // The sender's email domain is the canonical identifier — it cannot be confused
            // by AI reading quoted original messages. AI-extracted names are a last resort.

            // Step 1: resolve from sender domain (DomainMap exact → token-substring).
            // Direct supplier replies: use emailMeta.EmailFrom as-is (it IS the supplier's address).
            // Forwarded emails (from internal staff): parse the original sender from the body,
            // because the internal Hackensack/Jay address is never in the supplier reference list.
            static bool IsInternalDomain(string? addr) =>
                addr is not null && (
                    addr.EndsWith("@mithrilmetals.com",      StringComparison.OrdinalIgnoreCase) ||
                    addr.EndsWith("@metalsupermarkets.com",  StringComparison.OrdinalIgnoreCase));

            string? supplier = null;
            string  supplierFromDomain = string.Empty;   // for WHOIS warning
            {
                var resolveEmail = IsInternalDomain(emailMeta.EmailFrom)
                    ? TryExtractForwardedSenderEmail(emailMeta.EmailBody ?? emailMeta.BodyContext)
                      ?? emailMeta.EmailFrom
                    : emailMeta.EmailFrom;

                if (!string.IsNullOrWhiteSpace(resolveEmail) && resolveEmail.Contains('@'))
                {
                    var senderDomain = resolveEmail[(resolveEmail.IndexOf('@') + 1)..].ToLowerInvariant();
                    supplierFromDomain = senderDomain;
                    string resolveSource = resolveEmail == emailMeta.EmailFrom
                        ? "direct sender domain"
                        : "forwarded-body sender domain";

                    if (_suppliers.DomainMap.TryGetValue(senderDomain, out var domainMatch))
                    {
                        supplier = domainMatch;
                        _log.LogDebug("[SP] Supplier '{Supplier}' resolved by {Source} (exact) '{Domain}'",
                            supplier, resolveSource, senderDomain);
                    }
                    else
                    {
                        var tld2 = senderDomain.Split('.')[0];
                        var substringMatch = _suppliers.ResolveByDomainSubstring(tld2);
                        if (substringMatch is not null)
                        {
                            supplier = substringMatch;
                            _log.LogDebug("[SP] Supplier '{Supplier}' resolved by {Source} (substring) '{Tld2}'",
                                supplier, resolveSource, tld2);
                        }
                        else
                        {
                            _log.LogDebug(
                                "[SP] Domain lookup failed '{Domain}' (tld2='{Tld2}') — trying AI name fallback. " +
                                "Known suppliers: [{Known}]",
                                senderDomain, tld2,
                                string.Join(", ", _suppliers.CachedNames));
                        }
                    }
                }
            }

            // Step 2: AI-extracted name fallback — only when domain resolution failed.
            // Discard names that are our own company (AI reads quoted original messages
            // and may extract "Mithril Metals" or "Metal Supermarkets" as the supplier).
            if (supplier is null)
            {
                var rawAiSupplier = header.SupplierName;
                static bool IsOurCompanyName(string? s) =>
                    s is not null && (
                        s.Contains("mithril", StringComparison.OrdinalIgnoreCase) ||
                        s.Contains("metal supermarket", StringComparison.OrdinalIgnoreCase) ||
                        s.Contains("hackensack", StringComparison.OrdinalIgnoreCase));
                if (IsOurCompanyName(rawAiSupplier))
                    rawAiSupplier = null;

                if (!string.IsNullOrWhiteSpace(rawAiSupplier))
                {
                    supplier = _suppliers.ResolveSupplierName(rawAiSupplier);
                    if (supplier is not null)
                        _log.LogInformation(
                            "[SP] Supplier '{Supplier}' resolved by AI name '{Raw}' (domain lookup failed). " +
                            "Subject='{Subject}'",
                            supplier, rawAiSupplier, emailMeta.EmailSubject);
                }
            }

            if (supplier is null)
            {
                result.SupplierUnknown = true;
                supplier = "Unknown";
                jobRef   = "WHOIS";
                _log.LogWarning(
                    "[SP] Supplier not in reference list (domain='{Domain}', AI='{AiName}') — writing under [WHOIS]. " +
                    "Subject='{Subject}'  From='{From}'",
                    supplierFromDomain, header.SupplierName, emailMeta.EmailSubject, emailMeta.EmailFrom);
            }

            result.SupplierName = supplier;

            // ??"?????"??? Upsert SupplierResponses ??"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"???
            var srListId                   = await GetSupplierResponsesListIdAsync();
            var (srId, srNew, contactEmail) = await EnsureSupplierResponseAsync(
                siteId, srListId, jobRef, supplier, header, emailMeta, source, sourceFile, messageId);
            result.SpItemId = srId;
            result.Updated  = !srNew;   // true = existing row updated; false = new insert
            result.RfqId    = jobRef;   // resolved RFQ ID  -- may differ from req.JobRefs when email subject has no bracket ref

            // Upload the source attachment as a SharePoint list item attachment.
            if (result.SpItemId is not null &&
                emailMeta.SourceType == "attachment" &&
                !string.IsNullOrEmpty(emailMeta.FileName) &&
                !string.IsNullOrEmpty(emailMeta.Base64Data))
            {
                try
                {
                    var bytes = Convert.FromBase64String(emailMeta.Base64Data);
                    await UpsertItemAttachmentAsync(srId, srListId, emailMeta.FileName, bytes);
                }
                catch (Exception ex)
                {
                    _log.LogError(ex, "[SP] Attachment upload FAILED for SR {Id} ('{File}')  -- quote PDF will be missing from SharePoint", srId, emailMeta.FileName);
                }
            }

            // ??"?????"??? Upsert SupplierLineItems ??"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"???
            var sliListId = await GetSupplierLineItemsListIdAsync();

            // On the first product of a MessageId-bearing reprocess, purge any existing SLI
            // rows for this SR that have no MessageId.  These are stale rows from an earlier
            // extraction whose product names may differ enough to fail the fuzzy match, causing
            // the new extraction to insert duplicates instead of replacing them.
            // We only reach here after the AI has already returned
            // is safe  -- the new rows are about to be written immediately after.
            if (messageId is not null && rowIndex == 0)
                await PurgeNoMessageIdSliForSrAsync(siteId, sliListId, srId);

            await WriteSupplierLineItemAsync(
                siteId, sliListId, srId, jobRef, supplier, product, rowIndex,
                sourceFile, emailMeta.EmailFrom, messageId, header.QuoteReference,
                jobRefMismatchWarning);

            result.Success = true;
        }
        catch (Microsoft.Graph.Models.ODataErrors.ODataError odataEx)
        {
            result.Success = false;
            result.Error   = odataEx.Message;
            _log.LogError("[SP] ODataError: code={Code} msg={Msg}", odataEx.Error?.Code, odataEx.Error?.Message);
        }
        catch (Exception ex)
        {
            result.Success = false;
            result.Error   = ex.Message;
            _log.LogError(ex, "[SP] Failed to upsert product '{Name}'", product.ProductName);
        }
        return result;
    }

    // ??"?????"??? Upsert SupplierResponses (private) ??"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"???

    private async Task<(string Id, bool IsNew, string? ContactEmail)> EnsureSupplierResponseAsync(
        string siteId, string listId,
        string jobRef, string supplier,
        RfqExtraction header, ExtractRequest emailMeta,
        string source, string? sourceFile,
        string? messageId = null)
    {
        var existingId = await FindExistingSupplierResponseAsync(siteId, listId, jobRef, supplier, header.QuoteReference);
        bool isNew = existingId is null;

        var emailBodyTrunc = (emailMeta.EmailBody ?? emailMeta.BodyContext) is string body
            ? body[..Math.Min(body.Length,
                  int.TryParse(_config["SharePoint:MaxEmailBodyChars"], out var mebc) ? mebc : 10_000)]
            : null;

        // For forwarded emails (from internal staff) extract the original supplier sender address
        // embedded by Outlook in the forwarded body ("From: name <email>"). Store it as ContactEmail
        // so the conversation feature can reply to the real supplier rather than the internal forwarder.
        var forwardedSender = TryExtractForwardedSenderEmail(emailMeta.EmailBody ?? emailMeta.BodyContext);
        var contactEmail = forwardedSender is not null &&
                           !string.Equals(forwardedSender, emailMeta.EmailFrom, StringComparison.OrdinalIgnoreCase)
            ? forwardedSender
            : null;

        bool blanketRegret = HasRegretPhrase(emailMeta.EmailBody) || HasRegretPhrase(emailMeta.BodyContext);

        var title = $"[{jobRef}] {supplier} {(emailMeta.ReceivedAt is not null ? DateTime.Parse(emailMeta.ReceivedAt).ToString("yyyy-MM-dd") : "unknown")}";
        title = title[..Math.Min(title.Length, 255)];

        var fieldData = new Dictionary<string, object?>
        {
            ["Title"]                = title,
            ["RFQ_ID"]               = string.IsNullOrEmpty(jobRef) ? null : jobRef,
            ["SupplierName"]         = supplier,
            ["EmailFrom"]            = emailMeta.EmailFrom,
            ["ContactEmail"]         = contactEmail,
            ["ReceivedAt"]           = emailMeta.ReceivedAt,
            ["EmailSubject"]         = emailMeta.EmailSubject,
            ["EmailBody"]            = emailBodyTrunc,
            ["ProcessedAt"]          = DateTime.UtcNow.ToString("o"),
            ["ProcessingSource"]     = source,
            ["SourceFile"]           = sourceFile,
            ["MessageId"]            = messageId,
            ["QuoteReference"]       = header.QuoteReference,
            // DateOfQuote / EstimatedDeliveryDate intentionally omitted  --
            // dates come from the RFQ Reference record, not from extraction.
            ["FreightTerms"]         = header.FreightTerms,
            ["IsRegret"]             = blanketRegret,
        };

        // Build an NDJSON log entry for this AI response
        // overwritten, so every extraction is preserved for auditing and smart merging.
        var logEntry = System.Text.Json.JsonSerializer.Serialize(new
        {
            Ts  = DateTime.UtcNow.ToString("o"),
            Src = source,
            Ext = header,
        }, new System.Text.Json.JsonSerializerOptions
        {
            DefaultIgnoreCondition = System.Text.Json.Serialization.JsonIgnoreCondition.WhenWritingNull,
        });

        if (!isNew)
        {
            // Fetch the precious AI-extracted fields that already exist on this row.
            // We only overwrite them when the current value is blank  -- a good extraction
            // from an earlier pass (e.g. the attachment run) should never be clobbered by
            // a weaker body-only re-run that returns nulls or less detail.
            // ProcessingSource/SourceFile are also protected: attachment beats body.
            var precious = new[] { "QuoteReference", "DateOfQuote", "EstimatedDeliveryDate",
                                   "FreightTerms", "ProcessingSource", "SourceFile",
                                   "ContactEmail", "ClaudeResponseLog" };
            var currentItem = await GetGraph().Sites[siteId].Lists[listId].Items[existingId!]
                .GetAsync(r => r.QueryParameters.Expand =
                    [$"fields($select={string.Join(",", precious)})"]);
            var cur = currentItem?.Fields?.AdditionalData ?? new Dictionary<string, object?>();

            var update = new Dictionary<string, object?>(fieldData);
            foreach (var key in precious.Where(k => k != "ClaudeResponseLog"))
            {
                if (!update.TryGetValue(key, out var newVal)) continue;
                cur.TryGetValue(key, out var curVal);
                var curStr = curVal is JsonElement cje ? cje.GetString() : curVal?.ToString();
                var newStr = newVal is JsonElement nje ? nje.GetString() : newVal?.ToString();
                // Keep existing value when it is populated and the new value adds nothing
                if (!string.IsNullOrWhiteSpace(curStr) && string.IsNullOrWhiteSpace(newStr))
                    update.Remove(key);
                // Never downgrade attachment ?????' body
                if (key == "ProcessingSource" &&
                    string.Equals(curStr, "attachment", StringComparison.OrdinalIgnoreCase) &&
                    !string.Equals(newStr, "attachment", StringComparison.OrdinalIgnoreCase))
                    update.Remove(key);
            }

            // Append to the log  -- keep the last 50 entries so the field stays within SP limits
            cur.TryGetValue("ClaudeResponseLog", out var existingLog);
            var existingLogStr = existingLog is JsonElement lje ? lje.GetString() : existingLog?.ToString();
            var logLines = string.IsNullOrWhiteSpace(existingLogStr)
                ? []
                : existingLogStr!.Split('\n', StringSplitOptions.RemoveEmptyEntries).ToList();
            if (logLines.Count >= 50)
                logLines = logLines.TakeLast(49).ToList();
            logLines.Add(logEntry);
            update["ClaudeResponseLog"] = string.Join("\n", logLines);

            // Strip any remaining null-valued keys  -- no point patching fields to null
            foreach (var key in update.Keys.Where(k => update[k] is null).ToList())
                update.Remove(key);

            await PatchFieldsAsync(siteId, listId, existingId!, update);
            InvalidateSrCache();
            _log.LogInformation("[SP] Updated SupplierResponse {Id} for [{JobRef}] {Supplier}", existingId, jobRef, supplier);
            return (existingId!, false, contactEmail);
        }
        else
        {
            fieldData["ClaudeResponseLog"] = logEntry;
            var item = await PostListItemAsync(siteId, listId, fieldData);
            var newId = item!.Id!;
            InvalidateSrCache();
            _log.LogInformation("[SP] Created SupplierResponse {Id} for [{JobRef}] {Supplier}", newId, jobRef, supplier);
            return (newId, true, contactEmail);
        }
    }

    // Optional columns that may not exist on older list schemas.
    // When a PATCH or POST fails because one of these is unrecognised, it is
    // silently dropped and the request retried rather than failing the whole write.
    private static readonly string[] _optionalColumns = ["ClaudeResponseLog", "ContactEmail"];

    /// <summary>
    /// PATCH a SP list item's fields, retrying without any optional columns
    /// that don't exist in this list schema yet.
    /// </summary>
    private async Task PatchFieldsAsync(string siteId, string listId, string itemId,
        Dictionary<string, object?> data)
    {
        try
        {
            await GetGraph().Sites[siteId].Lists[listId].Items[itemId].Fields
                .PatchAsync(new FieldValueSet { AdditionalData = data });
        }
        catch (Microsoft.Graph.Models.ODataErrors.ODataError e)
            when (_optionalColumns.Any(c => e.Error?.Message?.Contains(c) == true))
        {
            var missing = _optionalColumns.First(c => e.Error?.Message?.Contains(c) == true);
            _log.LogDebug("[SP] Optional field '{Field}' absent  -- retrying PATCH without it", missing);
            data.Remove(missing);
            await GetGraph().Sites[siteId].Lists[listId].Items[itemId].Fields
                .PatchAsync(new FieldValueSet { AdditionalData = data });
        }
    }

    /// <summary>
    /// POST a new SP list item, retrying without any optional columns
    /// that don't exist in this list schema yet.
    /// </summary>
    private async Task<ListItem?> PostListItemAsync(string siteId, string listId,
        Dictionary<string, object?> data)
    {
        try
        {
            return await GetGraph().Sites[siteId].Lists[listId].Items
                .PostAsync(new ListItem { Fields = new FieldValueSet { AdditionalData = data } });
        }
        catch (Microsoft.Graph.Models.ODataErrors.ODataError e)
            when (_optionalColumns.Any(c => e.Error?.Message?.Contains(c) == true))
        {
            var missing = _optionalColumns.First(c => e.Error?.Message?.Contains(c) == true);
            _log.LogDebug("[SP] Optional field '{Field}' absent  -- retrying POST without it", missing);
            data.Remove(missing);
            return await GetGraph().Sites[siteId].Lists[listId].Items
                .PostAsync(new ListItem { Fields = new FieldValueSet { AdditionalData = data } });
        }
    }

    private async Task<string?> FindExistingSupplierResponseAsync(
        string siteId, string listId, string jobRef, string supplierName,
        string? quoteReference = null)
    {
        if (string.IsNullOrEmpty(jobRef)) return null;

        // Fetch all SR rows and filter client-side, following nextLink to cover lists > 2000 rows.
        // OData filter on non-indexed columns with HonorNonIndexedQueriesWarningMayFailRandomly
        // can silently return empty results, causing a new SR to be inserted instead of updating
        // the existing one  -- producing duplicate supplier response rows.
        // $select includes both RFQ_ID and RFQ_x005F_ID because SharePoint may return the
        // underscore column under either internal name depending on how the list was created.
        var page = await GetGraph().Sites[siteId].Lists[listId].Items
            .GetAsync(r =>
            {
                r.QueryParameters.Expand = ["fields($select=id,RFQ_ID,RFQ_x005F_ID,SupplierName,QuoteReference)"];
                r.QueryParameters.Top    = 2000;
            });

        string? nameMatch    = null;
        string? quoteRefMatch = null;

        while (page is not null)
        {
            foreach (var i in page.Value ?? [])
            {
                var d = i.Fields?.AdditionalData;
                if (d is null || i.Id is null) continue;
                var itemJobRef = (d.TryGetValue("RFQ_ID",       out var jv)  ? jv?.ToString()  : null)
                              ?? (d.TryGetValue("RFQ_x005F_ID", out var jv2) ? jv2?.ToString() : null);
                if (!string.Equals(itemJobRef, jobRef, StringComparison.OrdinalIgnoreCase)) continue;

                var itemSupplier  = d.TryGetValue("SupplierName",   out var sv) ? sv?.ToString() : null;
                var itemQuoteRef  = d.TryGetValue("QuoteReference",  out var qv) ? qv?.ToString() : null;

                // Strong match: same job ref + same quote reference (supplier-assigned number).
                // Matches regardless of email address or name variation.
                if (!string.IsNullOrEmpty(quoteReference) && !string.IsNullOrEmpty(itemQuoteRef)
                    && string.Equals(quoteReference, itemQuoteRef, StringComparison.OrdinalIgnoreCase))
                {
                    quoteRefMatch = i.Id;
                    break;
                }

                // Fallback: same job ref + same supplier name.
                if (nameMatch is null
                    && !string.IsNullOrEmpty(supplierName)
                    && string.Equals(itemSupplier, supplierName, StringComparison.OrdinalIgnoreCase))
                    nameMatch = i.Id;
            }

            if (quoteRefMatch is not null) break;
            if (page.OdataNextLink is null) break;
            var next = new Microsoft.Graph.Sites.Item.Lists.Item.Items.ItemsRequestBuilder(
                page.OdataNextLink, GetGraph().RequestAdapter);
            page = await next.GetAsync();
        }

        return quoteRefMatch ?? nameMatch;
    }

    // ??"?????"??? Write SupplierLineItems (private) ??"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"???

    private async Task WriteSupplierLineItemAsync(
        string siteId, string listId,
        string supplierResponseId, string jobRef, string supplier,
        ProductLine product, int rowIndex,
        string? sourceFile, string? emailFrom, string? messageId = null,
        string? quoteReference = null, string? jobRefMismatchWarning = null)
    {
        var prodName   = product.ProductName ?? $"Product {rowIndex + 1}";
        var prodTokens = ProductTokens(prodName);

        // If the AI resolved an MSPC directly from the RLI requested-items list, use it.
        // Otherwise fall back to the fuzzy catalog matcher on the supplier's product name.
        string? productSearchKey;
        string? catalogProductName;
        if (!string.IsNullOrEmpty(product.ProductSearchKey))
        {
            productSearchKey   = product.ProductSearchKey;
            var byKey          = _catalog.FindBySearchKey(product.ProductSearchKey);
            catalogProductName = byKey?.Name;
            _log.LogInformation(
                "[SP] RLI-anchored: [{RfqId}] {Supplier} '{Name}' ?????' MSPC={Key} catalog='{Catalog}'",
                jobRef, supplier, prodName, productSearchKey, catalogProductName ?? "(not in catalog)");
        }
        else
        {
            var catalogMatch   = _catalog.ResolveProduct(prodName);
            productSearchKey   = catalogMatch?.SearchKey;
            catalogProductName = catalogMatch?.Name;

            if (productSearchKey is null)
                _log.LogWarning(
                    "[SP] No catalog match for [{RfqId}] {Supplier} '{Name}'  -- ProductSearchKey will be null",
                    jobRef, supplier, prodName);
        }

        var title = $"[{jobRef}] {supplier} - {prodName}";
        title = title[..Math.Min(title.Length, 255)];

        var fieldData = new Dictionary<string, object?>
        {
            ["Title"]                    = title,
            ["SupplierResponseId"]       = supplierResponseId,
            ["RFQ_ID"]                   = string.IsNullOrEmpty(jobRef) ? null : jobRef,
            ["SupplierName"]             = supplier,
            ["QuoteReference"]           = quoteReference,
            ["ProductName"]              = prodName,
            ["SupplierProductName"]      = prodName,
            ["CatalogProductName"]       = catalogProductName,
            ["ProductSearchKey"]         = productSearchKey,
            ["SourceFile"]               = sourceFile,
            ["EmailFrom"]                = emailFrom,
            ["MessageId"]                = messageId,
            ["UnitsRequested"]           = product.UnitsRequested,
            ["UnitsQuoted"]              = product.UnitsQuoted,
            ["LengthPerUnit"]            = product.LengthPerUnit,
            ["LengthUnit"]               = product.LengthUnit,
            ["WeightPerUnit"]            = product.WeightPerUnit,
            ["WeightUnit"]               = product.WeightUnit,
            ["PricePerPound"]            = product.PricePerPound,
            ["PricePerFoot"]             = product.PricePerFoot,
            ["PricePerPiece"]            = product.PricePerPiece,
            ["TotalPrice"]               = product.TotalPrice ?? ComputeTotalPrice(product),
            ["LeadTimeText"]             = product.LeadTimeText,
            ["Certifications"]           = product.Certifications,
            ["SupplierProductComments"]  = string.IsNullOrEmpty(jobRefMismatchWarning)
                                               ? product.SupplierProductComments
                                               : string.IsNullOrEmpty(product.SupplierProductComments)
                                                   ? jobRefMismatchWarning
                                                   : $"{jobRefMismatchWarning} {product.SupplierProductComments}",
            // A quoted price means the supplier IS supplying — don't mark regret even if
            // cross-product comments contain regret language (e.g. "regrets on the pipe").
            ["IsRegret"]                 = !HasPrice(product) && HasRegretPhrase(product.SupplierProductComments),
            ["IsSubstitute"]             = product.IsSubstitute,
        };

        var existing = await FindExistingSupplierLineItemAsync(
            siteId, listId, supplierResponseId, jobRef, supplier, prodName, prodTokens, quoteReference, productSearchKey);

        if (existing is not null)
        {
            var update = new Dictionary<string, object?>(fieldData);
            update.Remove("ProductName");         // preserve canonical name from first write
            update.Remove("SupplierProductName"); // preserve original supplier name from first write

            // Preserve SupplierProductComments: AI commentary is cumulative.
            // Only overwrite when the existing row has no comment and the new run does.
            // If both have values and they differ, keep the existing one (first good extraction wins).
            var curComments = existing.Fields?.AdditionalData is { } d &&
                              d.TryGetValue("SupplierProductComments", out var cv)
                              ? (cv is JsonElement je ? je.GetString() : cv?.ToString())
                              : null;
            var newComments = product.SupplierProductComments;
            if (!string.IsNullOrWhiteSpace(curComments))
                update.Remove("SupplierProductComments"); // keep existing
            else if (string.IsNullOrWhiteSpace(newComments))
                update.Remove("SupplierProductComments"); // nothing to write

            // Pricing fields: when the new extraction has ANY pricing, include all price fields in
            // the PATCH even if null — this clears stale values when a supplier changes their format
            // (e.g. previously quoted $/lb, now quoting total-only).
            var pricingKeys = new HashSet<string>
                { "PricePerPound", "PricePerFoot", "PricePerPiece", "TotalPrice",
                  "UnitsQuoted", "LengthPerUnit", "LengthUnit", "WeightPerUnit", "WeightUnit" };
            bool hasNewPricing = pricingKeys.Any(k => update.TryGetValue(k, out var v) && v is not null);

            // Strip null values from non-pricing fields (don't null out other populated fields).
            // Preserve null pricing fields when hasNewPricing so stale price columns get cleared.
            foreach (var key in update.Keys.Where(k => update[k] is null && (!hasNewPricing || !pricingKeys.Contains(k))).ToList())
                update.Remove(key);

            await GetGraph().Sites[siteId].Lists[listId].Items[existing.Id!].Fields
                .PatchAsync(new FieldValueSet { AdditionalData = update });
            _log.LogInformation("[SP] Updated SupplierLineItem {Id} for '{Name}'", existing.Id, prodName);
        }
        else
        {
            var item = await GetGraph().Sites[siteId].Lists[listId].Items
                .PostAsync(new ListItem { Fields = new FieldValueSet { AdditionalData = fieldData } });
            _log.LogInformation("[SP] Created SupplierLineItem {Id} for '{Name}'", item!.Id, prodName);
        }
    }

    private async Task<ListItem?> FindExistingSupplierLineItemAsync(
        string siteId, string listId,
        string supplierResponseId, string jobRef, string supplierName,
        string productName, HashSet<string> productTokens,
        string? quoteReference = null, string? productSearchKey = null)
    {
        // Fetch all SLI rows and filter in memory, following nextLink so lists > 2000 rows
        // don't silently miss existing items and create duplicates.
        // OData filter on the non-indexed SupplierResponseId field is avoided here because
        // "HonorNonIndexedQueriesWarningMayFailRandomly" can silently return empty results.
        // SupplierProductComments is included so the caller can apply fill-blanks without
        // a second round-trip.
        var page = await GetGraph().Sites[siteId].Lists[listId].Items
            .GetAsync(r =>
            {
                r.QueryParameters.Expand = ["fields($select=id,SupplierResponseId,RFQ_ID,SupplierName,ProductName,ProductSearchKey,QuoteReference,SupplierProductComments)"];
                r.QueryParameters.Top    = 2000;
            });

        // Cross-email dedup: prefer a match within the same SR (same email), but fall back to a
        // matching row from a different SR for the same RFQ+supplier so repeated quotes from the
        // same supplier update the existing row rather than creating duplicates.
        bool canCrossSr = !string.IsNullOrEmpty(jobRef) && jobRef != "000000";
        ListItem? srMatch      = null;
        ListItem? crossSrMatch = null;

        while (page is not null)
        {
            foreach (var item in page.Value ?? [])
            {
                var d = item.Fields?.AdditionalData;
                if (d is null) continue;

                var srId  = d.TryGetValue("SupplierResponseId", out var sid) ? sid?.ToString() : null;
                bool sameSr = string.Equals(srId, supplierResponseId, StringComparison.OrdinalIgnoreCase);

                if (sameSr)
                {
                    if (srMatch is null && SliProductMatches(d, productName, productTokens, quoteReference, productSearchKey))
                        srMatch = item;
                }
                else if (canCrossSr && crossSrMatch is null)
                {
                    var spRfqId    = d.TryGetValue("RFQ_ID",       out var ri) ? ri?.ToString() : null;
                    var spSupplier = d.TryGetValue("SupplierName",  out var sn) ? sn?.ToString() : null;
                    if (string.Equals(spRfqId, jobRef, StringComparison.OrdinalIgnoreCase)
                        && string.Equals(spSupplier, supplierName, StringComparison.OrdinalIgnoreCase)
                        && SliProductMatches(d, productName, productTokens, quoteReference, productSearchKey))
                    {
                        crossSrMatch = item;
                    }
                }
            }

            if (srMatch is not null) return srMatch;

            if (page.OdataNextLink is null) break;
            var next = new Microsoft.Graph.Sites.Item.Lists.Item.Items.ItemsRequestBuilder(
                page.OdataNextLink, GetGraph().RequestAdapter);
            page = await next.GetAsync();
        }

        return srMatch ?? crossSrMatch;
    }

    private bool SliProductMatches(
        IDictionary<string, object?> d,
        string productName, HashSet<string> productTokens,
        string? quoteReference, string? productSearchKey)
    {
        // Strong match: same quote reference + same catalog product key.
        // If the supplier attached the same quote PDF twice (or it was forwarded
        // to multiple mailboxes), this reliably identifies the same line without
        // relying on raw product name extraction being identical.
        var spQuoteRef  = d.TryGetValue("QuoteReference",  out var qr) ? qr?.ToString() : null;
        var spSearchKey = d.TryGetValue("ProductSearchKey", out var sk) ? sk?.ToString() : null;
        if (!string.IsNullOrEmpty(quoteReference) && !string.IsNullOrEmpty(spQuoteRef)
            && string.Equals(quoteReference, spQuoteRef, StringComparison.OrdinalIgnoreCase))
        {
            // Same quote  -- match by catalog key + numeric tokens when available, else by name.
            // We must check numeric tokens even when the catalog key matches because two line items
            // on the same quote can share an MSPC but differ in cut size
            // (e.g. 4'×8' and 4'×10' cold-rolled sheet are both CSHCQ/048).
            // A pure catalog-key comparison would incorrectly collapse them into one row.
            var spProduct2 = d.TryGetValue("ProductName", out var p2) ? p2?.ToString() : null;
            if (!string.IsNullOrEmpty(productSearchKey) && !string.IsNullOrEmpty(spSearchKey))
            {
                if (!string.Equals(productSearchKey, spSearchKey, StringComparison.OrdinalIgnoreCase))
                    return false; // different catalog products
                // Same MSPC: still verify dimensions match so distinct cut sizes are kept separate.
                var spTok3 = ProductTokens(spProduct2 ?? string.Empty);
                return NumericTokensCompatible(productTokens, spTok3)
                    && ProductJaccard(spTok3, productTokens) >= 0.4;
            }
            if (NormalizeMatch(spProduct2, productName)) return true;
            var spTok2 = ProductTokens(spProduct2 ?? string.Empty);
            return NumericTokensCompatible(productTokens, spTok2)
                && ProductJaccard(spTok2, productTokens) >= 0.4;
        }

        // Standard fuzzy match by product name.
        var spProduct = d.TryGetValue("ProductName", out var p) ? p?.ToString() : null;
        if (NormalizeMatch(spProduct, productName)) return true;
        var spTokens = ProductTokens(spProduct ?? string.Empty);
        return NumericTokensCompatible(productTokens, spTokens)
            && ProductJaccard(spTokens, productTokens) >= 0.5;
    }

    // ??"?????"??? Purge stale no-MessageId SLI rows ??"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"???

    /// <summary>
    /// Deletes all SupplierLineItems rows for the given SR that have no MessageId.
    /// Called before inserting a new batch when the incoming email carries a MessageId,
    /// so stale rows from an earlier extraction (which lacked a MessageId and may have
    /// different product names) don't survive as duplicates.
    /// </summary>
    private async Task PurgeNoMessageIdSliForSrAsync(string siteId, string listId, string srId)
    {
        var page = await GetGraph().Sites[siteId].Lists[listId].Items
            .GetAsync(r =>
            {
                r.QueryParameters.Expand = ["fields($select=id,SupplierResponseId,MessageId)"];
                r.QueryParameters.Top    = 2000;
            });

        var toDelete = new List<string>();
        while (page is not null)
        {
            foreach (var item in page.Value ?? [])
            {
                var d = item.Fields?.AdditionalData;
                if (d is null || item.Id is null) continue;
                var itemSrId = d.TryGetValue("SupplierResponseId", out var s) ? s?.ToString() : null;
                if (!string.Equals(itemSrId, srId, StringComparison.OrdinalIgnoreCase)) continue;
                var msgId = d.TryGetValue("MessageId", out var m) ? m?.ToString() : null;
                if (string.IsNullOrEmpty(msgId))
                    toDelete.Add(item.Id);
            }
            if (page.OdataNextLink is null) break;
            var next = new Microsoft.Graph.Sites.Item.Lists.Item.Items.ItemsRequestBuilder(
                page.OdataNextLink, GetGraph().RequestAdapter);
            page = await next.GetAsync();
        }

        foreach (var id in toDelete)
        {
            await GetGraph().Sites[siteId].Lists[listId].Items[id].DeleteAsync();
            _log.LogInformation("[SP] Purged stale no-MessageId SLI {Id} for SR {SrId}", id, srId);
        }
    }

    // ??"?????"??? TotalPrice fallback calculation ??"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"???

    /// <summary>
    /// Mirrors the AI Step-8 forward calculation as a server-side fallback.
    /// Called when the AI returns a null totalPrice despite having valid unit prices and quantities.
    /// </summary>
    private static double? ComputeTotalPrice(ProductLine p)
    {
        var qty = (double?)(p.UnitsQuoted ?? p.UnitsRequested);

        // a. piece price ?? --  qty
        if (p.PricePerPiece.HasValue && qty.HasValue)
            return p.PricePerPiece.Value * qty.Value;
        // b. foot price ?? --  qty ?? --  length
        if (p.PricePerFoot.HasValue && qty.HasValue && p.LengthPerUnit.HasValue)
        {
            var ft = (p.LengthUnit ?? "ft").Trim().ToLowerInvariant() switch
            {
                "in" => p.LengthPerUnit.Value / 12.0,
                "m"  => p.LengthPerUnit.Value * 3.28084,
                "mm" => p.LengthPerUnit.Value / 304.8,
                "cm" => p.LengthPerUnit.Value / 30.48,
                _    => p.LengthPerUnit.Value
            };
            return p.PricePerFoot.Value * qty.Value * ft;
        }
        // c. pound price ?? --  qty ?? --  weight
        if (p.PricePerPound.HasValue && qty.HasValue && p.WeightPerUnit.HasValue)
        {
            var lb = (p.WeightUnit ?? "lb").Trim().ToLowerInvariant() switch
            {
                "kg" => p.WeightPerUnit.Value * 2.20462,
                "oz" => p.WeightPerUnit.Value / 16.0,
                "g"  => p.WeightPerUnit.Value / 453.592,
                _    => p.WeightPerUnit.Value
            };
            return p.PricePerPound.Value * qty.Value * lb;
        }
        return null;
    }

    // ??"?????"??? Regret detection ??"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"???

    private static bool HasRegretPhrase(string? text) =>
        text is not null &&
        _regretPhrases.Any(p => text.Contains(p, StringComparison.OrdinalIgnoreCase));

    private static bool HasPrice(ProductLine p) =>
        p.TotalPrice.HasValue || p.PricePerPound.HasValue ||
        p.PricePerFoot.HasValue || p.PricePerPiece.HasValue;

    // ??"?????"??? Forwarded-email original-sender extraction ??"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"???

    // Matches the "From:" line inside a forwarded message block.
    // Handles both plain-text and HTML-stripped formats, e.g.:
    //   From: Robert Smith <robert@pennstainless.com>
    //   From: robert@pennstainless.com
    //   From: "Penn Stainless" <quotes@pennstainless.com>
    // Anchored so it only fires after a line-start (not mid-sentence).
    private static readonly Regex _forwardedFromRegex = new(
        @"(?:^|[\r\n])\s*From:\s*(?:[^\r\n<@]*<)?\s*([a-zA-Z0-9.+\-]+@[a-zA-Z0-9.\-]+\.[a-zA-Z]{2,})",
        RegexOptions.Compiled | RegexOptions.IgnoreCase);

    /// <summary>
    /// Scans <paramref name="body"/> for the original sender's email address embedded
    /// by Outlook when it forwards a message (the "From: ..." line in the forwarded block).
    /// Returns the first email address found, or <see langword="null"/> if none is present.
    /// </summary>
    private static string? TryExtractForwardedSenderEmail(string? body)
    {
        if (string.IsNullOrWhiteSpace(body)) return null;
        var m = _forwardedFromRegex.Match(body);
        return m.Success ? m.Groups[1].Value.Trim() : null;
    }

    // ??"?????"??? OData helpers ??"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"???

    private static string EscapeOdata(string s) => s.Replace("'", "''");

    // ??"?????"??? Product tokenisation ??"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"???

    private static readonly Regex _normaliseRegex  = new(@"[\s\W]+", RegexOptions.Compiled);
    private static bool NormalizeMatch(string? a, string? b)
    {
        if (a is null && b is null) return true;
        if (a is null || b is null) return false;
        // Convert fractions before stripping so "3/8" → "3f8" and "1/2" → "1f2" stay distinct.
        // Without this, stripping "/" makes "3/8 x 12" and "3/8 x 1/2" both normalise to "38x12".
        static string N(string s) => _normaliseRegex.Replace(
            _dimFraction.Replace(s.Trim().ToLowerInvariant(), "$1f$2"), "");
        return N(a) == N(b);
    }

    private static readonly Regex _dimFraction  = new(@"(\d+)/(\d+)",                               RegexOptions.Compiled);
    private static readonly Regex _dimDecimal   = new(@"(\d+)\.(\d+)",                              RegexOptions.Compiled);
    private static readonly Regex _dimSeparator = new(@"(\d[a-z0-9]*)[""']?\s*[xX?? -- ]\s*[""']?(\d[a-z0-9]*)", RegexOptions.Compiled);
    private static readonly Regex _dimSplit     = new(@"[^a-z0-9]+",                                RegexOptions.Compiled);
    private static readonly Regex _orLength     = new(@"\bor\s+\d+[a-z""']*\b", RegexOptions.Compiled | RegexOptions.IgnoreCase);
    // Mixed number: "2-1/2" (whole-numerator/denominator). Must be handled before _dimFraction
    // so the compound value 2.5 becomes a single token rather than splitting into "2" and "1/2",
    // which would collapse "1/4 x 2 x 144" and "1/4 x 2-1/2 x 144" to identical numeric sets.
    private static readonly Regex _mixedNumber  = new(@"(\d+)-(\d+)/(\d+)",                         RegexOptions.Compiled);

    private static readonly Regex _trailingZeroInch = new(@"(\d+)\s*'\s*0\s*""", RegexOptions.Compiled);

    private static string PreprocessProduct(string s)
    {
        s = s.ToLowerInvariant();
        // "25' 0"" and "25'" are the same length  -- strip the zero-inch component
        // before any other processing so it doesn't produce a spurious numeric token.
        s = _trailingZeroInch.Replace(s, "$1'");
        s = _orLength.Replace(s, "");
        s = Regex.Replace(s, @"\brandom\s+lengths?\b|\bmill\s+lengths?\b|\bfull\s+lengths?\b|\blengths?\b", "");
        // Convert mixed numbers (e.g. "2-1/2") to their decimal value ("2.5") before the
        // fraction regex runs, so they produce a distinct token from bare fractions.
        s = _mixedNumber.Replace(s, m =>
        {
            if (decimal.TryParse(m.Groups[1].Value, out var whole)
                && decimal.TryParse(m.Groups[2].Value, out var num)
                && decimal.TryParse(m.Groups[3].Value, out var den)
                && den != 0)
                return (whole + num / den).ToString("G29");
            return m.Value;
        });
        s = _dimFraction.Replace(s, "$1f$2");
        s = _dimDecimal.Replace(s, "$1d$2");
        s = Regex.Replace(s, @"d(\d+)", m =>
        {
            var stripped = m.Groups[1].Value.TrimEnd('0');
            return "d" + (stripped.Length == 0 ? "0" : stripped);
        });
        s = _dimSeparator.Replace(s, "$1x$2");
        s = _dimSeparator.Replace(s, "$1x$2");
        s = Regex.Replace(s, @"[""']", "");
        return s;
    }

    private static HashSet<string> ProductTokens(string s)
    {
        var p = PreprocessProduct(s);
        return _dimSplit.Split(p)
                        .Where(t => t.Length > 1 || (t.Length == 1 && char.IsDigit(t[0])))
                        .ToHashSet();
    }

    private static double ProductJaccard(HashSet<string> a, HashSet<string> b)
    {
        if (a.Count == 0 && b.Count == 0) return 1.0;
        var intersection = a.Count(t => b.Contains(t));
        var union        = a.Count + b.Count - intersection;
        return union == 0 ? 0 : (double)intersection / union;
    }

    private static bool HasDigit(string t) => t.Any(char.IsDigit);
    private static bool IsDimToken(string t) =>
        t.Any(char.IsDigit) && t.Any(c => c == 'x' || c == 'f' || c == 'd');

    private static bool NumericTokensCompatible(HashSet<string> a, HashSet<string> b)
    {
        var numA = a.Where(HasDigit).ToHashSet();
        var numB = b.Where(HasDigit).ToHashSet();
        var dimA = numA.Where(IsDimToken).ToHashSet();
        var dimB = numB.Where(IsDimToken).ToHashSet();

        if (dimA.Count > 0 && dimB.Count > 0)
        {
            // Compare the underlying numeric values in the dim tokens rather than the
            // raw token strings.  "MC6 ?? --  18" and "MC 6 x 18" tokenise differently:
            // the former produces dim={6x3d5x0d475} + plain=18, the latter produces
            // dim={6x18, 6x3d5x0d475}.  Extracting digits from dim tokens and combining
            // with standalone plain-digit tokens gives the same number set for both.
            var plainDigitA = numA.Where(t => !IsDimToken(t) && t.All(char.IsDigit)).ToHashSet();
            var plainDigitB = numB.Where(t => !IsDimToken(t) && t.All(char.IsDigit)).ToHashSet();
            // Also include the leading digit run from mixed alloy/temper tokens like "3003h14" → "3003".
            // Without this, one extraction writing "3003H14" and another writing "3003" produce
            // different allNums sets, causing a false incompatibility for the same product.
            static IEnumerable<string> LeadingDigits(IEnumerable<string> tokens) =>
                tokens
                    .Where(t => !IsDimToken(t) && !t.All(char.IsDigit) && t.Length > 0 && char.IsDigit(t[0]))
                    .Select(t => new string(t.TakeWhile(char.IsDigit).ToArray()))
                    .Where(s => s.Length > 0);
            var allNumsA = ExtractDimNumbers(dimA).Union(plainDigitA).Union(LeadingDigits(numA)).ToHashSet();
            var allNumsB = ExtractDimNumbers(dimB).Union(plainDigitB).Union(LeadingDigits(numB)).ToHashSet();
            if (!allNumsA.SetEquals(allNumsB)) return false;

            var gradeA = numA.Where(t => !IsDimToken(t) && !t.All(char.IsDigit)).ToHashSet();
            var gradeB = numB.Where(t => !IsDimToken(t) && !t.All(char.IsDigit)).ToHashSet();
            return gradeA.IsSubsetOf(gradeB) || gradeB.IsSubsetOf(gradeA);
        }
        // One side has dim tokens, the other doesn't  -- e.g. "A500 Pipe 6 SCH 40 (6.625 OD x 0.280 wall) x 21'"
        // vs the simplified re-extraction "A500 Pipe 6 SCH 40 x 21'".
        // Treat as compatible when the simpler side's plain numbers are a subset of the
        // richer side's expanded numbers (the nominal size is still represented).
        if (dimA.Count > 0)
        {
            var pA       = numA.Where(t => !IsDimToken(t) && t.All(char.IsDigit)).ToHashSet();
            var allNumsA = ExtractDimNumbers(dimA).Union(pA).ToHashSet();
            var plainB   = numB.Where(t => !IsDimToken(t) && t.All(char.IsDigit)).ToHashSet();
            return plainB.IsSubsetOf(allNumsA);
        }
        if (dimB.Count > 0)
        {
            var pB       = numB.Where(t => !IsDimToken(t) && t.All(char.IsDigit)).ToHashSet();
            var allNumsB = ExtractDimNumbers(dimB).Union(pB).ToHashSet();
            var plainA   = numA.Where(t => !IsDimToken(t) && t.All(char.IsDigit)).ToHashSet();
            return plainA.IsSubsetOf(allNumsB);
        }

        var gA = numA.Where(t => !IsDimToken(t)).ToHashSet();
        var gB = numB.Where(t => !IsDimToken(t)).ToHashSet();
        return gA.IsSubsetOf(gB) || gB.IsSubsetOf(gA);
    }

    /// <summary>
    /// Splits dim tokens on the dimension separators (x, f, d) used by
    /// <see cref="PreprocessProduct"/> and returns the individual digit strings.
    /// e.g. "6x3d5x0d475" ?????' {"6","3","5","0","475"}
    /// </summary>
    private static HashSet<string> ExtractDimNumbers(HashSet<string> dimTokens)
    {
        var result = new HashSet<string>();
        foreach (var tok in dimTokens)
            foreach (var part in tok.Split(new[] { 'x', 'f', 'd' }, StringSplitOptions.RemoveEmptyEntries))
            {
                if (part.Length == 0) continue;
                if (part.All(char.IsDigit))
                {
                    result.Add(part);
                }
                else if (char.IsDigit(part[0]))
                {
                    // Digit-leading mixed token like "11ga" (gauge) or "14ga" — extract the
                    // leading numeric portion so gauge differences are compared as numbers.
                    // Letter-leading tokens like "t6511" (alloy) are skipped (not dim data).
                    var leadingDigits = new string(part.TakeWhile(char.IsDigit).ToArray());
                    if (leadingDigits.Length > 0) result.Add(leadingDigits);
                }
            }
        return result;
    }

    // ?????? Attachment upload (Graph drive API) ??????????????????????????????????????????????????????????????????????????????????????????????????????

    /// <summary>
    /// Cached drive ID for the RFQ site's default document library.
    /// </summary>
    private string? _rfqDriveId;

    private async Task<string> GetRfqDriveIdAsync()
    {
        if (_rfqDriveId is not null) return _rfqDriveId;
        var siteId = await GetSiteIdAsync();
        var drive  = await GetGraph().Sites[siteId].Drive.GetAsync();
        _rfqDriveId = drive!.Id ?? throw new Exception("Could not resolve RFQ site drive ID");
        return _rfqDriveId;
    }

    /// <summary>
    /// Uploads a file to the site's default drive at QuoteAttachments/{srItemId}/{fileName}.
    /// Uses Graph PUT content which supports app-only tokens.
    /// </summary>
    private async Task UpsertItemAttachmentAsync(string spItemId, string listId, string fileName, byte[] bytes)
    {
        var driveId = await GetRfqDriveIdAsync();
        var itemKey = $"root:/QuoteAttachments/{spItemId}/{fileName}:";

        using var stream = new MemoryStream(bytes);
        await GetGraph().Drives[driveId].Items[itemKey].Content.PutAsync(stream);

        _log.LogInformation("[SP] Uploaded attachment '{File}' ({Bytes} bytes) to drive for SR item {Id}",
            fileName, bytes.Length, spItemId);
    }

    // ??"?????"??? Backfill: write CatalogProductName / ProductSearchKey for existing SLI rows ??"?????"?????"?????"?????"???

    /// <summary>
    /// Iterates every SupplierLineItem and writes the current catalog match result
    /// to <c>CatalogProductName</c> and <c>ProductSearchKey</c>.
    /// Safe to run repeatedly  -- idempotent patch, no rows created or deleted.
    /// Returns (total rows visited, rows updated, rows with a match).
    /// </summary>
    public async Task<(int Total, int Updated, int Matched)> BackfillCatalogMatchesAsync(
        CancellationToken ct = default)
    {
        var siteId    = await GetSiteIdAsync();
        var sliListId = await GetSupplierLineItemsListIdAsync();
        int total = 0, updated = 0, matched = 0;

        // Page through all SLI rows, reading only the fields we need.
        var page = await GetGraph().Sites[siteId].Lists[sliListId].Items
            .GetAsync(r =>
            {
                r.QueryParameters.Expand = ["fields($select=id,ProductName,CatalogProductName,ProductSearchKey)"];
                r.QueryParameters.Top    = 500;
            }, ct);

        while (page?.Value is not null)
        {
            foreach (var item in page.Value)
            {
                ct.ThrowIfCancellationRequested();
                total++;

                var data    = item.Fields?.AdditionalData;
                var itemId  = item.Id;
                if (data is null || itemId is null) continue;

                var prodName = data.TryGetValue("ProductName", out var pn) ? pn?.ToString() : null;
                if (string.IsNullOrWhiteSpace(prodName)) continue;

                var match = _catalog.ResolveProduct(prodName);
                if (match is not null) matched++;

                // Always patch  -- clears stale values when catalog changes remove a match.
                var patch = new Dictionary<string, object?>
                {
                    ["CatalogProductName"] = match?.Name,
                    ["ProductSearchKey"]   = match?.SearchKey,
                };
                await GetGraph().Sites[siteId].Lists[sliListId].Items[itemId].Fields
                    .PatchAsync(new FieldValueSet { AdditionalData = patch }, cancellationToken: ct);
                updated++;
            }

            // Follow nextLink for the next page.
            if (page.OdataNextLink is null) break;
            page = await GetGraph().Sites[siteId].Lists[sliListId].Items
                .WithUrl(page.OdataNextLink)
                .GetAsync(cancellationToken: ct);
        }

        _log.LogInformation("[SP] Backfill complete: {Total} rows, {Updated} patched, {Matched} matched",
            total, updated, matched);
        return (total, updated, matched);
    }

    // ??"?????"??? Clean: delete all derived email-processing data ??"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"???

    /// <summary>
    /// Deletes a single SupplierLineItem row by its SharePoint item ID.
    /// </summary>
    public async Task DeleteSliAsync(string itemId)
    {
        var siteId    = await GetSiteIdAsync();
        var sliListId = await GetSupplierLineItemsListIdAsync();
        await GetGraph().Sites[siteId].Lists[sliListId].Items[itemId].DeleteAsync();
        _log.LogInformation("[SP] Deleted SLI item {Id}", itemId);
    }

    /// <summary>
    /// Deletes a SupplierResponse row and all SupplierLineItem rows that reference it.
    /// Returns (sliDeleted, srDeleted) counts.
    /// </summary>
    public async Task<(int SliDeleted, int SrDeleted)> DeleteSrAsync(string srId)
    {
        var siteId    = await GetSiteIdAsync();
        var sliListId = await GetSupplierLineItemsListIdAsync();
        var srListId  = await GetSupplierResponsesListIdAsync();

        // Find all SLI rows that reference this SR
        var page = await GetGraph().Sites[siteId].Lists[sliListId].Items
            .GetAsync(r =>
            {
                r.QueryParameters.Expand = ["fields($select=id,SupplierResponseId)"];
                r.QueryParameters.Top    = 2000;
            });

        var sliIds = new List<string>();
        while (page is not null)
        {
            foreach (var item in page.Value ?? [])
            {
                var d = item.Fields?.AdditionalData;
                if (d is null || item.Id is null) continue;
                var refId = d.TryGetValue("SupplierResponseId", out var v) ? v?.ToString() : null;
                if (string.Equals(refId, srId, StringComparison.OrdinalIgnoreCase))
                    sliIds.Add(item.Id);
            }
            if (page.OdataNextLink is null) break;
            var next = new Microsoft.Graph.Sites.Item.Lists.Item.Items.ItemsRequestBuilder(
                page.OdataNextLink, GetGraph().RequestAdapter);
            page = await next.GetAsync();
        }

        foreach (var id in sliIds)
        {
            await GetGraph().Sites[siteId].Lists[sliListId].Items[id].DeleteAsync();
            _log.LogInformation("[SP] Deleted SLI {Id} (child of SR {Sr})", id, srId);
        }

        await GetGraph().Sites[siteId].Lists[srListId].Items[srId].DeleteAsync();
        InvalidateSrCache();
        _log.LogInformation("[SP] Deleted SR {Id} ({SliCount} SLIs removed)", srId, sliIds.Count);

        return (sliIds.Count, 1);
    }

    /// <summary>
    /// Re-parents a SupplierResponse and all its child SupplierLineItems to a new RFQ ID.
    /// Updates RFQ_ID on both the SR row and every SLI row that references it via SupplierResponseId.
    /// </summary>
    public async Task ReparentSupplierResponseAsync(string srId, string rfqId)
    {
        var siteId    = await GetSiteIdAsync();
        var sliListId = await GetSupplierLineItemsListIdAsync();
        var srListId  = await GetSupplierResponsesListIdAsync();

        // Find child SLI rows by SupplierResponseId
        var page = await GetGraph().Sites[siteId].Lists[sliListId].Items
            .GetAsync(r =>
            {
                r.QueryParameters.Expand = ["fields($select=id,SupplierResponseId)"];
                r.QueryParameters.Top    = 2000;
            });

        var sliIds = new List<string>();
        while (page is not null)
        {
            foreach (var item in page.Value ?? [])
            {
                var d = item.Fields?.AdditionalData;
                if (d is null || item.Id is null) continue;
                var refId = d.TryGetValue("SupplierResponseId", out var v) ? v?.ToString() : null;
                if (string.Equals(refId, srId, StringComparison.OrdinalIgnoreCase))
                    sliIds.Add(item.Id);
            }
            if (page.OdataNextLink is null) break;
            var next = new Microsoft.Graph.Sites.Item.Lists.Item.Items.ItemsRequestBuilder(
                page.OdataNextLink, GetGraph().RequestAdapter);
            page = await next.GetAsync();
        }

        await GetGraph().Sites[siteId].Lists[srListId].Items[srId].Fields
            .PatchAsync(new FieldValueSet { AdditionalData = new Dictionary<string, object?> { ["RFQ_ID"] = rfqId } });

        foreach (var sliId in sliIds)
        {
            await GetGraph().Sites[siteId].Lists[sliListId].Items[sliId].Fields
                .PatchAsync(new FieldValueSet { AdditionalData = new Dictionary<string, object?> { ["RFQ_ID"] = rfqId } });
        }

        InvalidateSrCache();
        _log.LogInformation("[SP] Reparented SR {SrId} and {N} SLI row(s) to RFQ {RfqId}", srId, sliIds.Count, rfqId);
    }

    /// <summary>
    /// Re-parents a single SupplierLineItem to a new RFQ ID.
    /// If the source SR has no remaining SLI children, it is deleted.
    /// Returns true when the source SR was deleted.
    /// </summary>
    public async Task<bool> ReparentSliAsync(string sliItemId, string rfqId)
    {
        var siteId    = await GetSiteIdAsync();
        var sliListId = await GetSupplierLineItemsListIdAsync();
        var srListId  = await GetSupplierResponsesListIdAsync();

        // Read the SLI to discover its parent SR before patching
        var sliItem = await GetGraph().Sites[siteId].Lists[sliListId].Items[sliItemId]
            .GetAsync(r => r.QueryParameters.Expand = ["fields($select=SupplierResponseId)"]);
        var parentSrId = sliItem?.Fields?.AdditionalData
            ?.TryGetValue("SupplierResponseId", out var sv) == true ? sv?.ToString() : null;

        // Patch the SLI's own RFQ_ID to the target RFQ
        await GetGraph().Sites[siteId].Lists[sliListId].Items[sliItemId].Fields
            .PatchAsync(new FieldValueSet
            {
                AdditionalData = new Dictionary<string, object?> { ["RFQ_ID"] = rfqId }
            });

        _log.LogInformation("[SP] Reparented SLI {SliId} to RFQ {RfqId}", sliItemId, rfqId);

        if (string.IsNullOrEmpty(parentSrId)) return false;

        // Count SLI rows still referencing the source SR; break early once we know there are 2+
        var page = await GetGraph().Sites[siteId].Lists[sliListId].Items
            .GetAsync(r =>
            {
                r.QueryParameters.Expand = ["fields($select=SupplierResponseId)"];
                r.QueryParameters.Top    = 2000;
            });

        int siblingCount = 0;
        while (page is not null && siblingCount < 2)
        {
            foreach (var item in page.Value ?? [])
            {
                var d = item.Fields?.AdditionalData;
                if (d is null) continue;
                var refId = d.TryGetValue("SupplierResponseId", out var rv) ? rv?.ToString() : null;
                if (string.Equals(refId, parentSrId, StringComparison.OrdinalIgnoreCase))
                    siblingCount++;
                if (siblingCount >= 2) break;
            }
            if (page.OdataNextLink is null || siblingCount >= 2) break;
            var next = new Microsoft.Graph.Sites.Item.Lists.Item.Items.ItemsRequestBuilder(
                page.OdataNextLink, GetGraph().RequestAdapter);
            page = await next.GetAsync();
        }

        // siblingCount == 1 means only our just-reparented SLI remains (still points at old SR)
        if (siblingCount <= 1)
        {
            await GetGraph().Sites[siteId].Lists[srListId].Items[parentSrId].DeleteAsync();
            InvalidateSrCache();
            _log.LogInformation("[SP] Deleted empty SR {SrId} after last SLI reparented to {RfqId}", parentSrId, rfqId);
            return true;
        }

        return false;
    }

    /// <summary>
    /// Deletes every item in SupplierResponses and SupplierLineItems.
    /// Returns counts of items deleted from each list.
    /// Does NOT touch RFQ References (notes / dates).
    /// </summary>
    /// <summary>
    /// Deletes all rows from all four RFQ lists: RFQ References, RFQ Line Items,
    /// SupplierResponses, and SupplierLineItems.  Order: SLI ?????' SR ?????' RLI ?????' RR
    /// (child before parent to respect any referential integrity).
    /// </summary>
    public async Task<(int RefsDeleted, int RliDeleted, int SrDeleted, int SliDeleted)> CleanAllDataAsync()
    {
        var siteId    = await GetSiteIdAsync();
        var rrListId  = await GetRfqReferencesListIdAsync();
        var rliListId = await GetRfqLineItemsListIdAsync();
        var srListId  = await GetSupplierResponsesListIdAsync();
        var sliListId = await GetSupplierLineItemsListIdAsync();

        var sliDeleted = await DeleteAllItemsAsync(siteId, sliListId, "SupplierLineItems");
        var srDeleted  = await DeleteAllItemsAsync(siteId, srListId,  "SupplierResponses");
        InvalidateSrCache();
        var rliDeleted = await DeleteAllItemsAsync(siteId, rliListId, "RFQ Line Items");
        var rrDeleted  = await DeleteAllItemsAsync(siteId, rrListId,  "RFQ References");

        return (rrDeleted, rliDeleted, srDeleted, sliDeleted);
    }

    public async Task<(int SrDeleted, int SliDeleted)> CleanSupplierDataAsync()
    {
        var siteId    = await GetSiteIdAsync();
        var srListId  = await GetSupplierResponsesListIdAsync();
        var sliListId = await GetSupplierLineItemsListIdAsync();

        var srDeleted  = await DeleteAllItemsAsync(siteId, srListId,  "SupplierResponses");
        var sliDeleted = await DeleteAllItemsAsync(siteId, sliListId, "SupplierLineItems");
        InvalidateSrCache();

        return (srDeleted, sliDeleted);
    }

    /// <summary>
    /// Deletes SupplierResponses (and their child SLIs) where the email ReceivedAt date
    /// is older than <paramref name="days"/> days.
    /// </summary>
    public async Task<(int SrDeleted, int SliDeleted)> PurgeOldSupplierDataAsync(int days)
    {
        var cutoff    = DateTimeOffset.UtcNow.AddDays(-days);
        var siteId    = await GetSiteIdAsync();
        var srListId  = await GetSupplierResponsesListIdAsync();
        var sliListId = await GetSupplierLineItemsListIdAsync();

        // Collect SR item IDs where ReceivedAt < cutoff
        var oldSrIds = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        var srPage = await GetGraph().Sites[siteId].Lists[srListId].Items
            .GetAsync(req =>
            {
                req.QueryParameters.Expand = ["fields($select=id,ReceivedAt)"];
                req.QueryParameters.Top    = 2000;
            });

        while (srPage?.Value is not null)
        {
            foreach (var item in srPage.Value)
            {
                if (item.Id is null || item.Fields?.AdditionalData is not { } d) continue;
                var recStr = GetStr(d, "ReceivedAt");
                if (DateTimeOffset.TryParse(recStr, out var rec) && rec < cutoff)
                    oldSrIds.Add(item.Id);
            }
            if (srPage.OdataNextLink is null) break;
            srPage = await GetGraph().Sites[siteId].Lists[srListId].Items
                .WithUrl(srPage.OdataNextLink).GetAsync();
        }

        _log.LogInformation("[SP] Purge: {Count} SR row(s) with ReceivedAt older than {Days} days",
            oldSrIds.Count, days);

        if (oldSrIds.Count == 0) return (0, 0);

        // Delete child SLI rows first
        int sliDeleted = 0;
        var sliPage = await GetGraph().Sites[siteId].Lists[sliListId].Items
            .GetAsync(req =>
            {
                req.QueryParameters.Expand = ["fields($select=id,SupplierResponseId)"];
                req.QueryParameters.Top    = 2000;
            });

        while (sliPage?.Value is not null)
        {
            foreach (var item in sliPage.Value)
            {
                if (item.Id is null || item.Fields?.AdditionalData is not { } d) continue;
                var srId = GetStr(d, "SupplierResponseId");
                if (srId is null || !oldSrIds.Contains(srId)) continue;
                await GetGraph().Sites[siteId].Lists[sliListId].Items[item.Id].DeleteAsync();
                sliDeleted++;
            }
            if (sliPage.OdataNextLink is null) break;
            sliPage = await GetGraph().Sites[siteId].Lists[sliListId].Items
                .WithUrl(sliPage.OdataNextLink).GetAsync();
        }

        // Delete SR rows
        int srDeleted = 0;
        foreach (var spId in oldSrIds)
        {
            await GetGraph().Sites[siteId].Lists[srListId].Items[spId].DeleteAsync();
            srDeleted++;
        }

        _log.LogInformation("[SP] Purge complete  -- SR deleted={Sr}, SLI deleted={Sli}", srDeleted, sliDeleted);
        return (srDeleted, sliDeleted);
    }

    private async Task<int> DeleteAllItemsAsync(string siteId, string listId, string listName)
    {
        int deleted = 0;

        while (true)
        {
            // Fetch a page of item IDs only  -- no fields needed
            var page = await GetGraph().Sites[siteId].Lists[listId].Items
                .GetAsync(req => { req.QueryParameters.Top = 100; });

            var items = page?.Value;
            if (items is null || items.Count == 0) break;

            // Delete in parallel (Graph throttles at ~4 concurrent writes; keep modest)
            var tasks = items
                .Where(i => i.Id is not null)
                .Select(i => GetGraph().Sites[siteId].Lists[listId].Items[i.Id!].DeleteAsync());

            await Task.WhenAll(tasks);
            deleted += items.Count;
            _log.LogInformation("[SP] Deleted {Count} items from {List} (total so far: {Total})",
                items.Count, listName, deleted);

            // If fewer items than page size were returned, we're done
            if (items.Count < 100) break;
        }

        _log.LogInformation("[SP] Finished cleaning {List}: {Total} items deleted", listName, deleted);
        return deleted;
    }

    // ??"?????"??? Deduplicate SupplierResponses ??"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"???

    // Report models  -- populated in both live and dry-run modes
    public record DedupeReportSli(
        string Action,            // "delete" | "reparent"
        string SliId,
        string ProductName,
        bool   WouldRescueComments);

    public record DedupeReportRetiring(
        string   SrId,
        int      Score,
        string   ProcessingSource,
        string?  SourceFile,
        string[] FieldsToMerge,
        DedupeReportSli[] Slis);

    public record DedupeReportGroup(
        string   RfqId,
        string   Supplier,
        string   KeeperSrId,
        int      KeeperScore,
        string   KeeperProcessingSource,
        DedupeReportRetiring[] Retiring);

    public record DedupeReportSliDupe(
        string SliId,
        string ProductName,
        bool   WouldRescueComments);

    public record DedupeReportSliDupeGroup(
        string   SrId,
        string   RfqId,
        string   Supplier,
        string   KeeperSliId,
        string   KeeperProductName,
        DedupeReportSliDupe[] Retiring);

    public record DedupeSupplierResponsesResult(
        bool     DryRun,
        int      DuplicateGroups,
        int      SrDeleted,
        int      SliReparented,
        int      SliDeleted,
        List<DedupeReportGroup> Groups,
        int      SliDuplicateGroups,
        int      SliWithinSrDeleted,
        List<DedupeReportSliDupeGroup> SliGroups);

    /// <summary>
    /// Finds SupplierResponse rows that share the same (RFQ_ID, SupplierName) and merges
    /// each duplicate group into a single canonical row.
    ///
    /// For each duplicate group:
    ///    --?? Keep the SR with the best data: attachment rows beat body rows; priced SLIs beat
    ///     unpriceds; newest DateCreated breaks ties.
    ///    --?? For each SLI under a duplicate SR:
    ///        -- If the keeper already has an SLI for the same product ?????' delete the duplicate SLI.
    ///        -- Otherwise ?????' re-parent the SLI to the keeper.
    ///    --?? Delete the duplicate SR.
    /// </summary>
    public async Task<DedupeSupplierResponsesResult> DedupeSupplierResponsesAsync(bool dryRun = false, string? rfqId = null)
    {
        var siteId    = await GetSiteIdAsync();
        var srListId  = await GetSupplierResponsesListIdAsync();
        var sliListId = await GetSupplierLineItemsListIdAsync();

        // ??"?????"??? Fetch all SR rows ??"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"???
        // RFQ_x005F_ID is the SP internal column name in some list configurations;
        // select both so the fallback in FldRfqId always finds a value.
        var srResponse = await GetGraph().Sites[siteId].Lists[srListId].Items
            .GetAsync(r =>
            {
                r.QueryParameters.Expand = ["fields($select=id,RFQ_ID,RFQ_x005F_ID,SupplierName," +
                    "ProcessingSource,SourceFile,QuoteReference,DateOfQuote," +
                    "EstimatedDeliveryDate,FreightTerms,EmailBody)"];
                r.QueryParameters.Top = 5000;
            });
        var srItems = srResponse?.Value ?? [];
        if (srItems.Count >= 5000)
            _log.LogWarning("[Dedupe-SR] SR fetch hit the 5 000-row limit  -- re-run to catch any remaining duplicates");

        // ??"?????"??? Fetch all SLI rows ??"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"???
        // SupplierProductComments included so we can rescue AI commentary
        // before deleting a duplicate SLI that covers the same product.
        var sliResponse = await GetGraph().Sites[siteId].Lists[sliListId].Items
            .GetAsync(r =>
            {
                r.QueryParameters.Expand = ["fields($select=id,SupplierResponseId,ProductName," +
                    "PricePerPound,PricePerFoot,PricePerPiece,TotalPrice,SupplierProductComments,MSPCMatch)"];
                r.QueryParameters.Top = 5000;
            });
        var sliItems = sliResponse?.Value ?? [];
        if (sliItems.Count >= 5000)
            _log.LogWarning("[Dedupe-SR] SLI fetch hit the 5 000-row limit  -- re-run to catch any remaining duplicates");

        // ??"?????"??? Helpers ??"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"???
        static string? Fld(ListItem item, string key)
        {
            var d = item.Fields?.AdditionalData;
            if (d is null || !d.TryGetValue(key, out var v)) return null;
            return v is JsonElement je
                ? (je.ValueKind == JsonValueKind.String ? je.GetString() : je.ToString())
                : v?.ToString();
        }

        // RFQ_ID may be stored under either name depending on list creation history
        static string? FldRfqId(ListItem sr) =>
            Fld(sr, "RFQ_ID") ?? Fld(sr, "RFQ_x005F_ID");

        static bool SliHasPrice(ListItem sli)
        {
            var d = sli.Fields?.AdditionalData;
            if (d is null) return false;
            foreach (var key in (string[])["PricePerPound", "PricePerFoot", "PricePerPiece", "TotalPrice"])
            {
                if (!d.TryGetValue(key, out var v) || v is null) continue;
                var n = v is JsonElement je && je.ValueKind == JsonValueKind.Number
                    ? je.GetDouble() : (double?)null;
                if (n.HasValue && n.Value > 0) return true;
            }
            return false;
        }

        static bool ProductsMatch(HashSet<string> tokA, HashSet<string> tokB) =>
            NumericTokensCompatible(tokA, tokB) && ProductJaccard(tokA, tokB) >= 0.5;

        // Index SLIs by SupplierResponseId
        var slisBySrId = sliItems
            .GroupBy(sli => Fld(sli, "SupplierResponseId") ?? "")
            .ToDictionary(g => g.Key, g => g.ToList());

        // ??"?????"??? Find duplicate groups ??"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"???
        var rfqFilter = rfqId?.Trim().ToUpperInvariant();
        var duplicateGroups = srItems
            .GroupBy(sr => (
                RfqId:    (FldRfqId(sr) ?? "").ToUpperInvariant(),
                Supplier: (Fld(sr, "SupplierName") ?? "").ToLowerInvariant()))
            .Where(g => g.Key.RfqId.Length > 0 && g.Key.Supplier.Length > 0 && g.Count() > 1
                     && (rfqFilter is null || g.Key.RfqId == rfqFilter))
            .ToList();

        _log.LogInformation("[Dedupe-SR] Found {G} duplicate group(s) across {T} total SR rows",
            duplicateGroups.Count, srItems.Count);

        int srDeleted = 0, sliDeleted = 0, sliReparented = 0;
        var reportGroups = new List<DedupeReportGroup>();
        var dryTag = dryRun ? "[DRY] " : "";

        foreach (var group in duplicateGroups)
        {
            var srsInGroup = group.ToList();

            // Score: attachment > has priced SLI > newest DateCreated (tiebreak)
            int Score(ListItem sr)
            {
                int s    = 0;
                var src  = Fld(sr, "ProcessingSource") ?? "";
                var file = Fld(sr, "SourceFile")       ?? "";
                if (src.Equals("attachment", StringComparison.OrdinalIgnoreCase) || file.Length > 0)
                    s += 2;
                if (slisBySrId.GetValueOrDefault(sr.Id ?? "", []).Any(SliHasPrice))
                    s += 1;
                return s;
            }

            var keeper = srsInGroup
                .OrderByDescending(Score)
                .ThenByDescending(sr => sr.CreatedDateTime)
                .First();

            // keeper's SLIs as a mutable list  -- entries are added as SLIs are re-parented
            // so subsequent dupes in the same group see the updated coverage.
            var keeperSlis = slisBySrId.GetValueOrDefault(keeper.Id ?? "", []).ToList();

            _log.LogInformation("{Tag}[Dedupe-SR] Group [{Rfq}] {Supplier}: keeping SR {Keep}, retiring {Dupes}",
                dryTag, group.Key.RfqId, group.Key.Supplier, keeper.Id,
                string.Join(", ", srsInGroup.Where(s => s.Id != keeper.Id).Select(s => s.Id)));

            var reportRetiring = new List<DedupeReportRetiring>();

            foreach (var dupe in srsInGroup.Where(sr => sr.Id != keeper.Id))
            {
                // ?????? Merge SR-level AI content into keeper ??????????????????????????????
                // Only fill blank keeper fields. Update the in-memory dict after each
                // merge so subsequent dupes in the same group see the updated values
                // and don't try to merge the same field twice.
                var mergeFields = new[] { "QuoteReference", "DateOfQuote",
                                          "EstimatedDeliveryDate", "FreightTerms", "EmailBody" };
                var toMerge = new Dictionary<string, object?>();
                foreach (var mf in mergeFields)
                {
                    var keeperVal = Fld(keeper, mf);
                    var dupeVal   = Fld(dupe,   mf);
                    if (string.IsNullOrWhiteSpace(keeperVal) && !string.IsNullOrWhiteSpace(dupeVal))
                        toMerge[mf] = dupeVal;
                }
                if (toMerge.Count > 0)
                {
                    if (!dryRun)
                    {
                        await GetGraph().Sites[siteId].Lists[srListId].Items[keeper.Id!].Fields
                            .PatchAsync(new FieldValueSet { AdditionalData = toMerge! });
                        // Reflect in-memory so subsequent dupes don't re-merge the same field
                        if (keeper.Fields?.AdditionalData is { } keeperDict)
                            foreach (var (k, v) in toMerge)
                                keeperDict[k] = v;
                    }
                    _log.LogInformation("{Tag}[Dedupe-SR] Merged {Fields} from retiring SR {From} ?????' keeper {To}",
                        dryTag, string.Join(", ", toMerge.Keys), dupe.Id, keeper.Id);
                }

                // ??"?????"??? Handle this dupe's SLIs ??"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"???
                var dupeSlis = slisBySrId.GetValueOrDefault(dupe.Id ?? "", []);
                var reportSlis = new List<DedupeReportSli>();

                foreach (var sli in dupeSlis)
                {
                    var prodName     = Fld(sli, "ProductName") ?? "";
                    var prodTok      = ProductTokens(prodName);
                    var dupeComments = Fld(sli, "SupplierProductComments");

                    // Match against keeper SLIs by original name (NormalizeMatch) or
                    // token Jaccard  -- avoids the broken "join tokens ?????' NormalizeMatch" pattern.
                    var coveringKeeperSli = keeperSlis.FirstOrDefault(k =>
                    {
                        var kName = Fld(k, "ProductName") ?? "";
                        return NormalizeMatch(prodName, kName)
                            || ProductsMatch(prodTok, ProductTokens(kName));
                    });

                    if (coveringKeeperSli is not null)
                    {
                        // Product already covered by keeper.
                        // Rescue SupplierProductComments and MSPCMatch before deleting
                        // if the keeper SLI has none.
                        var rescueFields = new Dictionary<string, object?>();
                        bool wouldRescue = false;

                        if (!string.IsNullOrWhiteSpace(dupeComments) &&
                            string.IsNullOrWhiteSpace(Fld(coveringKeeperSli, "SupplierProductComments")))
                        {
                            rescueFields["SupplierProductComments"] = dupeComments;
                            wouldRescue = true;
                        }

                        var dupeMspc   = Fld(sli, "MSPCMatch");
                        var keeperMspc = Fld(coveringKeeperSli, "MSPCMatch");
                        if (!string.IsNullOrWhiteSpace(dupeMspc) && string.IsNullOrWhiteSpace(keeperMspc))
                            rescueFields["MSPCMatch"] = dupeMspc;

                        if (rescueFields.Count > 0)
                        {
                            if (!dryRun)
                            {
                                await GetGraph().Sites[siteId].Lists[sliListId]
                                    .Items[coveringKeeperSli.Id!].Fields
                                    .PatchAsync(new FieldValueSet { AdditionalData = rescueFields! });
                                // Reflect in-memory so subsequent dupes see the rescued values
                                if (coveringKeeperSli.Fields?.AdditionalData is { } sliDict)
                                    foreach (var (k, v) in rescueFields)
                                        sliDict[k] = v;
                            }
                            _log.LogInformation(
                                "{Tag}[Dedupe-SR] Rescued {Fields} from SLI {From} ?????' keeper SLI {To} ('{Product}')",
                                dryTag, string.Join(", ", rescueFields.Keys), sli.Id, coveringKeeperSli.Id, prodName);
                        }

                        if (!dryRun)
                            await GetGraph().Sites[siteId].Lists[sliListId].Items[sli.Id!].DeleteAsync();
                        sliDeleted++;
                        reportSlis.Add(new DedupeReportSli("delete", sli.Id!, prodName, wouldRescue));
                        _log.LogInformation("{Tag}[Dedupe-SR] Deleted duplicate SLI {Id} ('{Product}') from retiring SR {Sr}",
                            dryTag, sli.Id, prodName, dupe.Id);
                    }
                    else
                    {
                        // Not covered  -- re-parent to keeper and track in-memory
                        if (!dryRun)
                        {
                            await GetGraph().Sites[siteId].Lists[sliListId].Items[sli.Id!].Fields
                                .PatchAsync(new FieldValueSet
                                {
                                    AdditionalData = new Dictionary<string, object?>
                                        { ["SupplierResponseId"] = keeper.Id }
                                });
                            keeperSlis.Add(sli); // track so subsequent dupes see this product as covered
                        }
                        sliReparented++;
                        reportSlis.Add(new DedupeReportSli("reparent", sli.Id!, prodName, false));
                        _log.LogInformation("{Tag}[Dedupe-SR] Re-parented SLI {Id} ('{Product}') from SR {From} ?????' {To}",
                            dryTag, sli.Id, prodName, dupe.Id, keeper.Id);
                    }
                }

                if (!dryRun)
                    await GetGraph().Sites[siteId].Lists[srListId].Items[dupe.Id!].DeleteAsync();
                srDeleted++;
                reportRetiring.Add(new DedupeReportRetiring(
                    dupe.Id!,
                    Score(dupe),
                    Fld(dupe, "ProcessingSource") ?? "",
                    Fld(dupe, "SourceFile"),
                    [.. toMerge.Keys],
                    [.. reportSlis]));
                _log.LogInformation("{Tag}[Dedupe-SR] Deleted duplicate SR {Id} for [{Rfq}] {Supplier}",
                    dryTag, dupe.Id, group.Key.RfqId, group.Key.Supplier);
            }

            reportGroups.Add(new DedupeReportGroup(
                group.Key.RfqId,
                group.Key.Supplier,
                keeper.Id!,
                Score(keeper),
                Fld(keeper, "ProcessingSource") ?? "",
                [.. reportRetiring]));
        }

        _log.LogInformation("{Tag}[Dedupe-SR] Done  -- {G} groups, {Sr} SR deleted, {SliR} SLI re-parented, {SliD} SLI deleted",
            dryTag, duplicateGroups.Count, srDeleted, sliReparented, sliDeleted);

        // ??"?????"??? Pass 2: SLI-level dedup within each SR ??"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"???
        // Catches the case where the same attachment was processed multiple times
        // in a single run (SP write-lag means the just-written SLI isn't visible
        // to subsequent FindExistingSupplierLineItemAsync calls), producing several
        // SLI rows with slightly different product name wording but identical pricing
        // all under the same SR.
        int sliWithinSrDeleted = 0;
        var reportSliGroups    = new List<DedupeReportSliDupeGroup>();

        // Reverse index: SR ID ?????' SR item (for RFQ_ID / SupplierName in report)
        var srById = srItems.ToDictionary(sr => sr.Id ?? "", sr => sr);

        int SliScore(ListItem sli)
        {
            var d = sli.Fields?.AdditionalData;
            if (d is null) return 0;
            int s = 0;
            foreach (var key in (string[])["PricePerPound", "PricePerFoot", "PricePerPiece", "TotalPrice"])
            {
                if (!d.TryGetValue(key, out var v) || v is null) continue;
                var n = v is JsonElement je && je.ValueKind == JsonValueKind.Number ? je.GetDouble() : (double?)null;
                if (n.HasValue && n.Value > 0) s += 4;
            }
            foreach (var key in (string[])["UnitsQuoted", "WeightPerUnit", "LengthPerUnit", "Certifications"])
                if (d.TryGetValue(key, out var v) && !string.IsNullOrWhiteSpace(v?.ToString())) s += 1;
            // Prefer rows that have a confirmed catalog match
            if (d.TryGetValue("MSPCMatch", out var mspc) && !string.IsNullOrWhiteSpace(mspc?.ToString()))
                s += 4;
            return s;
        }

        foreach (var (srId, slis) in slisBySrId)
        {
            if (slis.Count <= 1) continue;

            // Greedy clustering: each SLI joins the first cluster whose representative
            // it matches. Two criteria (any one is sufficient):
            //   1. Exact-normalised product name (strips punctuation/whitespace), OR
            //   2. NumericTokensCompatible + Jaccard >= 0.75 (strict — within one email the AI
            //      should describe the same product almost identically; the high threshold
            //      prevents genuinely different products like "Round Bar" vs "Square Bar" or
            //      "11GA" vs "16GA" sheet from being collapsed into one row).
            // The previous criterion 3 (same price + Jaccard >= 0.3) was too loose and caused
            // different products with coincidentally equal prices to be merged.

            var clusters = new List<List<ListItem>>();
            foreach (var sli in slis)
            {
                var prodName = Fld(sli, "ProductName") ?? "";
                var prodTok  = ProductTokens(prodName);
                var cluster  = clusters.FirstOrDefault(c =>
                {
                    var repName = Fld(c[0], "ProductName") ?? "";
                    var repTok  = ProductTokens(repName);
                    if (NormalizeMatch(prodName, repName)) return true;
                    return NumericTokensCompatible(prodTok, repTok) && ProductJaccard(prodTok, repTok) >= 0.75;
                });
                if (cluster is not null) cluster.Add(sli);
                else clusters.Add([sli]);
            }

            foreach (var cluster in clusters.Where(c => c.Count > 1))
            {
                var sliKeeper = cluster
                    .OrderByDescending(SliScore)
                    .ThenByDescending(s => s.CreatedDateTime)
                    .First();

                var sr          = srById.GetValueOrDefault(srId);
                var sliRfqId    = sr is not null ? (FldRfqId(sr) ?? "") : "";
                var supplier    = sr is not null ? (Fld(sr, "SupplierName") ?? "") : "";
                var reportDupes = new List<DedupeReportSliDupe>();

                _log.LogInformation("{Tag}[Dedupe-SLI] SR {Sr} [{Rfq}] {Supplier}: keeping SLI {Keep}, retiring {Dupes}",
                    dryTag, srId, sliRfqId, supplier, sliKeeper.Id,
                    string.Join(", ", cluster.Where(s => s.Id != sliKeeper.Id).Select(s => s.Id)));

                foreach (var dupe in cluster.Where(s => s.Id != sliKeeper.Id))
                {
                    var dupeComments = Fld(dupe, "SupplierProductComments");
                    var rescueFields = new Dictionary<string, object?>();
                    bool wouldRescue = false;

                    if (!string.IsNullOrWhiteSpace(dupeComments) &&
                        string.IsNullOrWhiteSpace(Fld(sliKeeper, "SupplierProductComments")))
                    {
                        rescueFields["SupplierProductComments"] = dupeComments;
                        wouldRescue = true;
                    }

                    var dupeMspc   = Fld(dupe, "MSPCMatch");
                    var keeperMspc = Fld(sliKeeper, "MSPCMatch");
                    if (!string.IsNullOrWhiteSpace(dupeMspc) && string.IsNullOrWhiteSpace(keeperMspc))
                        rescueFields["MSPCMatch"] = dupeMspc;

                    if (rescueFields.Count > 0)
                    {
                        if (!dryRun)
                        {
                            await GetGraph().Sites[siteId].Lists[sliListId]
                                .Items[sliKeeper.Id!].Fields
                                .PatchAsync(new FieldValueSet { AdditionalData = rescueFields! });
                            if (sliKeeper.Fields?.AdditionalData is { } kd)
                                foreach (var (k, v) in rescueFields)
                                    kd[k] = v;
                        }
                        _log.LogInformation(
                            "{Tag}[Dedupe-SLI] Rescued {Fields} from SLI {From} ?????' keeper SLI {To}",
                            dryTag, string.Join(", ", rescueFields.Keys), dupe.Id, sliKeeper.Id);
                    }

                    if (!dryRun)
                        await GetGraph().Sites[siteId].Lists[sliListId].Items[dupe.Id!].DeleteAsync();
                    sliWithinSrDeleted++;
                    reportDupes.Add(new DedupeReportSliDupe(
                        dupe.Id!, Fld(dupe, "ProductName") ?? "", wouldRescue));
                    _log.LogInformation(
                        "{Tag}[Dedupe-SLI] Deleted duplicate SLI {Id} ('{Product}') within SR {Sr}",
                        dryTag, dupe.Id, Fld(dupe, "ProductName"), srId);
                }

                reportSliGroups.Add(new DedupeReportSliDupeGroup(
                    srId, rfqId, supplier,
                    sliKeeper.Id!, Fld(sliKeeper, "ProductName") ?? "",
                    [.. reportDupes]));
            }
        }

        _log.LogInformation("{Tag}[Dedupe-SLI] Done  -- {G} within-SR duplicate groups, {D} SLI deleted",
            dryTag, reportSliGroups.Count, sliWithinSrDeleted);

        return new DedupeSupplierResponsesResult(
            dryRun, duplicateGroups.Count, srDeleted, sliReparented, sliDeleted, reportGroups,
            reportSliGroups.Count, sliWithinSrDeleted, reportSliGroups);
    }

    // ??"?????"??? Fetch SP list item attachment (SupplierResponses PDF) ??"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"???

    /// <summary>
    /// Downloads the named attachment from a SupplierResponses list item via the SP REST API.
    /// Returns null if the item or file is not found.
    /// </summary>
    public async Task<(string ContentType, byte[] Bytes, string FileName)?> GetSpItemAttachmentAsync(
        string srItemId, string fileName)
    {
        var driveId = await GetRfqDriveIdAsync();
        var itemKey = $"root:/QuoteAttachments/{srItemId}/{fileName}:";

        try
        {
            using var stream = await GetGraph().Drives[driveId].Items[itemKey].Content.GetAsync();
            if (stream is null) return null;

            using var ms = new MemoryStream();
            await stream.CopyToAsync(ms);
            var bytes = ms.ToArray();

            var ext = Path.GetExtension(fileName).ToLowerInvariant();
            var ct = ext switch
            {
                ".pdf"  => "application/pdf",
                ".docx" => "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                ".doc"  => "application/msword",
                _        => "application/octet-stream",
            };

            _log.LogInformation("[SP] Fetched attachment '{File}' ({Bytes} bytes) from drive for SR item {Id}", fileName, bytes.Length, srItemId);
            return (ct, bytes, fileName);
        }
        catch (Microsoft.Graph.Models.ODataErrors.ODataError ex) when (ex.ResponseStatusCode == 404)
        {
            _log.LogWarning("[SP] Attachment not found in drive: srItemId={Id} file={File}", srItemId, fileName);
            return null;
        }
    }

    // ??"?????"??? Publish folder (Graph) ??"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"???

    // Cached publish site + drive IDs (separate from the RFQ/QC site).
    private string? _publishSiteId;
    private string? _publishDriveId;

    private async Task<(string SiteId, string DriveId)> GetPublishDriveAsync()
    {
        if (_publishSiteId is not null && _publishDriveId is not null)
            return (_publishSiteId, _publishDriveId);

        var siteUrl = _config["Publish:SiteUrl"]
            ?? throw new InvalidOperationException("Publish:SiteUrl is not configured in appsettings.json");

        var uri  = new Uri(siteUrl);
        var host = uri.Host;
        var path = uri.AbsolutePath;

        var site = await GetGraph().Sites[$"{host}:{path}"].GetAsync();
        _publishSiteId = site!.Id ?? throw new Exception($"Could not resolve publish SharePoint site ID for {siteUrl}");

        var drive = await GetGraph().Sites[_publishSiteId].Drive.GetAsync();
        _publishDriveId = drive!.Id ?? throw new Exception("Could not resolve publish drive ID");

        _log.LogInformation("[Publish] Site: {SiteId}  Drive: {DriveId}", _publishSiteId, _publishDriveId);
        return (_publishSiteId, _publishDriveId);
    }

    /// <summary>Returns the SharePoint folder path for the given channel.</summary>
    private string ResolvePublishFolder(string? channel) =>
        string.Equals(channel, "dev", StringComparison.OrdinalIgnoreCase)
            ? "publish/dev"
            : (_config["Publish:FolderPath"] ?? "publish/current").Trim('/');

    /// <summary>Reads version.txt from the configured SharePoint publish folder via Graph.</summary>
    /// <param name="channel">"dev" ?????' publish/dev; anything else ?????' publish/current (or Publish:FolderPath config).</param>
    public async Task<string> GetPublishVersionAsync(string? channel = null)
    {
        var (_, driveId) = await GetPublishDriveAsync();
        var folderPath   = ResolvePublishFolder(channel);
        // Graph SDK v5: Drives[id].Items["root:/path/to/file:"].Content
        var itemKey = $"root:/{folderPath}/version.txt:";

        using var stream = await GetGraph().Drives[driveId].Items[itemKey].Content.GetAsync();
        if (stream is null) throw new Exception($"version.txt not found at '{folderPath}/version.txt'");

        using var reader = new StreamReader(stream);
        return (await reader.ReadToEndAsync()).Trim().Split('+')[0].Trim();
    }

    /// <summary>
    /// Downloads a file from the configured SharePoint publish folder via Graph.
    /// Returns (contentType, bytes, fileName).
    /// Throws if the file is not found or the name is not a simple filename (no path traversal).
    /// </summary>
    public async Task<(string ContentType, byte[] Bytes, string FileName)> GetPublishFileAsync(string fileName, string? channel = null)
    {
        // Guard against path traversal
        if (string.IsNullOrWhiteSpace(fileName) ||
            fileName.IndexOfAny(Path.GetInvalidFileNameChars()) >= 0 ||
            fileName.Contains('/') || fileName.Contains('\\'))
            throw new ArgumentException($"Invalid file name: '{fileName}'");

        var (_, driveId) = await GetPublishDriveAsync();
        var folderPath   = ResolvePublishFolder(channel);
        var itemKey      = $"root:/{folderPath}/{fileName}:";

        using var stream = await GetGraph().Drives[driveId].Items[itemKey].Content.GetAsync();
        if (stream is null) throw new Exception($"File '{fileName}' not found at '{folderPath}/{fileName}'");

        using var ms = new MemoryStream();
        await stream.CopyToAsync(ms);
        var bytes = ms.ToArray();

        var ext = Path.GetExtension(fileName).ToLowerInvariant();
        var contentType = ext switch
        {
            ".ps1"  => "text/plain",
            ".exe"  => "application/octet-stream",
            ".txt"  => "text/plain",
            ".json" => "application/json",
            ".bat"  => "text/plain",
            _       => "application/octet-stream",
        };

        _log.LogInformation("[Publish] Served '{File}' ({Bytes} bytes)", fileName, bytes.Length);
        return (contentType, bytes, fileName);
    }

    /// <summary>
    /// Streams all files in the SharePoint publish folder as a ZIP into
    /// <paramref name="destination"/>. Preserves the Proxy/ subdirectory.
    /// Writes to a temp file first so the central directory is fully finalised
    /// before any bytes reach the response stream.
    /// </summary>
    public async Task WritePublishPackageZipAsync(Stream destination, string? channel = null)
    {
        var (_, driveId) = await GetPublishDriveAsync();
        var folderPath   = ResolvePublishFolder(channel);

        var tempPath = Path.Combine(Path.GetTempPath(), $"ShredderPackage_{Guid.NewGuid():N}.zip");
        try
        {
            // Build complete ZIP on disk  -- Dispose() writes the central directory before we stream
            using (var tempFile = File.OpenWrite(tempPath))
            using (var zip = new ZipArchive(tempFile, ZipArchiveMode.Create, leaveOpen: false))
                await AddFolderToZipAsync(zip, driveId, folderPath, "");

            // Stream the finished file to the caller
            using var fs = File.OpenRead(tempPath);
            await fs.CopyToAsync(destination);
        }
        finally
        {
            if (File.Exists(tempPath)) File.Delete(tempPath);
        }
    }

    private async Task AddFolderToZipAsync(ZipArchive zip, string driveId, string spPath, string zipPrefix)
    {
        var itemKey = $"root:/{spPath}:";
        var items   = await GetGraph().Drives[driveId].Items[itemKey].Children.GetAsync();

        foreach (var item in items?.Value ?? [])
        {
            var name    = item.Name ?? "";
            var zipName = zipPrefix == "" ? name : $"{zipPrefix}/{name}";

            if (item.Folder is not null)
            {
                await AddFolderToZipAsync(zip, driveId, $"{spPath}/{name}", zipName);
            }
            else if (item.File is not null)
            {
                var fileKey    = $"root:/{spPath}/{name}:";
                var fileStream = await GetGraph().Drives[driveId].Items[fileKey].Content.GetAsync();
                if (fileStream is null) continue;
                var entry = zip.CreateEntry(zipName, CompressionLevel.NoCompression);
                using var entryStream = entry.Open();
                await fileStream.CopyToAsync(entryStream);
                _log.LogInformation("[Update] Packaged {Name}", zipName);
            }
        }
    }

    // ??"?????"??? Diagnostics ??"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"???

    public async Task<object> DiagnoseAsync()
    {
        var steps = new List<object>();
        try
        {
            steps.Add(new { step = "token", status = "trying" });
            var tenantId     = _config["SharePoint:TenantId"]     ?? throw new Exception("SharePoint:TenantId not set");
            var clientId     = _config["SharePoint:ClientId"]     ?? throw new Exception("SharePoint:ClientId not set");
            var clientSecret = _config["SharePoint:ClientSecret"] ?? throw new Exception("SharePoint:ClientSecret not set");
            var credential   = new Azure.Identity.ClientSecretCredential(tenantId, clientId, clientSecret);
            var tokenCtx     = new Azure.Core.TokenRequestContext(["https://graph.microsoft.com/.default"]);
            var token        = await credential.GetTokenAsync(tokenCtx);
            var jwtParts     = token.Token.Split('.');
            var claimsJson   = jwtParts.Length > 1
                ? System.Text.Encoding.UTF8.GetString(
                    Convert.FromBase64String(jwtParts[1].PadRight((jwtParts[1].Length + 3) & ~3, '=')))
                : "{}";
            using var claimsDoc = System.Text.Json.JsonDocument.Parse(claimsJson);
            var roles = claimsDoc.RootElement.TryGetProperty("roles", out var r) ? r.ToString() : "NONE";
            var aud   = claimsDoc.RootElement.TryGetProperty("aud",   out var a) ? a.ToString() : "?";
            var tid   = claimsDoc.RootElement.TryGetProperty("tid",   out var t) ? t.ToString() : "?";
            steps[^1] = new { step = "token", status = "ok", expiresOn = token.ExpiresOn, aud, tid, roles };

            var graph = GetGraph();

            steps.Add(new { step = "sites/root", status = "trying" });
            var root = await graph.Sites["root"].GetAsync();
            steps[^1] = new { step = "sites/root", status = "ok", siteId = root?.Id, webUrl = root?.WebUrl };

            var siteUrl = _config["SharePoint:SiteUrl"] ?? "https://metalsupermarkets-my.sharepoint.com/personal/angus_mithrilmetals_com";
            var uri     = new Uri(siteUrl);
            var siteKey = $"{uri.Host}:{uri.AbsolutePath}";
            steps.Add(new { step = $"sites/{siteKey}", status = "trying" });
            var site = await graph.Sites[siteKey].GetAsync();
            steps[^1] = new { step = $"sites/{siteKey}", status = "ok", siteId = site?.Id };

            foreach (var listName in new[] { "SupplierResponses", "SupplierLineItems", "RFQ References" })
            {
                steps.Add(new { step = "list lookup", status = "trying" });
                var lists = await graph.Sites[site!.Id!].Lists
                    .GetAsync(r => r.QueryParameters.Filter = $"displayName eq '{listName}'");
                var found = lists?.Value?.FirstOrDefault();
                steps[^1] = new { step = "list lookup",
                                  status = found != null ? "ok" : "not_found",
                                  listId = found?.Id, listName };
            }
        }
        catch (Microsoft.Graph.Models.ODataErrors.ODataError ex)
        {
            steps.Add(new { step = "error", code = ex.Error?.Code, message = ex.Error?.Message });
        }
        catch (Exception ex)
        {
            steps.Add(new { step = "error", message = ex.Message });
        }
        return new { steps };
    }

    // ??"?????"??? Provision new supplier lists (run once) ??"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"???

    public async Task<Dictionary<string, object>> EnsureSupplierListsAsync()
    {
        var siteId  = await GetSiteIdAsync();
        var results = new Dictionary<string, object>
        {
            ["SupplierResponses"]  = await EnsureListColumnsAsync(siteId, "SupplierResponses",
            [
                ("RFQ_ID",               "text"),
                ("SupplierName",         "text"),
                ("EmailFrom",            "text"),
                ("ReceivedAt",           "dateTime"),
                ("EmailSubject",         "text"),
                ("EmailBody",            "note"),
                ("ProcessedAt",          "dateTime"),
                ("ProcessingSource",     "text"),
                ("SourceFile",           "text"),
                ("MessageId",            "text"),
                ("QuoteReference",       "text"),
                ("DateOfQuote",          "dateTime"),
                ("EstimatedDeliveryDate","dateTime"),
                ("FreightTerms",         "text"),
                ("IsRegret",             "boolean"),
                ("ClaudeResponseLog",    "note"),
            ]),
            ["SupplierLineItems"] = await EnsureListColumnsAsync(siteId, "SupplierLineItems",
            [
                ("SupplierResponseId",       "text"),
                ("RFQ_ID",                   "text"),
                ("SupplierName",             "text"),
                ("SourceFile",               "text"),
                ("EmailFrom",                "text"),
                ("MessageId",                "text"),
                ("QuoteReference",           "text"),
                ("ProductName",              "text"),
                ("CatalogProductName",       "text"),
                ("ProductSearchKey",         "text"),
                ("UnitsRequested",           "number"),
                ("UnitsQuoted",              "number"),
                ("LengthPerUnit",            "number"),
                ("LengthUnit",               "text"),
                ("WeightPerUnit",            "number"),
                ("WeightUnit",               "text"),
                ("PricePerPound",            "number"),
                ("PricePerFoot",             "number"),
                ("PricePerPiece",            "number"),
                ("TotalPrice",               "number"),
                ("LeadTimeText",             "text"),
                ("Certifications",           "text"),
                ("SupplierProductComments",  "note"),
                ("IsRegret",                 "boolean"),
                ("IsSubstitute",             "boolean"),
                ("IsPurchased",              "boolean"),
                ("PurchaseRecordId",         "text"),
            ]),
            ["RFQ Line Items"] = await EnsureListColumnsAsync(siteId,
                _config["SharePoint:ListName"] ?? "RFQ Line Items",
            [
                ("IsPurchased", "boolean"),
                ("PoNumber",    "text"),
            ]),
            ["PurchaseOrders"] = await EnsureListColumnsAsync(siteId, "PurchaseOrders",
            [
                ("RFQ_ID",       "text"),
                ("SupplierName", "text"),
                ("PoNumber",     "text"),
                ("ReceivedAt",   "dateTime"),
                ("MessageId",    "text"),
                ("LineItems",    "note"),
                ("PdfUrl",       "text"),
            ]),
            ["SupplierConversations"] = await EnsureListColumnsAsync(siteId, "SupplierConversations",
            [
                ("RFQ_ID",             "text"),
                ("SupplierName",       "text"),
                ("SupplierResponseId", "text"),
                ("Direction",          "text"),
                ("MessageId",          "text"),
                ("InReplyTo",          "text"),
                ("SentAt",             "dateTime"),
                ("EmailSubject",       "text"),
                ("BodyText",           "note"),
                ("HasAttachments",     "boolean"),
                ("ExtractedPricing",   "boolean"),
            ]),
        };
        if (results.TryGetValue("PurchaseOrders", out var poMap) &&
            poMap is Dictionary<string, string> poDict &&
            poDict.TryGetValue("RFQ_ID", out _))
        {
            // Cache list ID so we don't resolve again immediately
            _poListId = null; // will be lazy-resolved on next use
        }

        // Index the hot filter columns so grid/conversation queries hit a SharePoint
        // column index instead of scanning the list (which is what the
        // "HonorNonIndexedQueriesWarningMayFailRandomly" hint falls back to).
        results["Indexes:SupplierConversations"] = await EnsureColumnIndexesAsync(
            siteId, "SupplierConversations", "RFQ_ID", "SupplierName");
        results["Indexes:SupplierResponses"] = await EnsureColumnIndexesAsync(
            siteId, "SupplierResponses",     "RFQ_ID", "SupplierName", "MessageId");
        results["Indexes:SupplierLineItems"] = await EnsureColumnIndexesAsync(
            siteId, "SupplierLineItems",     "RFQ_ID", "SupplierName", "MessageId");
        results["Indexes:PurchaseOrders"] = await EnsureColumnIndexesAsync(
            siteId, "PurchaseOrders",        "RFQ_ID", "MessageId");

        return results;
    }

    /// <summary>
    /// Ensures each named column on the list has its <c>Indexed</c> flag set. Idempotent  --  columns
    /// already indexed or missing from the list are returned with a status string and no PATCH is issued.
    /// SharePoint allows at most 20 indexed columns per list; we're well under that.
    /// </summary>
    private async Task<Dictionary<string, string>> EnsureColumnIndexesAsync(
        string siteId, string listName, params string[] columnNames)
    {
        var results = new Dictionary<string, string>();

        var lists = await GetGraph().Sites[siteId].Lists
            .GetAsync(r => r.QueryParameters.Filter = $"displayName eq '{listName}'");
        var listId = lists?.Value?.FirstOrDefault()?.Id;
        if (listId is null)
        {
            foreach (var n in columnNames) results[n] = "list not found";
            return results;
        }

        var cols = await GetGraph().Sites[siteId].Lists[listId].Columns
            .GetAsync(r => r.QueryParameters.Select = ["id", "name", "indexed"]);

        var byName = cols?.Value?
            .Where(c => c.Name is not null)
            .ToDictionary(c => c.Name!, c => c, StringComparer.OrdinalIgnoreCase)
            ?? new Dictionary<string, ColumnDefinition>(StringComparer.OrdinalIgnoreCase);

        foreach (var name in columnNames)
        {
            if (!byName.TryGetValue(name, out var col) || col.Id is null)
            {
                results[name] = "column not found";
                continue;
            }
            if (col.Indexed == true)
            {
                results[name] = "already indexed";
                continue;
            }

            try
            {
                await GetGraph().Sites[siteId].Lists[listId].Columns[col.Id]
                    .PatchAsync(new ColumnDefinition { Indexed = true });
                results[name] = "indexed";
                _log.LogInformation("[SP] Indexed column '{Col}' on '{List}'", name, listName);
            }
            catch (Exception ex)
            {
                results[name] = $"error: {ex.Message}";
                _log.LogWarning("[SP] Failed to index '{Col}' on '{List}': {Err}", name, listName, ex.Message);
            }
        }
        return results;
    }

    private async Task<Dictionary<string, string>> EnsureListColumnsAsync(
        string siteId, string listName, (string Name, string Type)[] columns)
    {
        // Create list if absent
        var lists = await GetGraph().Sites[siteId].Lists
            .GetAsync(r => r.QueryParameters.Filter = $"displayName eq '{listName}'");
        string listId;
        if (lists?.Value?.FirstOrDefault() is null)
        {
            var newList = await GetGraph().Sites[siteId].Lists.PostAsync(new List
            {
                DisplayName = listName,
                ListProp    = new ListInfo { Template = "genericList" },
            });
            listId = newList?.Id ?? throw new Exception($"Failed to create list '{listName}'");
            _log.LogInformation("[SP] Created list '{Name}' -> id: {Id}", listName, listId);
        }
        else
        {
            listId = lists.Value.First().Id!;
        }

        // Cache the newly-resolved IDs
        if (listName == "SupplierResponses") _srListId  = listId;
        if (listName == "SupplierLineItems") _sliListId = listId;

        var existing = await GetGraph().Sites[siteId].Lists[listId].Columns.GetAsync();
        var existingNames = existing?.Value?
            .Select(c => c.Name ?? "").ToHashSet(StringComparer.OrdinalIgnoreCase) ?? [];

        var results = new Dictionary<string, string>();
        foreach (var (name, type) in columns)
        {
            if (existingNames.Contains(name)) { results[name] = "exists"; continue; }
            try
            {
                var col = type switch
                {
                    "text"     => new ColumnDefinition { Name = name, Text     = new TextColumn() },
                    "number"   => new ColumnDefinition { Name = name, Number   = new NumberColumn() },
                    "dateTime" => new ColumnDefinition { Name = name, DateTime = new DateTimeColumn() },
                    "note"     => new ColumnDefinition { Name = name, Text     = new TextColumn { AllowMultipleLines = true, LinesForEditing = 6 } },
                    "boolean"  => new ColumnDefinition { Name = name, Boolean  = new BooleanColumn() },
                    _          => new ColumnDefinition { Name = name, Text     = new TextColumn() }
                };
                await GetGraph().Sites[siteId].Lists[listId].Columns.PostAsync(col);
                results[name] = "created";
                _log.LogInformation("[SP] Created column '{Name}' ({Type}) on '{List}'", name, type, listName);
            }
            catch (Exception ex)
            {
                results[name] = $"error: {ex.Message}";
                _log.LogWarning("[SP] Column '{Name}' on '{List}': {Err}", name, listName, ex.Message);
            }
        }
        return results;
    }

    // ??"?????"??? Legacy: provision old RFQ Line Items list (kept for recovery) ??"?????"?????"?????"?????"?????"?????"?????"???

    public async Task<Dictionary<string, string>> EnsureColumnsAsync()
    {
        var siteId  = await GetSiteIdAsync();
        var listName = _config["SharePoint:ListName"] ?? "RFQ Line Items";

        var lists = await GetGraph().Sites[siteId].Lists
            .GetAsync(r => r.QueryParameters.Filter = $"displayName eq '{listName}'");
        if (lists?.Value?.FirstOrDefault() is null)
        {
            var newList = await GetGraph().Sites[siteId].Lists.PostAsync(new List
            {
                DisplayName = listName,
                ListProp    = new ListInfo { Template = "genericList" },
            });
            _listId = newList?.Id ?? throw new Exception($"Failed to create list '{listName}'");
        }

        var listId = await GetListIdAsync();
        var results = new Dictionary<string, string>();
        var existing = await GetGraph().Sites[siteId].Lists[listId].Columns.GetAsync();
        var existingNames = existing?.Value?.Select(c => c.Name ?? "").ToHashSet(StringComparer.OrdinalIgnoreCase) ?? [];

        var columns = new (string Name, string Type)[]
        {
            ("JobReference",            "text"),
            ("EmailFrom",               "text"),
            ("ReceivedAt",              "dateTime"),
            ("ProcessedAt",             "dateTime"),
            ("ProcessingSource",        "text"),
            ("SourceFile",              "text"),
            ("SupplierName",            "text"),
            ("QuoteReference",          "text"),
            ("DateOfQuote",             "dateTime"),
            ("EstimatedDeliveryDate",   "dateTime"),
            ("ProductName",             "text"),
            ("UnitsRequested",          "number"),
            ("UnitsQuoted",             "number"),
            ("LengthPerUnit",           "number"),
            ("LengthUnit",              "text"),
            ("WeightPerUnit",           "number"),
            ("WeightUnit",              "text"),
            ("PricePerPound",           "number"),
            ("PricePerFoot",            "number"),
            ("PricePerPiece",           "number"),
            ("TotalPrice",              "number"),
            ("LeadTimeText",            "text"),
            ("Certifications",          "text"),
            ("FreightTerms",            "text"),
            ("SupplierProductComments", "note"),
            ("CatalogProductName",      "text"),
            ("ProductSearchKey",        "text"),
            ("EmailBody",               "note"),
        };

        foreach (var (name, type) in columns)
        {
            if (existingNames.Contains(name)) { results[name] = "exists"; continue; }
            try
            {
                var col = type switch
                {
                    "text"     => new ColumnDefinition { Name = name, Text     = new TextColumn() },
                    "number"   => new ColumnDefinition { Name = name, Number   = new NumberColumn() },
                    "dateTime" => new ColumnDefinition { Name = name, DateTime = new DateTimeColumn() },
                    "note"     => new ColumnDefinition { Name = name, Text     = new TextColumn { AllowMultipleLines = true, LinesForEditing = 6 } },
                    _          => new ColumnDefinition { Name = name, Text     = new TextColumn() }
                };
                await GetGraph().Sites[siteId].Lists[listId].Columns.PostAsync(col);
                results[name] = "created";
            }
            catch (Exception ex)
            {
                results[name] = $"error: {ex.Message}";
                _log.LogError(ex, "[SP] Failed to create column '{Column}' on list '{List}'", name, listName);
            }
        }
        return results;
    }

    // ??"?????"??? Legacy: read from old RFQ Line Items (kept for migration) ??"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"???

    public async Task<List<Dictionary<string, object?>>> ReadItemsAsync(int top = 500)
    {
        var siteId = await GetSiteIdAsync();
        var listId = await GetListIdAsync();

        var result = await GetGraph().Sites[siteId].Lists[listId].Items
            .GetAsync(req => { req.QueryParameters.Expand = ["fields"]; req.QueryParameters.Top = top; });

        static bool IsAppField(string key) =>
            !key.StartsWith('@') && !key.StartsWith('_') &&
            key is not ("LinkTitle" or "LinkTitleNoMenu" or "ContentType" or "Edit"
                     or "Attachments" or "ItemChildCount" or "FolderChildCount"
                     or "Modified" or "Created"
                     or "AuthorLookupId" or "EditorLookupId"
                     or "AppAuthorLookupId" or "AppEditorLookupId");

        return result?.Value?
            .Where(i => i.Fields?.AdditionalData is not null)
            .Select(i => i.Fields!.AdditionalData!
                .Where(kv => IsAppField(kv.Key))
                .ToDictionary(kv => kv.Key, kv => (object?)kv.Value))
            .ToList()
            ?? [];
    }

    // ??"?????"??? QC list ??"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"???

    /// <summary>
    /// Resolves and caches the QC SP site + list IDs, returning the list object.
    /// </summary>
    private async Task<(string SiteId, string ListId, Microsoft.Graph.Models.List List)> ResolveQcAsync()
    {
        var graph = GetGraph();

        if (_qcSiteId is null || _qcListId is null)
        {
            var siteUrl = _config["QC:SiteUrl"]
                ?? "https://metalsupermarkets.sharepoint.com/sites/hackensack";
            var uri     = new Uri(siteUrl);
            var siteKey = $"{uri.Host}:{uri.AbsolutePath}";

            var site = await graph.Sites[siteKey].GetAsync();
            if (site?.Id is null)
                throw new Exception($"[QC] Could not resolve site '{siteKey}'");

            var lists = await graph.Sites[site.Id].Lists
                .GetAsync(r => r.QueryParameters.Filter = "displayName eq 'QC'");

            var list = lists?.Value?.FirstOrDefault();
            if (list?.Id is null)
                throw new Exception("[QC] List 'QC' not found");

            _qcSiteId = site.Id;
            _qcListId = list.Id;
            return (_qcSiteId, _qcListId, list);
        }

        // Already cached  -- fetch list object for LastModifiedDateTime
        var cachedList = await graph.Sites[_qcSiteId].Lists[_qcListId].GetAsync();
        if (cachedList is null)
            throw new Exception("[QC] Could not retrieve cached QC list");

        return (_qcSiteId, _qcListId, cachedList);
    }

    /// <summary>
    /// Returns the last-modified UTC datetime of the QC SharePoint list.
    /// </summary>
    public async Task<DateTime?> GetQcLastModifiedAsync()
    {
        var (_, _, list) = await ResolveQcAsync();
        return list.LastModifiedDateTime?.UtcDateTime;
    }

    /// <summary>
    /// Reads the QC SharePoint list and returns normalised columns and rows.
    /// Targets the columns: Metal, Shape, QC, Title (returned as Notes).
    /// Multi-value choice fields (Metal, Shape) are joined with "; ".
    /// Site URL is read from config key QC:SiteUrl.
    /// </summary>
    public async Task<QcListResult> ReadQcListAsync()
    {
        var (siteId, listId, list) = await ResolveQcAsync();
        var graph = GetGraph();

        // ??"?????"??? Discover columns ??"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"???
        // Metal and Shape are multi-value lookup fields; QC Cut contains notes.
        var wantedDisplay = new[] { "Metal", "Shape", "Title", "QC", "QC Cut", "LQ", "LQ Count", "LQ Min", "LQ Max" };

        var colsResp = await graph.Sites[siteId].Lists[listId].Columns.GetAsync();

        var fields = (colsResp?.Value ?? [])
            .Where(c => !string.IsNullOrEmpty(c.Name)
                     && !string.IsNullOrEmpty(c.DisplayName)
                     && wantedDisplay.Contains(c.DisplayName, StringComparer.OrdinalIgnoreCase)
                     // Exclude SP's auto-generated LinkTitle / LinkTitleNoMenu columns
                     && !c.Name!.StartsWith("LinkTitle", StringComparison.OrdinalIgnoreCase))
            .Select(c => (Display: c.DisplayName!, Internal: c.Name!))
            .OrderBy(c => Array.FindIndex(wantedDisplay,
                w => w.Equals(c.Display, StringComparison.OrdinalIgnoreCase)))
            .ToArray();

        // ??"?????"??? Fetch items ??"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"???
        var selectFields = string.Join(",", fields.Select(f => f.Internal));
        var items = await graph.Sites[siteId].Lists[listId].Items
            .GetAsync(r =>
            {
                r.QueryParameters.Expand = [$"fields($select={selectFields})"];
                r.QueryParameters.Top    = 5000;
            });

        var itemData = (items?.Value ?? [])
            .Where(i => i.Fields?.AdditionalData is not null)
            .Select(i =>
            {
                var d = i.Fields!.AdditionalData!;
                var row = fields.Select(f =>
                    d.TryGetValue(f.Internal, out var v) ? SerializeQcValue(v) : ""
                ).ToArray();
                return (Id: i.Id ?? "", Row: row);
            })
            .ToArray();

        var rows    = itemData.Select(x => x.Row).ToArray();
        var itemIds = itemData.Select(x => x.Id).ToArray();

        // Map display names: "QC Cut" -> "Notes" for the client
        var outputColumns = fields
            .Select(f => f.Display.Equals("QC Cut", StringComparison.OrdinalIgnoreCase) ? "Notes" : f.Display)
            .ToArray();

        return new QcListResult(outputColumns, rows, itemIds, list.LastModifiedDateTime?.UtcDateTime);
    }

    // ??"?????"??? QC row update ??"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"???

    /// <summary>
    /// Patches the QC and QC Cut fields of a single QC list item by SP item ID.
    /// </summary>
    public async Task UpdateQcRowAsync(string itemId, string? qc, string? qcCut)
    {
        var (siteId, listId, _) = await ResolveQcAsync();
        var graph = GetGraph();

        var colsResp = await graph.Sites[siteId].Lists[listId].Columns.GetAsync();

        string? qcInternal    = null;
        string? qcCutInternal = null;

        foreach (var col in colsResp?.Value ?? [])
        {
            if (string.IsNullOrEmpty(col.Name) || string.IsNullOrEmpty(col.DisplayName)) continue;
            if (col.Name.StartsWith("LinkTitle", StringComparison.OrdinalIgnoreCase)) continue;

            if (col.DisplayName.Equals("QC", StringComparison.OrdinalIgnoreCase))
                qcInternal = col.Name;
            else if (col.DisplayName.Equals("QC Cut", StringComparison.OrdinalIgnoreCase))
                qcCutInternal = col.Name;
        }

        var patch = new Dictionary<string, object?>();
        if (qcInternal    is not null) patch[qcInternal]    = qc    ?? "";
        if (qcCutInternal is not null) patch[qcCutInternal] = qcCut ?? "";

        if (patch.Count == 0)
            throw new InvalidOperationException("[QC] Could not resolve QC or QC Cut column internal names");

        await graph.Sites[siteId].Lists[listId].Items[itemId].Fields
            .PatchAsync(new FieldValueSet { AdditionalData = patch });

        _log.LogInformation("[QC] Patched item {ItemId}: QC={Qc} QcCut={QcCut}", itemId, qc, qcCut);
    }

    // ??"?????"??? LQ update ??"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"???

    /// <summary>
    /// Joins supplier quotes ?????' RFQ Line Items (canonical product names) ?????' QC Metal+Shape rows,
    /// derives $/lb for each quote, patches the QC list 'LQ' column, and returns a match/miss log.
    ///
    /// Join chain:
    ///   SupplierLineItem.RFQ_ID ?????' RFQLineItem.RFQ_ID ?????' RFQLineItem.Product
    ///   RFQLineItem.Product (text-containment) ?????' QC row Metal+Shape
    /// </summary>
    public async Task<LqUpdateResult> UpdateQcLqAsync()
    {
        var (qcSiteId, qcListId, _) = await ResolveQcAsync();
        var graph = GetGraph();

        // ??"?????"??? Helper: extract number from an object? that may be JsonElement ??"?????"?????"?????"???
        static double? GetNum(Dictionary<string, object?> row, string key)
        {
            if (!row.TryGetValue(key, out var v) || v is null) return null;
            return v switch
            {
                double d => d,
                int    i => (double)i,
                JsonElement je when je.ValueKind == JsonValueKind.Number => je.GetDouble(),
                _ => double.TryParse(v.ToString(), out var d) ? d : null
            };
        }

        static bool IsRegret(Dictionary<string, object?> row)
        {
            if (!row.TryGetValue("IsRegret", out var v) || v is null) return false;
            return v switch
            {
                bool b => b,
                JsonElement je => je.ValueKind == JsonValueKind.True,
                _ => false
            };
        }

        static double ToPounds(double weight, string? unit) => (unit?.ToLowerInvariant()) switch
        {
            "kg" => weight * 2.20462,
            "oz" => weight / 16.0,
            "g"  => weight / 453.592,
            _    => weight
        };

        // Derive $/lb from a supplier quote row; returns null if not computable.
        static double? DerivePerPound(Dictionary<string, object?> row)
        {
            var ppp = GetNum(row, "PricePerPound");
            if (ppp is > 0) return ppp;

            var total  = GetNum(row, "TotalPrice");
            var qty    = GetNum(row, "UnitsQuoted") ?? GetNum(row, "UnitsRequested");
            var weight = GetNum(row, "WeightPerUnit");
            if (total is > 0 && qty is > 0 && weight is > 0)
            {
                var unit    = row.TryGetValue("WeightUnit", out var wu) ? wu?.ToString() : null;
                var totalLb = qty.Value * ToPounds(weight.Value, unit);
                if (totalLb > 0) return total.Value / totalLb;
            }

            return null;
        }

        // ??"?????"??? 1. Fetch QC rows with item IDs ??"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"???
        var wantedDisplay = new[] { "Metal", "Shape", "Title", "LQ", "LQ Count", "LQ Min", "LQ Max" };
        var colsResp = await graph.Sites[qcSiteId].Lists[qcListId].Columns.GetAsync();
        var fields = (colsResp?.Value ?? [])
            .Where(c => !string.IsNullOrEmpty(c.Name) && !string.IsNullOrEmpty(c.DisplayName)
                     && wantedDisplay.Contains(c.DisplayName, StringComparer.OrdinalIgnoreCase)
                     && !c.Name!.StartsWith("LinkTitle", StringComparison.OrdinalIgnoreCase))
            .Select(c => (Display: c.DisplayName!, Internal: c.Name!))
            .ToArray();

        string? ColInternal(string display) =>
            fields.FirstOrDefault(f => f.Display.Equals(display, StringComparison.OrdinalIgnoreCase)).Internal;

        var metalField = ColInternal("Metal") ?? throw new Exception("[QC] 'Metal' column not found");
        var shapeField = ColInternal("Shape") ?? throw new Exception("[QC] 'Shape' column not found");
        var titleField = ColInternal("Title") ?? "Title";
        var lqField    = ColInternal("LQ")    ?? throw new Exception("[QC] 'LQ' column not found  -- create it in the QC list first");

        // Auto-create 'LQ Count', 'LQ Min', 'LQ Max' number columns if missing
        async Task<string> EnsureNumberColumn(string display, string fallback)
        {
            var existing = ColInternal(display);
            if (existing is not null) return existing;
            _log.LogInformation("[LQ] '{Display}' column not found  -- creating it", display);
            var created = await graph.Sites[qcSiteId].Lists[qcListId].Columns
                .PostAsync(new ColumnDefinition
                {
                    Name        = fallback,
                    DisplayName = display,
                    Number      = new NumberColumn()
                });
            var name = created?.Name ?? fallback;
            _log.LogInformation("[LQ] Created '{Display}' column (internal: {Name})", display, name);
            return name;
        }

        var lqCountField = await EnsureNumberColumn("LQ Count", "LQ_Count");
        var lqMinField   = await EnsureNumberColumn("LQ Min",   "LQ_Min");
        var lqMaxField   = await EnsureNumberColumn("LQ Max",   "LQ_Max");

        var qcItemsResp = await graph.Sites[qcSiteId].Lists[qcListId].Items
            .GetAsync(r =>
            {
                r.QueryParameters.Expand = [$"fields($select={metalField},{shapeField},{titleField})"];
                r.QueryParameters.Top    = 5000;
            });

        var qcRows = (qcItemsResp?.Value ?? [])
            .Where(i => i.Id is not null && i.Fields?.AdditionalData is not null)
            .Select(i =>
            {
                var d       = i.Fields!.AdditionalData!;
                var metals  = (d.TryGetValue(metalField, out var mv) ? SerializeQcValue(mv) : "")
                              .Split(';').Select(m => m.Trim()).Where(m => m.Length > 0).ToArray();
                var shapes  = (d.TryGetValue(shapeField, out var sv) ? SerializeQcValue(sv) : "")
                              .Split(';').Select(s => s.Trim()).Where(s => s.Length > 0).ToArray();
                // Title contains pipe-separated extra match terms (e.g. "pipe|tube")
                var titles  = (d.TryGetValue(titleField, out var tv) ? SerializeQcValue(tv) : "")
                              .Split('|').Select(t => t.Trim().ToLowerInvariant()).Where(t => t.Length > 0).ToArray();
                return (Id: i.Id!, Metals: metals, Shapes: shapes, Titles: titles);
            })
            .Where(r => r.Metals.Length > 0 && r.Shapes.Length > 0)
            .ToArray();

        // ??"?????"??? 2. Fetch priced supplier quotes, group by RFQ_ID ??"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"???
        var lookbackDays = int.TryParse(_config["QC:LqLookbackDays"], out var ld) ? ld : 7;
        var cutoff       = DateTime.UtcNow.AddDays(-lookbackDays);

        var (allSli, _) = await ReadSupplierItemsAsync(top: 5000);

        // rfqId ?????' list of $/lb values from priced non-regret quotes within the lookback window
        var pricesByRfq = new Dictionary<string, List<double>>(StringComparer.OrdinalIgnoreCase);
        var unpricedCount = 0;
        var staleCount    = 0;

        foreach (var sli in allSli)
        {
            if (IsRegret(sli)) continue;

            // Filter by when the data was processed/written to SP (Modified), not when the
            // email arrived (ReceivedAt)  -- emails can sit in the inbox for days before
            // being processed, making ReceivedAt an unreliable freshness indicator.
            DateTime? quoteDate = null;
            foreach (var key in new[] { "Modified", "ProcessedAt", "DateOfQuote", "ReceivedAt" })
            {
                if (!sli.TryGetValue(key, out var dv) || dv is null) continue;
                var ds = dv is JsonElement je ? je.ToString() : dv.ToString();
                if (DateTime.TryParse(ds, null, System.Globalization.DateTimeStyles.RoundtripKind, out var dt))
                { quoteDate = dt.ToUniversalTime(); break; }
            }
            if (quoteDate.HasValue && quoteDate.Value < cutoff) { staleCount++; continue; }

            var rfqId = sli.TryGetValue("JobReference", out var jr) ? jr?.ToString()
                      : sli.TryGetValue("RFQ_ID",       out var ri) ? ri?.ToString()
                      : null;
            if (string.IsNullOrEmpty(rfqId)) continue;

            var ppp = DerivePerPound(sli);
            if (ppp is null or <= 0) { unpricedCount++; continue; }

            if (!pricesByRfq.TryGetValue(rfqId, out var list))
                pricesByRfq[rfqId] = list = [];
            list.Add(ppp.Value);
        }

        _log.LogInformation("[LQ] {QcRows} QC rows, {RfqCount} RFQs with prices in last {Days}d, {Unpriced} unpriced, {Stale} outside window",
            qcRows.Length, pricesByRfq.Count, lookbackDays, unpricedCount, staleCount);

        // ??"?????"??? 3. Fetch RFQ Line Items for RFQs that have quotes ??"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"???
        // rfqId ?????' canonical product names (lower-cased)
        var rfqProducts = new Dictionary<string, List<string>>(StringComparer.OrdinalIgnoreCase);
        var allRfqLines = await ReadAllRfqLineItemsAsync();

        foreach (var (rfqId, _, product, _, _, _, _) in allRfqLines)
        {
            if (string.IsNullOrEmpty(rfqId) || string.IsNullOrEmpty(product)) continue;
            if (!pricesByRfq.ContainsKey(rfqId)) continue;   // no quotes for this RFQ
            if (!rfqProducts.TryGetValue(rfqId, out var prods))
                rfqProducts[rfqId] = prods = [];
            prods.Add(product.ToLowerInvariant());
        }

        _log.LogInformation("[LQ] {Count} RFQ Line Item products across quoted RFQs", rfqProducts.Values.Sum(v => v.Count));

        // ??"?????"??? 4. Build: QC row ?????' list of prices whose RFQ products match ??"?????"?????"?????"?????"?????"?????"?????"???
        // For each QC row, find RFQs where any product contains Metal AND Shape,
        // then collect all prices from those RFQs.
        var updated = new List<LqMatch>();
        var misses  = new List<string>();

        foreach (var qcRow in qcRows)
        {
            var matchedPrices = new List<double>();
            var matchedRfqs   = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

            foreach (var (rfqId, products) in rfqProducts)
            {
                foreach (var product in products)
                {
                    bool metalMatch = qcRow.Metals.Any(m => product.Contains(m.ToLowerInvariant()));
                    bool shapeMatch = qcRow.Shapes.Any(s => product.Contains(s.ToLowerInvariant()));
                    bool titleMatch = qcRow.Titles.Length == 0 || qcRow.Titles.Any(t => product.Contains(t));
                    if (metalMatch && shapeMatch && titleMatch)
                    {
                        matchedRfqs.Add(rfqId);
                        matchedPrices.AddRange(pricesByRfq[rfqId]);
                        break; // one match per RFQ is enough
                    }
                }
            }

            var metalLabel = string.Join("; ", qcRow.Metals);
            var shapeLabel = string.Join("; ", qcRow.Shapes);

            if (matchedPrices.Count > 0)
            {
                var lq    = matchedPrices.Average();
                var lqMin = matchedPrices.Min();
                var lqMax = matchedPrices.Max();
                await graph.Sites[qcSiteId].Lists[qcListId].Items[qcRow.Id].Fields
                    .PatchAsync(new FieldValueSet
                    {
                        AdditionalData = new Dictionary<string, object?>
                        {
                            [lqField]      = lq,
                            [lqCountField] = (double)matchedPrices.Count,
                            [lqMinField]   = lqMin,
                            [lqMaxField]   = lqMax
                        }
                    });
                updated.Add(new LqMatch(metalLabel, shapeLabel, lq, matchedPrices.Count, lqMin, lqMax));
                _log.LogInformation("[LQ] Updated {Metal}/{Shape} LQ={Lq:F4} min={Min:F4} max={Max:F4} (n={N}, last {Days}d) RFQs: {Rfqs}",
                    metalLabel, shapeLabel, lq, lqMin, lqMax, matchedPrices.Count, lookbackDays, string.Join(", ", matchedRfqs));
            }
            else
            {
                misses.Add($"[QC ROW - NO QUOTES] {metalLabel} / {shapeLabel}");
            }
        }

        // ??"?????"??? 5. Log RFQ Line Item products that matched no QC row ??"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"???
        var matchedProducts = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        foreach (var qcRow in qcRows)
        foreach (var (_, products) in rfqProducts)
        foreach (var product in products)
        {
            if (qcRow.Metals.Any(m => product.Contains(m.ToLowerInvariant())) &&
                qcRow.Shapes.Any(s => product.Contains(s.ToLowerInvariant())) &&
                (qcRow.Titles.Length == 0 || qcRow.Titles.Any(t => product.Contains(t))))
                matchedProducts.Add(product);
        }

        foreach (var (rfqId, products) in rfqProducts)
        foreach (var product in products)
            if (!matchedProducts.Contains(product))
                misses.Add($"[RFQ PRODUCT - NO QC ROW] {rfqId}: {product}");

        return new LqUpdateResult(updated, misses);
    }

    /// <summary>
    /// Converts a SharePoint field value to a display string.
    /// Multi-value fields (UntypedArray) are joined with "; ".
    /// </summary>
    private static string SerializeQcValue(object? v)
    {
        return v switch
        {
            null                => "",
            UntypedArray arr    => string.Join("; ", arr.GetValue()
                                        .Select(ExtractNodeString)
                                        .Where(s => s.Length > 0)),
            UntypedString str   => str.GetValue() ?? "",
            UntypedNull         => "",
            UntypedDouble dbl   => dbl.GetValue().ToString(),
            UntypedInteger intV => intV.GetValue().ToString(),
            UntypedBoolean b    => b.GetValue().ToString(),
            _                   => v.ToString() ?? ""
        };
    }

    /// <summary>
    /// Extracts a display string from an UntypedNode array element.
    /// Lookup fields return UntypedObject with a "LookupValue" key.
    /// </summary>
    private static string ExtractNodeString(UntypedNode n)
    {
        if (n is UntypedString s) return s.GetValue() ?? "";
        if (n is UntypedObject obj)
        {
            var props = obj.GetValue();
            if (props.TryGetValue("LookupValue", out var lv) && lv is UntypedString ls)
                return ls.GetValue() ?? "";
        }
        return "";
    }

    // ??"?????"??? Read: Product Catalog ??"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"???

    /// <summary>
    /// Returns all rows from the "Product Catalog" SP list.
    /// Tries common internal-name variants for each field so the list column names
    /// don't need to exactly match hard-coded strings.
    /// </summary>
    public async Task<List<ProductCatalogDto>> ReadProductCatalogAsync()
    {
        var listName = _config["ProductCatalog:ListName"] ?? "Product Catalog";
        var siteId   = await GetSiteIdAsync();
        var listId   = await ResolveListIdAsync(listName);

        _log.LogInformation("[SP] ReadProductCatalog: list='{Name}'", listName);

        // Fetch catalog items + lookup tables for Category and Shape in parallel.
        var itemsTask = GetGraph().Sites[siteId].Lists[listId].Items
            .GetAsync(req =>
            {
                req.QueryParameters.Expand = ["fields"];
                req.QueryParameters.Top    = 5000;
            });

        // Category lookup: Metals list item-id ?????' category name
        var categoryMapTask = BuildLookupMapAsync(siteId, ["Metals", "Metal", "Product Categories"], "Title");

        // Shape lookup: Shapes list item-id ?????' shape name
        var shapeMapTask = BuildLookupMapAsync(siteId, ["Shapes", "Shape"], "ProductShape", "Title");

        await Task.WhenAll(itemsTask, categoryMapTask, shapeMapTask);

        var raw         = itemsTask.Result?.Value ?? [];
        var categoryMap = categoryMapTask.Result;
        var shapeMap    = shapeMapTask.Result;

        return raw
            .Where(i => i.Fields?.AdditionalData is not null)
            .Select(i =>
            {
                var d = i.Fields!.AdditionalData!;

                // Resolve lookup IDs ?????' text values via the maps built above.
                var catId   = RfqField(d, "Product_x0020_CategoryLookupId");
                var shapeId = RfqField(d, "Product_x0020_ShapeLookupId");
                var cat     = catId   is not null && categoryMap.TryGetValue(catId,   out var c) ? c : null;
                var shape   = shapeId is not null && shapeMap.TryGetValue(shapeId,    out var s) ? s : null;

                return new ProductCatalogDto
                {
                    // "Line_x0020_No" = "Line No" / MSPC equivalent
                    Mspc             = RfqField(d, "Line_x0020_No", "MSPC", "mspc", "Mspc"),
                    // "Product_x0020_SearchKey" = "Product SearchKey" column
                    ProductSearchKey = RfqField(d, "Product_x0020_SearchKey", "ProductSearchKey", "SearchKey"),
                    // "Product" is the internal field name for the product name column
                    ProductName      = RfqField(d, "Product", "ProductName", "Title"),
                    Category         = cat ?? RfqField(d, "Product_x0020_Category", "Metal", "ProductCategory", "Category"),
                    Shape            = shape ?? RfqField(d, "Product_x0020_Shape", "ProductShape", "Shape"),
                };
            })
            .Where(p => !string.IsNullOrWhiteSpace(p.ProductName))
            .ToList();
    }

    /// <summary>
    /// Reads items from the first matching list name and returns a map of item-ID ?????' text value
    /// (tries each field key in order; used to resolve SharePoint lookup column IDs).
    /// </summary>
    private async Task<Dictionary<string, string>> BuildLookupMapAsync(
        string siteId, string[] listNames, params string[] fieldKeys)
    {
        foreach (var name in listNames)
        {
            try
            {
                var lid   = await ResolveListIdAsync(name);
                var items = await GetGraph().Sites[siteId].Lists[lid].Items
                    .GetAsync(req =>
                    {
                        req.QueryParameters.Expand = ["fields"];
                        req.QueryParameters.Top    = 500;
                    });

                return (items?.Value ?? [])
                    .Where(i => i.Id is not null && i.Fields?.AdditionalData is not null)
                    .Select(i => (Id: i.Id!, Val: RfqField(i.Fields!.AdditionalData!, fieldKeys)))
                    .Where(x => x.Val is not null)
                    .ToDictionary(x => x.Id, x => x.Val!);
            }
            catch { /* try next list name */ }
        }

        return [];
    }

    // ??"?????"??? Read: Metal Categories ??"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"???

    /// <summary>
    /// Returns all ProductCategory values from the "Metals" SP list, sorted alphabetically.
    /// Falls back to Title if ProductCategory column is absent.
    /// </summary>
    public async Task<List<string>> ReadMetalCategoriesAsync()
    {
        var siteId = await GetSiteIdAsync();

        // Try several possible list names; fall through to product-catalog derivation if none found.
        string? listId = null;
        foreach (var name in new[] { "Metals", "Metal", "Product Categories", "ProductCategories" })
        {
            try { listId = await ResolveListIdAsync(name); break; }
            catch { /* list not found  -- try next */ }
        }

        if (listId is not null)
        {
            _log.LogInformation("[SP] ReadMetalCategories from list");

            var items = await GetGraph().Sites[siteId].Lists[listId].Items
                .GetAsync(req =>
                {
                    req.QueryParameters.Expand = ["fields"];
                    req.QueryParameters.Top    = 1000;
                });

            if (items?.Value?.Count > 0)
            {
                var sampleKeys = string.Join(", ", items.Value[0].Fields?.AdditionalData?.Keys ?? []);
                _log.LogInformation("[SP] Metals first-item fields: [{Keys}]", sampleKeys);
            }

            var vals = (items?.Value ?? [])
                .Where(i => i.Fields?.AdditionalData is not null)
                .Select(i => RfqField(i.Fields!.AdditionalData!, "Title", "ProductCategory", "Metal", "Category"))
                .Where(s => !string.IsNullOrWhiteSpace(s))
                .Select(s => s!)
                .Distinct(StringComparer.OrdinalIgnoreCase)
                .OrderBy(s => s, StringComparer.OrdinalIgnoreCase)
                .ToList();

            if (vals.Count > 0) return vals;
        }

        // Fallback: derive distinct Metal/Category values from the Product Catalog.
        _log.LogInformation("[SP] ReadMetalCategories  -- Metals list not found, deriving from Product Catalog");
        var catalogListName = _config["ProductCatalog:ListName"] ?? "Product Catalog";
        var catalogListId   = await ResolveListIdAsync(catalogListName);
        var catalogItems = await GetGraph().Sites[siteId].Lists[catalogListId].Items
            .GetAsync(req =>
            {
                req.QueryParameters.Expand = ["fields"];
                req.QueryParameters.Top    = 5000;
            });

        return (catalogItems?.Value ?? [])
            .Where(i => i.Fields?.AdditionalData is not null)
            .Select(i => RfqField(i.Fields!.AdditionalData!, "Metal", "ProductCategory", "Category"))
            .Where(s => !string.IsNullOrWhiteSpace(s))
            .Select(s => s!)
            .Distinct(StringComparer.OrdinalIgnoreCase)
            .OrderBy(s => s, StringComparer.OrdinalIgnoreCase)
            .ToList();
    }

    // ??"?????"??? Read: Product Shapes ??"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"???

    /// <summary>
    /// Returns all ProductShape values from the "Shapes" SP list, sorted alphabetically.
    /// Falls back to Title if ProductShape column is absent.
    /// </summary>
    public async Task<List<string>> ReadProductShapesAsync()
    {
        var siteId = await GetSiteIdAsync();
        var listId = await ResolveListIdAsync("Shapes");

        _log.LogInformation("[SP] ReadProductShapes");

        var items = await GetGraph().Sites[siteId].Lists[listId].Items
            .GetAsync(req =>
            {
                req.QueryParameters.Expand = ["fields"];
                req.QueryParameters.Top    = 1000;
            });

        return (items?.Value ?? [])
            .Where(i => i.Fields?.AdditionalData is not null)
            .Select(i => RfqField(i.Fields!.AdditionalData!, "ProductShape", "Title"))
            .Where(s => !string.IsNullOrWhiteSpace(s))
            .Select(s => s!)
            .Distinct(StringComparer.OrdinalIgnoreCase)
            .OrderBy(s => s, StringComparer.OrdinalIgnoreCase)
            .ToList();
    }

    // ??"?????"??? Read: Supplier Relationships ??"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"???

    /// <summary>
    /// Returns all rows from the "Supplier Relationships" SP list.
    /// Metal is the primary lookup key; Shape is secondary (null = match any shape for this metal).
    /// Rows missing an email or metal are excluded.
    /// </summary>
    public async Task<List<SupplierRelationshipDto>> ReadSupplierRelationshipsAsync()
    {
        var siteId = await GetSiteIdAsync();

        // ??"?????"??? Step 1: read Supplier Relationships rows (lookup IDs only) ??"?????"?????"?????"?????"?????"?????"?????"???
        var relListId = await ResolveListIdAsync("Supplier Relationships");
        _log.LogInformation("[SP] ReadSupplierRelationships");

        var relItems = await GetGraph().Sites[siteId].Lists[relListId].Items
            .GetAsync(req =>
            {
                req.QueryParameters.Expand = ["fields"];
                req.QueryParameters.Top    = 5000;
            });

        var relRaw = relItems?.Value ?? [];
        if (relRaw.Count == 0) return [];

        if (relRaw.Count > 0)
        {
            var sampleKeys = string.Join(", ", relRaw[0].Fields?.AdditionalData?.Keys ?? []);
            _log.LogInformation("[SP] SupplierRelationships first-item fields: [{Keys}]", sampleKeys);
        }

        // ??"?????"??? Step 2: build Suppliers map id?????'(name,email) ??"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"???
        var suppListId = await ResolveListIdAsync("Suppliers");
        var suppItems  = await GetGraph().Sites[siteId].Lists[suppListId].Items
            .GetAsync(req =>
            {
                req.QueryParameters.Expand = ["fields($select=Title,ContactEmail)"];
                req.QueryParameters.Top    = 1000;
            });

        var suppMap = (suppItems?.Value ?? [])
            .Where(i => i.Id is not null && i.Fields?.AdditionalData is not null)
            .ToDictionary(
                i => i.Id!,
                i => (
                    Name:  i.Fields!.AdditionalData!.TryGetValue("Title",        out var t) ? t?.ToString() ?? "" : "",
                    Email: i.Fields!.AdditionalData!.TryGetValue("ContactEmail", out var e) ? e?.ToString() ?? "" : ""
                )
            );

        // ??"?????"??? Step 3: build Metals map id?????'name ??"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"???
        string? metalListId = null;
        foreach (var n in new[] { "Metals", "Metal", "Product Categories", "ProductCategories" })
        {
            try { metalListId = await ResolveListIdAsync(n); break; }
            catch { /* try next */ }
        }

        var metalMap = new Dictionary<string, string>();
        if (metalListId is not null)
        {
            var metalItems = await GetGraph().Sites[siteId].Lists[metalListId].Items
                .GetAsync(req =>
                {
                    req.QueryParameters.Expand = ["fields($select=Title)"];
                    req.QueryParameters.Top    = 200;
                });

            metalMap = (metalItems?.Value ?? [])
                .Where(i => i.Id is not null && i.Fields?.AdditionalData is not null)
                .ToDictionary(
                    i => i.Id!,
                    i => i.Fields!.AdditionalData!.TryGetValue("Title", out var t) ? t?.ToString() ?? "" : ""
                );
        }

        // ??"?????"??? Step 4: build Shapes map id?????'name (optional  -- list may not exist) ??"???
        Dictionary<string, string> shapeMap = [];
        try
        {
            var shapeListId = await ResolveListIdAsync("Shapes");
            var shapeItems  = await GetGraph().Sites[siteId].Lists[shapeListId].Items
                .GetAsync(req =>
                {
                    req.QueryParameters.Expand = ["fields($select=Title,ProductShape)"];
                    req.QueryParameters.Top    = 200;
                });

            shapeMap = (shapeItems?.Value ?? [])
                .Where(i => i.Id is not null && i.Fields?.AdditionalData is not null)
                .ToDictionary(
                    i => i.Id!,
                    i =>
                    {
                        var d = i.Fields!.AdditionalData!;
                        return (d.TryGetValue("ProductShape", out var ps) ? ps?.ToString() : null)
                            ?? (d.TryGetValue("Title",        out var tt) ? tt?.ToString() : null)
                            ?? "";
                    }
                );
        }
        catch { /* Shapes list is optional */ }

        // ??"?????"??? Step 5: join ??"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"???
        return relRaw
            .Where(i => i.Fields?.AdditionalData is not null)
            .Select(i =>
            {
                var d = i.Fields!.AdditionalData!;

                // Lookup ID fields: "SKLookupId" ?????' supplier item ID, "MetalLookupId" ?????' metal item ID
                var suppId   = RfqField(d, "SKLookupId",    "SupplierLookupId");
                var metalId  = RfqField(d, "MetalLookupId", "Metal_x0020_CategoryLookupId");
                var shapeId  = RfqField(d, "ShapeLookupId", "Product_x0020_ShapeLookupId");

                var supp  = suppId  is not null && suppMap.TryGetValue(suppId,  out var s) ? s : default;
                var metal = metalId is not null && metalMap.TryGetValue(metalId, out var m) ? m : "";
                var shape = shapeId is not null && shapeMap.TryGetValue(shapeId, out var sh) ? sh : null;

                // Fallback: try direct text fields in case the list is configured differently
                if (string.IsNullOrEmpty(supp.Name))
                    supp = (Name: RfqField(d, "Title", "SupplierName", "Supplier") ?? "", Email: RfqField(d, "Email", "SupplierEmail", "ContactEmail") ?? "");
                if (string.IsNullOrEmpty(metal))
                    metal = RfqField(d, "Metal", "ProductCategory", "Category") ?? "";

                return new SupplierRelationshipDto
                {
                    SupplierName = supp.Name,
                    Email        = supp.Email,
                    Metal        = metal,
                    Shape        = string.IsNullOrEmpty(shape) ? null : shape,
                };
            })
            .Where(r => !string.IsNullOrWhiteSpace(r.Metal) && !string.IsNullOrWhiteSpace(r.Email))
            .ToList();
    }

    /// <summary>
    /// Tries each key in order; returns the first non-empty string value found in the field dictionary.
    /// </summary>
    private static string? RfqField(IDictionary<string, object?> d, params string[] keys)
    {
        foreach (var k in keys)
            if (d.TryGetValue(k, out var v) && v?.ToString() is string s && s.Length > 0)
                return s;
        return null;
    }

    // ??"?????"??? ShredderConfig list ??"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"???

    /// <summary>
    /// Returns the ShredderConfig list ID, creating the list and its columns
    /// (Value, Comments) on first call if they do not exist.
    /// </summary>
    private async Task<string> GetOrCreateShredderConfigListIdAsync()
    {
        if (_shredderConfigListId is not null) return _shredderConfigListId;

        var siteId = await GetSiteIdAsync();
        const string listName = "ShredderConfig";

        var lists = await GetGraph().Sites[siteId].Lists
            .GetAsync(r => r.QueryParameters.Filter = $"displayName eq '{listName}'");

        string listId;
        if (lists?.Value?.FirstOrDefault() is null)
        {
            _log.LogInformation("[SP] Creating ShredderConfig list --??");
            var newList = await GetGraph().Sites[siteId].Lists.PostAsync(new List
            {
                DisplayName = listName,
                ListProp    = new ListInfo { Template = "genericList" },
            });
            listId = newList?.Id ?? throw new Exception("Failed to create ShredderConfig list");
            _log.LogInformation("[SP] Created ShredderConfig list -> {Id}", listId);

            // Add Value and Comments columns (Title = Name already exists)
            foreach (var (col, typ) in new[] { ("Value", "text"), ("Comments", "note") })
            {
                var def = typ == "note"
                    ? new ColumnDefinition { Name = col, Text = new TextColumn { AllowMultipleLines = true } }
                    : new ColumnDefinition { Name = col, Text = new TextColumn() };
                await GetGraph().Sites[siteId].Lists[listId].Columns.PostAsync(def);
                _log.LogInformation("[SP] Created ShredderConfig column '{Col}'", col);
            }
        }
        else
        {
            listId = lists.Value.First().Id!;
        }

        // Ensure Value and Comments columns exist (idempotent  -- skips already-present columns).
        var existing = await GetGraph().Sites[siteId].Lists[listId].Columns.GetAsync();
        var existingNames = existing?.Value?
            .Select(c => c.Name ?? "").ToHashSet(StringComparer.OrdinalIgnoreCase) ?? [];

        foreach (var (col, isNote) in new[] { ("Value", false), ("Comments", true) })
        {
            if (existingNames.Contains(col)) continue;
            var def = isNote
                ? new ColumnDefinition { Name = col, Text = new TextColumn { AllowMultipleLines = true } }
                : new ColumnDefinition { Name = col, Text = new TextColumn() };
            await GetGraph().Sites[siteId].Lists[listId].Columns.PostAsync(def);
            _log.LogInformation("[SP] Created ShredderConfig column '{Col}'", col);
        }

        _shredderConfigListId = listId;
        return listId;
    }

    /// <summary>
    /// Returns the Value field for the ShredderConfig row with the given name,
    /// or null if the row does not exist or the list does not exist.
    /// </summary>
    public async Task<(string? Value, string? Comments)?> GetShredderConfigAsync(string name)
    {
        try
        {
            var siteId = await GetSiteIdAsync();
            var listId = await GetOrCreateShredderConfigListIdAsync();

            // Fetch all rows (config list is always tiny) and filter client-side.
            // SP OData filter on Title fails unless the column is indexed.
            var items = await GetGraph().Sites[siteId].Lists[listId].Items
                .GetAsync(req => { req.QueryParameters.Expand = ["fields"]; req.QueryParameters.Top = 100; });

            var item = items?.Value?
                .FirstOrDefault(i => string.Equals(
                    i.Fields?.AdditionalData.TryGetValue("Title", out var t) == true ? t?.ToString() : null,
                    name, StringComparison.OrdinalIgnoreCase));
            if (item?.Fields?.AdditionalData is null) return null;

            var d        = item.Fields.AdditionalData;
            var value    = d.TryGetValue("Value",    out var v) ? v?.ToString() : null;
            var comments = d.TryGetValue("Comments", out var c) ? c?.ToString() : null;
            return (value, comments);
        }
        catch (Exception ex)
        {
            _log.LogWarning(ex, "[SP] GetShredderConfig('{Name}') failed", name);
            return null;
        }
    }

    /// <summary>
    /// Creates or updates the ShredderConfig row with the given name (upsert by Title).
    /// </summary>
    public async Task UpsertShredderConfigAsync(string name, string value, string comments)
    {
        var siteId = await GetSiteIdAsync();
        var listId = await GetOrCreateShredderConfigListIdAsync();

        // Fetch all rows and filter client-side (avoids unindexed-column OData filter error).
        var items = await GetGraph().Sites[siteId].Lists[listId].Items
            .GetAsync(req => { req.QueryParameters.Expand = ["fields"]; req.QueryParameters.Top = 100; });

        var existing = items?.Value?
            .FirstOrDefault(i => string.Equals(
                i.Fields?.AdditionalData.TryGetValue("Title", out var t) == true ? t?.ToString() : null,
                name, StringComparison.OrdinalIgnoreCase));
        if (existing is null)
        {
            await GetGraph().Sites[siteId].Lists[listId].Items
                .PostAsync(new ListItem
                {
                    Fields = new FieldValueSet
                    {
                        AdditionalData = new Dictionary<string, object?>
                        {
                            ["Title"]    = name,
                            ["Value"]    = value,
                            ["Comments"] = comments,
                        }
                    }
                });
            _log.LogInformation("[SP] Created ShredderConfig '{Name}' = '{Value}'", name, value);
        }
        else
        {
            await GetGraph().Sites[siteId].Lists[listId].Items[existing.Id!].Fields
                .PatchAsync(new FieldValueSet
                {
                    AdditionalData = new Dictionary<string, object?>
                    {
                        ["Value"]    = value,
                        ["Comments"] = comments,
                    }
                });
            _log.LogInformation("[SP] Updated ShredderConfig '{Name}' = '{Value}'", name, value);
        }
    }

    // ??"?????"??? PurchaseOrders list ??"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"???

    private async Task<string> GetPurchaseOrdersListIdAsync()
    {
        if (_poListId is not null) return _poListId;
        _poListId = await ResolveListIdAsync("PurchaseOrders");
        return _poListId;
    }

    /// <summary>
    /// Writes a purchase order row to the PurchaseOrders SharePoint list.
    /// Skips the write if a row with the same MessageId already exists (dedup).
    /// Returns the new SP item ID if written, null if skipped as a duplicate.
    /// </summary>
    public async Task<string?> WritePurchaseOrderAsync(
        string rfqId, string supplierName, string? poNumber,
        string receivedAt, string? messageId, string lineItemsJson)
    {
        var siteId  = await GetSiteIdAsync();
        var listId  = await GetPurchaseOrdersListIdAsync();

        // Dedup: skip if this exact email was already processed
        if (!string.IsNullOrEmpty(messageId))
        {
            var existing = await GetGraph().Sites[siteId].Lists[listId].Items
                .GetAsync(req =>
                {
                    req.QueryParameters.Expand = ["fields($select=MessageId)"];
                    req.QueryParameters.Top    = 1;
                    req.QueryParameters.Filter = $"fields/MessageId eq '{messageId.Replace("'", "''")}'";
                });
            if (existing?.Value?.Count > 0)
            {
                _log.LogInformation("[PO] Skipping duplicate PO  -- MessageId already in SP: {Id}", messageId);
                return null;
            }
        }

        var title = $"[{rfqId}] {supplierName} PO";
        var data  = new Dictionary<string, object?>
        {
            ["Title"]        = title[..Math.Min(title.Length, 255)],
            ["RFQ_ID"]       = rfqId,
            ["SupplierName"] = supplierName,
            ["PoNumber"]     = poNumber,
            ["ReceivedAt"]   = receivedAt,
            ["MessageId"]    = messageId,
            ["LineItems"]    = lineItemsJson,
        };

        var item = await GetGraph().Sites[siteId].Lists[listId].Items
            .PostAsync(new ListItem { Fields = new FieldValueSet { AdditionalData = data } });

        _log.LogInformation("[PO] Wrote PurchaseOrder to SP: [{RfqId}] {Supplier}", rfqId, supplierName);
        return item?.Id;
    }

    /// <summary>
    /// Patches the LineItems JSON field on an existing PurchaseOrders row.
    /// </summary>
    public async Task UpdatePurchaseOrderLineItemsAsync(string spItemId, string lineItemsJson)
    {
        var siteId = await GetSiteIdAsync();
        var listId = await GetPurchaseOrdersListIdAsync();

        await GetGraph().Sites[siteId].Lists[listId].Items[spItemId].Fields
            .PatchAsync(new FieldValueSet
            {
                AdditionalData = new Dictionary<string, object?> { ["LineItems"] = lineItemsJson }
            });
    }

    /// <summary>
    /// Deletes all rows from the PurchaseOrders SharePoint list.
    /// </summary>
    public async Task<int> CleanPurchaseOrdersAsync()
    {
        var siteId = await GetSiteIdAsync();
        var listId = await GetPurchaseOrdersListIdAsync();
        return await DeleteAllItemsAsync(siteId, listId, "PurchaseOrders");
    }

    /// <summary>
    /// Returns all rows from the PurchaseOrders SharePoint list.
    /// </summary>
    public async Task<List<PurchaseOrderRecord>> ReadPurchaseOrdersAsync()
    {
        var siteId = await GetSiteIdAsync();
        var listId = await GetPurchaseOrdersListIdAsync();

        var results = new List<PurchaseOrderRecord>();
        var page    = await GetGraph().Sites[siteId].Lists[listId].Items
            .GetAsync(req =>
            {
                req.QueryParameters.Expand = ["fields($select=RFQ_ID,SupplierName,PoNumber,ReceivedAt,MessageId,LineItems,PdfUrl)"];
                req.QueryParameters.Top    = 5000;
            });

        while (page?.Value is not null)
        {
            foreach (var item in page.Value)
            {
                var f = item.Fields?.AdditionalData;
                if (f is null) continue;
                var rfq      = GetStr(f, "RFQ_ID") ?? GetStr(f, "RFQ_x005F_ID") ?? "";
                var supplier = GetStr(f, "SupplierName") ?? "";
                if (string.IsNullOrEmpty(rfq) || string.IsNullOrEmpty(supplier)) continue;

                results.Add(new PurchaseOrderRecord
                {
                    SpItemId     = item.Id ?? "",
                    RfqId        = rfq,
                    SupplierName = supplier,
                    PoNumber     = GetStr(f, "PoNumber"),
                    ReceivedAt   = GetStr(f, "ReceivedAt"),
                    MessageId    = GetStr(f, "MessageId"),
                    LineItems    = GetStr(f, "LineItems") ?? "[]",
                    PdfUrl       = GetStr(f, "PdfUrl"),
                });
            }

            if (page.OdataNextLink is null) break;
            page = await GetGraph().Sites[siteId].Lists[listId].Items
                .WithUrl(page.OdataNextLink)
                .GetAsync();
        }

        return results;
    }

    // ?????? Supplier conversations ?????????????????????????????????????????????????????????????????????????????????????????????????????????????????????????????????????????????

    private string? _conversationsListId;

    private async Task<string> GetConversationsListIdAsync()
    {
        if (_conversationsListId is not null) return _conversationsListId;

        var siteId = await GetSiteIdAsync();
        var lists  = await GetGraph().Sites[siteId].Lists
            .GetAsync(r => r.QueryParameters.Filter = "displayName eq 'SupplierConversations'");

        var listId = lists?.Value?.FirstOrDefault()?.Id
            ?? throw new InvalidOperationException(
                "SupplierConversations list not found. Run POST /api/setup-supplier-lists once.");

        _conversationsListId = listId;
        return listId;
    }

    /// <summary>
    /// Appends one message (inbound or outbound) to the SupplierConversations list.
    /// Dedupes on MessageId when provided so re-runs and the mail poller don't duplicate.
    /// If a duplicate exists and msg.ExtractedPricing is true, patches ExtractedPricing on
    /// the existing row (so a SHR-routed row written before extraction can be promoted).
    /// Returns the new SP item ID, or null if a duplicate was skipped.
    /// </summary>
    public async Task<string?> WriteConversationMessageAsync(Models.ConversationMessage msg)
    {
        var siteId = await GetSiteIdAsync();
        var listId = await GetConversationsListIdAsync();

        if (!string.IsNullOrEmpty(msg.MessageId))
        {
            var existing = await GetGraph().Sites[siteId].Lists[listId].Items
                .GetAsync(req =>
                {
                    req.QueryParameters.Expand = ["fields($select=MessageId,ExtractedPricing)"];
                    req.QueryParameters.Top    = 1;
                    req.QueryParameters.Filter = $"fields/MessageId eq '{msg.MessageId.Replace("'", "''")}'";

                });
            if (existing?.Value?.Count > 0)
            {
                var existingItem = existing.Value[0];
                // If this write carries ExtractedPricing=true (post-extraction hook) and the
                // existing row was written before extraction ran (ExtractedPricing=false), patch it.
                if (msg.ExtractedPricing &&
                    existingItem.Fields?.AdditionalData.TryGetValue("ExtractedPricing", out var ep) == true &&
                    ep is false or null)
                {
                    await GetGraph().Sites[siteId].Lists[listId].Items[existingItem.Id].Fields
                        .PatchAsync(new FieldValueSet
                        {
                            AdditionalData = new Dictionary<string, object?> { ["ExtractedPricing"] = true }
                        });
                    _log.LogDebug("[Conv] Patched ExtractedPricing=true on existing row for MessageId {Id}", msg.MessageId);
                }
                else
                {
                    _log.LogDebug("[Conv] Skipping duplicate message  --  MessageId already in SP: {Id}", msg.MessageId);
                }
                return null;
            }
        }

        var title = $"[{msg.RfqId}] {msg.SupplierName} {msg.Direction}";
        var data  = new Dictionary<string, object?>
        {
            ["Title"]              = title[..Math.Min(title.Length, 255)],
            ["RFQ_ID"]             = msg.RfqId,
            ["SupplierName"]       = msg.SupplierName,
            ["SupplierResponseId"] = msg.SupplierResponseId,
            ["Direction"]          = msg.Direction,
            ["MessageId"]          = msg.MessageId,
            ["InReplyTo"]          = msg.InReplyTo,
            ["SentAt"]             = msg.SentAt.ToString("o"),
            ["EmailSubject"]       = msg.Subject,
            ["BodyText"]           = msg.BodyText,
            ["HasAttachments"]     = msg.HasAttachments,
            ["ExtractedPricing"]   = msg.ExtractedPricing,
        };

        var item = await GetGraph().Sites[siteId].Lists[listId].Items
            .PostAsync(new ListItem { Fields = new FieldValueSet { AdditionalData = data } });

        _log.LogInformation("[Conv] Wrote {Direction} message for [{RfqId}] {Supplier}",
            msg.Direction, msg.RfqId, msg.SupplierName);
        return item?.Id;
    }

    /// <summary>
    /// Returns all conversation messages for a given RFQ + supplier, merged from
    /// SupplierConversations (outbound + any logged inbound follow-ups) and
    /// SupplierResponses (inbound priced responses), ordered by SentAt ascending.
    /// The two underlying SharePoint list queries run in parallel.
    /// </summary>
    public async Task<List<Models.ConversationMessage>> ReadConversationAsync(
        string rfqId, string supplierName)
    {
        // Resolve site/list IDs first (these are typically cached after first resolve).
        var siteId     = await GetSiteIdAsync();
        var convListId = await GetConversationsListIdAsync();
        var srListId   = await GetSupplierResponsesListIdAsync();

        var rfq      = rfqId.Replace("'", "''");
        var supplier = supplierName.Replace("'", "''");
        var filter   = $"fields/RFQ_ID eq '{rfq}' and fields/SupplierName eq '{supplier}'";

        // Fire both list-item queries concurrently  --  they're independent.
        var convTask = ReadConversationRowsAsync(siteId, convListId, filter, rfqId, supplierName);
        var srTask   = ReadInboundSrRowsAsync(siteId, srListId, filter, rfqId, supplierName);

        await Task.WhenAll(convTask, srTask);

        var convRows = convTask.Result;
        var srRows   = srTask.Result;

        // Dedup SR rows against any convRows that already captured the same MessageId
        // (e.g. if the Phase 3 poller hook has logged the inbound already).
        if (convRows.Count > 0)
        {
            var seen = convRows
                .Where(r => !string.IsNullOrEmpty(r.MessageId))
                .Select(r => r.MessageId!)
                .ToHashSet(StringComparer.OrdinalIgnoreCase);
            srRows = srRows
                .Where(r => string.IsNullOrEmpty(r.MessageId) || !seen.Contains(r.MessageId!))
                .ToList();
        }

        return convRows.Concat(srRows).OrderBy(r => r.SentAt).ToList();
    }

    /// <summary>
    /// Reads only the outbound conversation rows (and any inbound follow-ups logged by
    /// the poller) from the SupplierConversations list  --  skips the slow SupplierResponses
    /// scan. Use when the caller already has the SR-derived inbound messages in memory.
    /// </summary>
    public async Task<List<Models.ConversationMessage>> ReadOutboundConversationAsync(
        string rfqId, string supplierName)
    {
        var siteId     = await GetSiteIdAsync();
        var convListId = await GetConversationsListIdAsync();
        var rfq        = rfqId.Replace("'", "''");
        var supplier   = supplierName.Replace("'", "''");
        var filter     = $"fields/RFQ_ID eq '{rfq}' and fields/SupplierName eq '{supplier}'";
        var rows       = await ReadConversationRowsAsync(siteId, convListId, filter, rfqId, supplierName);
        return rows.OrderBy(r => r.SentAt).ToList();
    }

    private async Task<List<Models.ConversationMessage>> ReadConversationRowsAsync(
        string siteId, string convListId, string filter, string rfqId, string supplierName)
    {
        var results = new List<Models.ConversationMessage>();
        var page = await GetGraph().Sites[siteId].Lists[convListId].Items
            .GetAsync(req =>
            {
                req.QueryParameters.Expand = ["fields($select=RFQ_ID,SupplierName,SupplierResponseId,Direction,MessageId,InReplyTo,SentAt,EmailSubject,BodyText,HasAttachments,ExtractedPricing)"];
                req.QueryParameters.Top    = 500;
                req.QueryParameters.Filter = filter;
            });

        while (page?.Value is not null)
        {
            foreach (var item in page.Value)
            {
                var f = item.Fields?.AdditionalData;
                if (f is null) continue;
                results.Add(new Models.ConversationMessage
                {
                    SpItemId           = item.Id,
                    RfqId              = GetStr(f, "RFQ_ID") ?? GetStr(f, "RFQ_x005F_ID") ?? rfqId,
                    SupplierName       = GetStr(f, "SupplierName") ?? supplierName,
                    SupplierResponseId = GetStr(f, "SupplierResponseId"),
                    Direction          = GetStr(f, "Direction") ?? "out",
                    MessageId          = GetStr(f, "MessageId"),
                    InReplyTo          = GetStr(f, "InReplyTo"),
                    SentAt             = DateTimeOffset.TryParse(GetStr(f, "SentAt"), out var dt) ? dt : default,
                    Subject            = GetStr(f, "EmailSubject"),
                    BodyText           = GetStr(f, "BodyText"),
                    HasAttachments     = f.TryGetValue("HasAttachments", out var ha) && ha is bool hab && hab,
                    ExtractedPricing   = f.TryGetValue("ExtractedPricing", out var ep) && ep is bool epb && epb,
                });
            }
            if (page.OdataNextLink is null) break;
            page = await GetGraph().Sites[siteId].Lists[convListId].Items
                .WithUrl(page.OdataNextLink).GetAsync();
        }

        return results;
    }

    private async Task<List<Models.ConversationMessage>> ReadInboundSrRowsAsync(
        string siteId, string srListId, string filter, string rfqId, string supplierName)
    {
        var results = new List<Models.ConversationMessage>();
        var page = await GetGraph().Sites[siteId].Lists[srListId].Items
            .GetAsync(req =>
            {
                req.QueryParameters.Expand = ["fields($select=RFQ_ID,SupplierName,EmailFrom,ContactEmail,ReceivedAt,EmailSubject,EmailBody,MessageId,SourceFile)"];
                req.QueryParameters.Top    = 200;
                req.QueryParameters.Filter = filter;
            });

        while (page?.Value is not null)
        {
            foreach (var item in page.Value)
            {
                var f = item.Fields?.AdditionalData;
                if (f is null) continue;
                results.Add(new Models.ConversationMessage
                {
                    SpItemId           = item.Id,
                    RfqId              = GetStr(f, "RFQ_ID") ?? GetStr(f, "RFQ_x005F_ID") ?? rfqId,
                    SupplierName       = GetStr(f, "SupplierName") ?? supplierName,
                    SupplierResponseId = item.Id,
                    Direction          = "in",
                    MessageId          = GetStr(f, "MessageId"),
                    SentAt             = DateTimeOffset.TryParse(GetStr(f, "ReceivedAt"), out var dt) ? dt : default,
                    Subject            = GetStr(f, "EmailSubject"),
                    BodyText           = GetStr(f, "EmailBody"),
                    HasAttachments     = !string.IsNullOrEmpty(GetStr(f, "SourceFile")),
                    ExtractedPricing   = true,
                    ContactEmail       = GetStr(f, "ContactEmail"),
                });
            }
            if (page.OdataNextLink is null) break;
            page = await GetGraph().Sites[siteId].Lists[srListId].Items
                .WithUrl(page.OdataNextLink).GetAsync();
        }

        return results;
    }

    /// <summary>
    /// Uploads a PO PDF to the SharePoint site's default drive under PurchaseOrders/{poNumber}.pdf,
    /// then patches the PurchaseOrders list item with the resulting web URL.
    /// Returns the web URL, or null if the upload fails.
    /// </summary>
    public async Task<string?> UploadPoAttachmentAsync(
        string spItemId, string poNumber, string fileName, byte[] pdfBytes)
    {
        try
        {
            var siteId = await GetSiteIdAsync();
            // Sanitise filename  -- keep PO number as the primary name
            var safeName = $"{poNumber}.pdf";

            // Resolve the default drive for the site
            var drive   = await GetGraph().Sites[siteId].Drive.GetAsync();
            var driveId = drive?.Id ?? throw new Exception("Could not resolve site drive ID");

            // Upload to the default document library under a PurchaseOrders subfolder
            // Graph SDK v5 path-based item key: "root:/folder/file:"
            var itemKey   = $"root:/PurchaseOrders/{safeName}:";
            var driveItem = await GetGraph().Drives[driveId].Items[itemKey].Content
                .PutAsync(new MemoryStream(pdfBytes));

            var webUrl = driveItem?.WebUrl;
            if (string.IsNullOrEmpty(webUrl))
            {
                _log.LogWarning("[PO] Upload succeeded but no WebUrl returned for {PoNumber}", poNumber);
                return null;
            }

            // Patch the PO list item with the URL
            var listId = await GetPurchaseOrdersListIdAsync();
            await GetGraph().Sites[siteId].Lists[listId].Items[spItemId].Fields
                .PatchAsync(new FieldValueSet
                {
                    AdditionalData = new Dictionary<string, object?> { ["PdfUrl"] = webUrl }
                });

            _log.LogInformation("[PO] PDF uploaded for {PoNumber} ?????' {Url}", poNumber, webUrl);
            return webUrl;
        }
        catch (Exception ex)
        {
            _log.LogWarning(ex, "[PO] Failed to upload PDF for {PoNumber}  -- continuing without attachment", poNumber);
            return null;
        }
    }

    // ??"?????"??? RLI purchase status update ??"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"???

    /// <summary>
    /// Returns true if the RFQ Reference for <paramref name="rfqId"/> has Complete = true.
    /// Fetches all RFQ References and filters in memory (OData filter on non-indexed columns
    /// is unreliable in SharePoint).
    /// </summary>
    private async Task<bool> IsRfqCompleteAsync(string siteId, string rfqId)
    {
        var listId = await GetRfqReferencesListIdAsync();
        var col    = await ResolveRfqIdColumnAsync(siteId, listId);

        var all = await GetGraph().Sites[siteId].Lists[listId].Items
            .GetAsync(req =>
            {
                req.QueryParameters.Expand = [$"fields($select={col},Complete)"];
                req.QueryParameters.Top    = 500;
            });

        return (all?.Value ?? []).Any(i =>
        {
            var d = i.Fields?.AdditionalData;
            if (d is null) return false;
            var itemRfqId = d.TryGetValue(col, out var v) ? v?.ToString()
                          : d.TryGetValue("RFQ_x005F_ID", out var v2) ? v2?.ToString() : null;
            if (!string.Equals(itemRfqId, rfqId, StringComparison.OrdinalIgnoreCase)) return false;
            return d.TryGetValue("Complete", out var c) &&
                   c is true or JsonElement { ValueKind: JsonValueKind.True };
        });
    }

    /// <summary>
    /// Marks SupplierLineItems as purchased by patching <c>IsPurchased = true</c> and
    /// <c>PurchaseRecordId</c> on SLI rows that match the given PO's rfqId + supplierName +
    /// MSPC (ProductSearchKey).  If the PO has no MSPC data all matching SLI rows are marked.
    /// Skips the update entirely when the parent RFQ Reference has <c>Complete = true</c>.
    /// </summary>
    public async Task UpdateRliPurchaseStatusAsync(
        string rfqId, string supplierName, string poSpItemId, List<PoLineItem> poLineItems,
        string? poNumber = null)
    {
        var siteId = await GetSiteIdAsync();

        if (await IsRfqCompleteAsync(siteId, rfqId))
        {
            _log.LogInformation("[PO] Skipping RLI update for [{RfqId}]  -- RFQ is marked Complete", rfqId);
            return;
        }

        var sliListId = await GetSupplierLineItemsListIdAsync();

        // Build MSPC set from PO line items; empty means no MSPC data ?????' mark all supplier rows
        var mspcSet = poLineItems
            .Select(li => li.Mspc?.Trim())
            .Where(m => !string.IsNullOrEmpty(m))
            .ToHashSet(StringComparer.OrdinalIgnoreCase);
        bool matchOnMspc = mspcSet.Count > 0;

        int patched = 0;

        var page = await GetGraph().Sites[siteId].Lists[sliListId].Items
            .GetAsync(req =>
            {
                req.QueryParameters.Expand = ["fields($select=id,RFQ_ID,SupplierName,ProductSearchKey,IsPurchased)"];
                req.QueryParameters.Top    = 2000;
            });

        while (page?.Value is not null)
        {
            foreach (var item in page.Value)
            {
                if (item.Id is null || item.Fields?.AdditionalData is not { } d) continue;

                var itemRfqId    = GetStr(d, "RFQ_ID") ?? GetStr(d, "RFQ_x005F_ID") ?? "";
                var itemSupplier = GetStr(d, "SupplierName") ?? "";

                if (!string.Equals(itemRfqId,    rfqId,        StringComparison.OrdinalIgnoreCase)) continue;
                if (!string.Equals(itemSupplier, supplierName, StringComparison.OrdinalIgnoreCase)) continue;

                // Skip rows already marked to avoid redundant writes
                if (d.TryGetValue("IsPurchased", out var existing) &&
                    existing is true or JsonElement { ValueKind: JsonValueKind.True }) continue;

                var searchKey = GetStr(d, "ProductSearchKey");
                bool matches = !matchOnMspc ||
                               (searchKey is not null && mspcSet.Contains(searchKey));
                if (!matches) continue;

                await GetGraph().Sites[siteId].Lists[sliListId].Items[item.Id].Fields
                    .PatchAsync(new FieldValueSet
                    {
                        AdditionalData = new Dictionary<string, object?>
                        {
                            ["IsPurchased"]     = true,
                            ["PurchaseRecordId"] = poSpItemId,
                        }
                    });

                patched++;
                _log.LogInformation("[PO] Marked SLI {SliId} purchased (MSPC={Key}, PO={PoId})",
                    item.Id, searchKey ?? "n/a", poSpItemId);
            }

            if (page.OdataNextLink is null) break;
            page = await GetGraph().Sites[siteId].Lists[sliListId].Items
                .WithUrl(page.OdataNextLink)
                .GetAsync();
        }

        _log.LogInformation("[PO] Updated {Count} SLI item(s) as purchased for [{RfqId}] {Supplier}",
            patched, rfqId, supplierName);

        // Also update matching RLI rows so Shredder can show the PO badge on the group header.
        // UpdateRliPurchaseStatusAsync previously only patched SLI, leaving RfqLineItemData.IsPurchased
        // and RfqLineItemData.PoNumber blank  -- causing the PO badge to never appear.
        if (!string.IsNullOrWhiteSpace(poNumber))
        {
            var rliListId = await GetRfqLineItemsListIdAsync();
            var rliCol    = await ResolveRfqIdColumnAsync(siteId, rliListId);

            var rliPage = await GetGraph().Sites[siteId].Lists[rliListId].Items
                .GetAsync(req =>
                {
                    req.QueryParameters.Expand = [$"fields($select=id,{rliCol},MSPC,IsPurchased)"];
                    req.QueryParameters.Top    = 5000;
                });

            while (rliPage?.Value is not null)
            {
                foreach (var item in rliPage.Value)
                {
                    if (item.Id is null || item.Fields?.AdditionalData is not { } d) continue;

                    var itemRfqId = d.TryGetValue(rliCol,         out var v0) ? v0?.ToString()
                                  : d.TryGetValue("RFQ_x005F_ID", out var v1) ? v1?.ToString()
                                  : d.TryGetValue("RFQ_ID",        out var v2) ? v2?.ToString()
                                  : null;
                    if (!string.Equals(itemRfqId, rfqId, StringComparison.OrdinalIgnoreCase)) continue;

                    // Skip if already marked
                    if (d.TryGetValue("IsPurchased", out var ip) &&
                        ip is true or JsonElement { ValueKind: JsonValueKind.True }) continue;

                    var rliMspc = d.TryGetValue("MSPC", out var vm) ? vm?.ToString() : null;
                    bool rliMatches = !matchOnMspc ||
                                     (rliMspc is not null && mspcSet.Contains(rliMspc));
                    if (!rliMatches) continue;

                    await GetGraph().Sites[siteId].Lists[rliListId].Items[item.Id].Fields
                        .PatchAsync(new FieldValueSet
                        {
                            AdditionalData = new Dictionary<string, object?>
                            {
                                ["IsPurchased"] = true,
                                ["PoNumber"]    = poNumber,
                            }
                        });
                    _log.LogInformation("[PO] Marked RLI {Id} purchased (MSPC={Mspc}, PO={PoNum}, RFQ={Rfq})",
                        item.Id, rliMspc ?? "n/a", poNumber, rfqId);
                }

                if (rliPage.OdataNextLink is null) break;
                rliPage = await GetGraph().Sites[siteId].Lists[rliListId].Items
                    .WithUrl(rliPage.OdataNextLink).GetAsync();
            }
        }

        await CheckAndCompleteRfqAsync(siteId, rfqId);
    }

    /// <summary>
    /// Checks whether all RFQ Line Items for <paramref name="rfqId"/> are fully covered by
    /// purchase orders (total PO quantity >= requested quantity per MSPC). If so, marks the
    /// RFQ Reference as Complete.
    /// </summary>
    private async Task CheckAndCompleteRfqAsync(string siteId, string rfqId)
    {
        // ??"?????"??? 1. Read RFQ Line Items for this rfqId ??"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"???
        var rliListId = await GetRfqLineItemsListIdAsync();
        var rliCol    = await ResolveRfqIdColumnAsync(siteId, rliListId);

        var rliPage = await GetGraph().Sites[siteId].Lists[rliListId].Items
            .GetAsync(req =>
            {
                req.QueryParameters.Expand = [$"fields($select={rliCol},MSPC,Units)"];
                req.QueryParameters.Top    = 5000;
            });

        var rliItems = new List<(string Mspc, double Units)>();
        foreach (var item in rliPage?.Value ?? [])
        {
            if (item.Fields?.AdditionalData is not { } d) continue;
            var itemRfqId = d.TryGetValue(rliCol,         out var v0) ? v0?.ToString()
                          : d.TryGetValue("RFQ_x005F_ID", out var v1) ? v1?.ToString()
                          : d.TryGetValue("RFQ_ID",        out var v2) ? v2?.ToString()
                          : null;
            if (!string.Equals(itemRfqId, rfqId, StringComparison.OrdinalIgnoreCase)) continue;

            var mspc = d.TryGetValue("MSPC", out var vm) ? vm?.ToString() : null;
            if (string.IsNullOrEmpty(mspc)) continue;

            double units = 0;
            if (d.TryGetValue("Units", out var vu) && vu is not null)
            {
                var s = vu is JsonElement je ? je.ToString() : vu.ToString() ?? "";
                double.TryParse(s,
                    System.Globalization.NumberStyles.Any,
                    System.Globalization.CultureInfo.InvariantCulture,
                    out units);
            }
            rliItems.Add((mspc, units));
        }

        if (rliItems.Count == 0)
        {
            _log.LogInformation("[PO] No RFQ Line Items found for [{RfqId}]  -- skipping completion check", rfqId);
            return;
        }

        // ??"?????"??? 2. Aggregate PO quantities per MSPC for this rfqId ??"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"???
        var poListId = await GetPurchaseOrdersListIdAsync();
        var poQtyByMspc = new Dictionary<string, double>(StringComparer.OrdinalIgnoreCase);

        var poPage = await GetGraph().Sites[siteId].Lists[poListId].Items
            .GetAsync(req =>
            {
                req.QueryParameters.Expand = ["fields($select=RFQ_ID,LineItems)"];
                req.QueryParameters.Top    = 5000;
            });

        var jsonOpts = new JsonSerializerOptions { PropertyNameCaseInsensitive = true };
        while (poPage?.Value is not null)
        {
            foreach (var item in poPage.Value)
            {
                if (item.Fields?.AdditionalData is not { } d) continue;
                var poRfqId = GetStr(d, "RFQ_ID") ?? GetStr(d, "RFQ_x005F_ID") ?? "";
                if (!string.Equals(poRfqId, rfqId, StringComparison.OrdinalIgnoreCase)) continue;

                var lineItemsJson = GetStr(d, "LineItems") ?? "[]";
                try
                {
                    var lineItems = JsonSerializer.Deserialize<List<PoLineItem>>(lineItemsJson, jsonOpts) ?? [];
                    foreach (var li in lineItems)
                    {
                        if (string.IsNullOrEmpty(li.Mspc)) continue;
                        poQtyByMspc.TryGetValue(li.Mspc, out var existing);
                        poQtyByMspc[li.Mspc] = existing + (li.Quantity ?? 0);
                    }
                }
                catch { /* malformed LineItems JSON  -- skip */ }
            }

            if (poPage.OdataNextLink is null) break;
            poPage = await GetGraph().Sites[siteId].Lists[poListId].Items
                .WithUrl(poPage.OdataNextLink).GetAsync();
        }

        // ??"?????"??? 3. Check every RLI item is fully covered ??"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"???
        foreach (var (mspc, requestedQty) in rliItems)
        {
            poQtyByMspc.TryGetValue(mspc, out var poQty);
            bool covered = requestedQty > 0 ? poQty >= requestedQty : poQty > 0;
            if (!covered)
            {
                _log.LogInformation(
                    "[PO] [{RfqId}] not complete  -- MSPC {Mspc} needs {Required}, POs have {Covered}",
                    rfqId, mspc, requestedQty, poQty);
                return;
            }
        }

        _log.LogInformation("[PO] All RFQ Line Items covered  -- marking [{RfqId}] Complete", rfqId);
        await SetRfqCompleteAsync(rfqId, true);
    }

    /// <summary>
    /// Matches PO line items to RFQ Line Items by MSPC when the PO has no RFQ ID.
    /// Patches matched RLI rows with <c>IsPurchased=true</c> and <c>PoNumber</c>.
    /// Skips RLIs whose parent RFQ Reference is already Complete.
    /// After patching, runs a completion check on each affected RFQ.
    /// </summary>
    public async Task<HashSet<string>> MatchAndMarkRliByMspcAsync(
        string supplierName, string? poNumber, List<PoLineItem> poLineItems)
    {
        // Only valid MSPC codes contain a forward slash (e.g. ASH3003/040).
        // Filter out supplier codes, part numbers, or other non-MSPC values.
        var mspcSet = poLineItems
            .Select(li => li.Mspc?.Trim())
            .Where(m => !string.IsNullOrEmpty(m) && m!.Contains('/'))
            .ToHashSet(StringComparer.OrdinalIgnoreCase);

        if (mspcSet.Count == 0)
        {
            _log.LogInformation("[PO] No valid MSPC codes (containing '/')  -- skipping RLI MSPC match for {Supplier}", supplierName);
            return [];
        }

        var siteId = await GetSiteIdAsync();
        var listId = await GetRfqLineItemsListIdAsync();
        var col    = await ResolveRfqIdColumnAsync(siteId, listId);

        // Read all RFQ Line Items, collect candidates matching the MSPC set.
        // Also track already-purchased RLI items so their rfqId+mspc can drive SLI updates
        // when SLI never got stamped (e.g. PO was processed before this logic existed).
        var candidates    = new List<(string ItemId, string RfqId, string? Mspc)>();
        var alreadyPurchased = new List<(string RfqId, string? Mspc)>(); // RLI rows already marked

        var page = await GetGraph().Sites[siteId].Lists[listId].Items
            .GetAsync(req =>
            {
                req.QueryParameters.Expand = [$"fields($select=id,{col},MSPC,IsPurchased)"];
                req.QueryParameters.Top    = 5000;
            });

        while (page?.Value is not null)
        {
            foreach (var item in page.Value)
            {
                if (item.Id is null || item.Fields?.AdditionalData is not { } d) continue;

                var mspc = d.TryGetValue("MSPC", out var vm) ? vm?.ToString() : null;
                if (string.IsNullOrEmpty(mspc) || !mspcSet.Contains(mspc)) continue;

                var rfqId = d.TryGetValue(col,            out var v0) ? v0?.ToString()
                          : d.TryGetValue("RFQ_x005F_ID", out var v1) ? v1?.ToString()
                          : d.TryGetValue("RFQ_ID",        out var v2) ? v2?.ToString()
                          : null;
                if (string.IsNullOrEmpty(rfqId)) continue;

                // Track already-purchased separately so SLI can still be stamped
                if (d.TryGetValue("IsPurchased", out var ip) &&
                    ip is true or JsonElement { ValueKind: JsonValueKind.True })
                {
                    alreadyPurchased.Add((rfqId, mspc));
                    continue;
                }

                candidates.Add((item.Id, rfqId, mspc));
            }

            if (page.OdataNextLink is null) break;
            page = await GetGraph().Sites[siteId].Lists[listId].Items
                .WithUrl(page.OdataNextLink).GetAsync();
        }

        if (candidates.Count == 0 && alreadyPurchased.Count == 0)
        {
            _log.LogInformation("[PO] No matching RLI items found for MSPCs [{Mspc}]",
                string.Join(", ", mspcSet));
            return [];
        }

        if (candidates.Count == 0)
        {
            _log.LogInformation("[PO] No unmatched RLI items found for MSPCs [{Mspc}]  -- all already purchased; checking SLI",
                string.Join(", ", mspcSet));
        }

        // Cache complete-RFQ checks to avoid repeated full-list reads
        var completeCache = new Dictionary<string, bool>(StringComparer.OrdinalIgnoreCase);
        async Task<bool> IsComplete(string id)
        {
            if (!completeCache.TryGetValue(id, out var v))
                completeCache[id] = v = await IsRfqCompleteAsync(siteId, id);
            return v;
        }

        int patched = 0;
        var affectedRfqIds = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

        foreach (var (itemId, rfqId, mspc) in candidates)
        {
            if (await IsComplete(rfqId)) continue;

            await GetGraph().Sites[siteId].Lists[listId].Items[itemId].Fields
                .PatchAsync(new FieldValueSet
                {
                    AdditionalData = new Dictionary<string, object?>
                    {
                        ["IsPurchased"] = true,
                        ["PoNumber"]    = poNumber,
                    }
                });

            affectedRfqIds.Add(rfqId);
            patched++;
            _log.LogInformation("[PO] Marked RLI {Id} purchased (MSPC={Mspc}, PO={PoNum}, RFQ={Rfq})",
                itemId, mspc, poNumber ?? "n/a", rfqId);
        }

        _log.LogInformation("[PO] MSPC match: {Count} RLI item(s) marked purchased for {Supplier}",
            patched, supplierName);

        // Also stamp matching SLI rows so Shredder can show the PO badge on group headers.
        // Covers both freshly-patched RLI items AND already-purchased RLI items whose
        // corresponding SLI rows may never have been stamped (backfill scenario).
        {
            // Build rfqId ?????' mspc set from all matching RLI rows (patched now + already purchased)
            var sliTargets = new Dictionary<string, HashSet<string>>(StringComparer.OrdinalIgnoreCase);

            void AddToSliTargets(string rfqId, string? mspc)
            {
                if (string.IsNullOrEmpty(mspc)) return;
                if (!sliTargets.TryGetValue(rfqId, out var set))
                    sliTargets[rfqId] = set = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
                set.Add(mspc);
            }

            foreach (var c in candidates.Where(c => affectedRfqIds.Contains(c.RfqId)))
                AddToSliTargets(c.RfqId, c.Mspc);
            foreach (var (rfqId, mspc) in alreadyPurchased)
                AddToSliTargets(rfqId, mspc);

            if (sliTargets.Count > 0)
            {
                var sliListId  = await GetSupplierLineItemsListIdAsync();
                int sliPatched = 0;

                var sliPage = await GetGraph().Sites[siteId].Lists[sliListId].Items
                    .GetAsync(req =>
                    {
                        req.QueryParameters.Expand = ["fields($select=id,RFQ_ID,SupplierName,ProductSearchKey,IsPurchased)"];
                        req.QueryParameters.Top    = 2000;
                    });

                while (sliPage?.Value is not null)
                {
                    foreach (var item in sliPage.Value)
                    {
                        if (item.Id is null || item.Fields?.AdditionalData is not { } d) continue;

                        var itemRfqId = GetStr(d, "RFQ_ID") ?? GetStr(d, "RFQ_x005F_ID") ?? "";
                        if (!sliTargets.TryGetValue(itemRfqId, out var targetMspcs)) continue;

                        var searchKey = GetStr(d, "ProductSearchKey");
                        if (searchKey is null || !targetMspcs.Contains(searchKey)) continue;

                        // Skip rows already marked
                        if (d.TryGetValue("IsPurchased", out var existing) &&
                            existing is true or JsonElement { ValueKind: JsonValueKind.True }) continue;

                        await GetGraph().Sites[siteId].Lists[sliListId].Items[item.Id].Fields
                            .PatchAsync(new FieldValueSet
                            {
                                AdditionalData = new Dictionary<string, object?>
                                {
                                    ["IsPurchased"]      = true,
                                    ["PurchaseRecordId"] = poNumber,
                                }
                            });

                        sliPatched++;
                        _log.LogInformation("[PO] Marked SLI {SliId} purchased via MSPC match (MSPC={Key}, PO={PoNum}, RFQ={Rfq})",
                            item.Id, searchKey, poNumber ?? "n/a", itemRfqId);
                    }

                    if (sliPage.OdataNextLink is null) break;
                    sliPage = await GetGraph().Sites[siteId].Lists[sliListId].Items
                        .WithUrl(sliPage.OdataNextLink).GetAsync();
                }

                _log.LogInformation("[PO] MSPC match: {Count} SLI item(s) marked purchased for {Supplier}",
                    sliPatched, supplierName);
            }
        }

        foreach (var rfqId in affectedRfqIds)
            await CheckRliAllPurchasedAsync(siteId, rfqId);

        return affectedRfqIds;
    }

    /// <summary>
    /// Marks the RFQ Reference as Complete when every RFQ Line Item for
    /// <paramref name="rfqId"/> has <c>IsPurchased = true</c>.
    /// Used after MSPC-based PO matching where quantity data is unavailable.
    /// </summary>
    private async Task CheckRliAllPurchasedAsync(string siteId, string rfqId)
    {
        var listId = await GetRfqLineItemsListIdAsync();
        var col    = await ResolveRfqIdColumnAsync(siteId, listId);

        var page = await GetGraph().Sites[siteId].Lists[listId].Items
            .GetAsync(req =>
            {
                req.QueryParameters.Expand = [$"fields($select={col},IsPurchased)"];
                req.QueryParameters.Top    = 5000;
            });

        bool hasItems = false;
        bool allPurchased = true;

        foreach (var item in page?.Value ?? [])
        {
            if (item.Fields?.AdditionalData is not { } d) continue;
            var itemRfqId = d.TryGetValue(col,            out var v0) ? v0?.ToString()
                          : d.TryGetValue("RFQ_x005F_ID", out var v1) ? v1?.ToString()
                          : d.TryGetValue("RFQ_ID",        out var v2) ? v2?.ToString()
                          : null;
            if (!string.Equals(itemRfqId, rfqId, StringComparison.OrdinalIgnoreCase)) continue;

            hasItems = true;
            if (!d.TryGetValue("IsPurchased", out var ip) ||
                ip is not (true or JsonElement { ValueKind: JsonValueKind.True }))
            {
                allPurchased = false;
                break;
            }
        }

        if (hasItems && allPurchased)
        {
            _log.LogInformation("[PO] All RLI items purchased  -- marking [{RfqId}] Complete", rfqId);
            await SetRfqCompleteAsync(rfqId, true);
        }
    }

    // ??"?????"??? MessageId backfill ??"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"???

    /// <summary>
    /// Scans SupplierResponses rows from the last <paramref name="days"/> days that have no MessageId,
    /// finds the matching Graph email by sender + received time, and patches MessageId onto the SR
    /// and all linked SLI rows.
    /// Returns (patched, skipped) counts.
    /// </summary>
    public async Task<(int Patched, int Skipped)> BackfillMessageIdsAsync(
        MailService mail, int days = 7, CancellationToken ct = default)
    {
        var siteId    = await GetSiteIdAsync();
        var srListId  = await GetSupplierResponsesListIdAsync();
        var sliListId = await GetSupplierLineItemsListIdAsync();
        var cutoff    = DateTimeOffset.UtcNow.AddDays(-days);

        // Load all SR rows within the window
        var srPage = await GetGraph().Sites[siteId].Lists[srListId].Items
            .GetAsync(req =>
            {
                req.QueryParameters.Expand = ["fields($select=id,EmailFrom,ReceivedAt,MessageId)"];
                req.QueryParameters.Top    = 2000;
            });

        var candidates = new List<(string SpId, string EmailFrom, DateTimeOffset ReceivedAt)>();
        while (srPage?.Value is not null)
        {
            foreach (var item in srPage.Value)
            {
                if (item.Id is null || item.Fields?.AdditionalData is not { } d) continue;
                // Skip rows that already have a MessageId
                var existing = GetStr(d, "MessageId");
                if (!string.IsNullOrWhiteSpace(existing)) continue;

                var from = GetStr(d, "EmailFrom");
                if (string.IsNullOrWhiteSpace(from)) continue;

                var recStr = GetStr(d, "ReceivedAt");
                if (!DateTimeOffset.TryParse(recStr, out var rec)) continue;
                if (rec < cutoff) continue;

                candidates.Add((item.Id, from, rec));
            }
            if (srPage.OdataNextLink is null) break;
            srPage = await GetGraph().Sites[siteId].Lists[srListId].Items
                .WithUrl(srPage.OdataNextLink).GetAsync();
        }

        _log.LogInformation("[Backfill] {Count} SR row(s) in last {Days} days missing MessageId", candidates.Count, days);
        int patched = 0, skipped = 0;

        foreach (var (srSpId, emailFrom, receivedAt) in candidates)
        {
            ct.ThrowIfCancellationRequested();

            var msgId = await mail.FindMessageIdAsync(emailFrom, receivedAt);
            if (msgId is null)
            {
                _log.LogDebug("[Backfill] No Graph message found for SR {Id} ({From} ~{At})", srSpId, emailFrom, receivedAt);
                skipped++;
                continue;
            }

            // Patch SR with MessageId
            await GetGraph().Sites[siteId].Lists[srListId].Items[srSpId].Fields
                .PatchAsync(new FieldValueSet
                {
                    AdditionalData = new Dictionary<string, object?> { ["MessageId"] = msgId }
                });

            // Patch all linked SLI rows
            var sliPage = await GetGraph().Sites[siteId].Lists[sliListId].Items
                .GetAsync(req =>
                {
                    req.QueryParameters.Expand = ["fields($select=id,SupplierResponseId,MessageId)"];
                    req.QueryParameters.Top    = 500;
                });
            while (sliPage?.Value is not null)
            {
                foreach (var sliItem in sliPage.Value)
                {
                    if (sliItem.Id is null || sliItem.Fields?.AdditionalData is not { } sd) continue;
                    var srId = GetStr(sd, "SupplierResponseId");
                    if (!string.Equals(srId, srSpId, StringComparison.OrdinalIgnoreCase)) continue;
                    var existingMsgId = GetStr(sd, "MessageId");
                    if (!string.IsNullOrWhiteSpace(existingMsgId)) continue;

                    await GetGraph().Sites[siteId].Lists[sliListId].Items[sliItem.Id].Fields
                        .PatchAsync(new FieldValueSet
                        {
                            AdditionalData = new Dictionary<string, object?> { ["MessageId"] = msgId }
                        });
                }
                if (sliPage.OdataNextLink is null) break;
                sliPage = await GetGraph().Sites[siteId].Lists[sliListId].Items
                    .WithUrl(sliPage.OdataNextLink).GetAsync();
            }

            patched++;
            _log.LogInformation("[Backfill] Patched MessageId on SR {Id} ({From})", srSpId, emailFrom);
        }

        _log.LogInformation("[Backfill] MessageId backfill complete  -- patched={Patched}, skipped={Skipped}", patched, skipped);
        return (patched, skipped);
    }

    // ??"?????"??? QuoteReference backfill ??"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"???

    /// <summary>
    /// Scans SupplierResponse rows that have a MessageId but no QuoteReference,
    /// re-runs AI extraction on the original email/attachment, and patches
    /// QuoteReference onto the SR and all its child SLI rows.
    /// </summary>
    public async Task<(int Patched, int Skipped)> BackfillQuoteReferencesAsync(
        MailService mail, IAiExtractionService ai, string mailbox,
        int days = 90, CancellationToken ct = default)
    {
        var siteId    = await GetSiteIdAsync();
        var srListId  = await GetSupplierResponsesListIdAsync();
        var sliListId = await GetSupplierLineItemsListIdAsync();
        var cutoff    = DateTimeOffset.UtcNow.AddDays(-days);

        // Load SR rows that have a MessageId but no QuoteReference
        var candidates = new List<(string SpId, string MessageId, string ProcSrc, string? EmailBody)>();
        string? nextLink = null;
        do
        {
            var page = nextLink is null
                ? await GetGraph().Sites[siteId].Lists[srListId].Items.GetAsync(req =>
                {
                    req.QueryParameters.Expand = ["fields($select=id,MessageId,ReceivedAt,QuoteReference,ProcessingSource,EmailBody)"];
                    req.QueryParameters.Top    = 500;
                })
                : await GetGraph().Sites[siteId].Lists[srListId].Items.WithUrl(nextLink).GetAsync();

            foreach (var item in page?.Value ?? [])
            {
                if (item.Id is null || item.Fields?.AdditionalData is not { } d) continue;
                var msgId  = GetStr(d, "MessageId");
                var qRef   = GetStr(d, "QuoteReference");
                var recStr = GetStr(d, "ReceivedAt");
                if (string.IsNullOrWhiteSpace(msgId)) continue;
                if (!string.IsNullOrWhiteSpace(qRef)) continue;
                if (DateTimeOffset.TryParse(recStr, out var rec) && rec < cutoff) continue;
                var procSrc = GetStr(d, "ProcessingSource") ?? "body";
                var body    = GetStr(d, "EmailBody");
                candidates.Add((item.Id, msgId, procSrc, body));
            }
            nextLink = page?.OdataNextLink;
        } while (nextLink is not null);

        _log.LogInformation("[Backfill] {Count} SR row(s) missing QuoteReference (last {Days} days)", candidates.Count, days);
        int patched = 0, skipped = 0;

        foreach (var (srSpId, messageId, procSrc, emailBody) in candidates)
        {
            ct.ThrowIfCancellationRequested();
            string? quoteRef = null;

            _log.LogInformation("[Backfill] Processing SR {Id} procSrc={ProcSrc} msg={MsgId}", srSpId, procSrc, messageId);

            try
            {
                // Prefer attachment extraction (same source as original processing)
                if (procSrc == "attachment")
                {
                    List<Attachment> attachments;
                    try
                    {
                        attachments = await mail.GetAttachmentsAsync(mailbox, messageId);
                    }
                    catch (Exception ex)
                    {
                        _log.LogWarning(ex, "[Backfill] SR {Id}  -- could not fetch attachments from Graph (msg={MsgId})", srSpId, messageId);
                        skipped++;
                        continue;
                    }

                    var pdf = attachments
                        .OfType<FileAttachment>()
                        .FirstOrDefault(a =>
                            a.ContentType?.Contains("pdf", StringComparison.OrdinalIgnoreCase) == true
                            || a.Name?.EndsWith(".pdf", StringComparison.OrdinalIgnoreCase) == true);

                    if (pdf?.ContentBytes is { } bytes)
                    {
                        _log.LogInformation("[Backfill] SR {Id}  -- running AI on PDF attachment '{Name}'", srSpId, pdf.Name);
                        var req = new ExtractRequest
                        {
                            Content     = string.Empty,
                            SourceType  = "attachment",
                            FileName    = pdf.Name,
                            Base64Data  = Convert.ToBase64String(bytes),
                            ContentType = pdf.ContentType ?? "application/pdf",
                            BodyContext = emailBody,
                        };
                        var result = await ai.ExtractRfqAsync(req, ct);
                        quoteRef = result?.QuoteReference;
                        _log.LogInformation("[Backfill] SR {Id}  -- attachment extraction quoteRef={QuoteRef}", srSpId, quoteRef ?? "(null)");
                    }
                    else
                    {
                        _log.LogInformation("[Backfill] SR {Id}  -- no PDF attachment found ({Count} attachment(s) total)", srSpId, attachments.Count);
                    }
                }

                // Fall back to body extraction
                if (string.IsNullOrWhiteSpace(quoteRef))
                {
                    if (!string.IsNullOrWhiteSpace(emailBody))
                    {
                        _log.LogInformation("[Backfill] SR {Id}  -- running AI on stored email body ({Len} chars)", srSpId, emailBody!.Length);
                        var req = new ExtractRequest { Content = emailBody, SourceType = "body" };
                        var result = await ai.ExtractRfqAsync(req, ct);
                        quoteRef = result?.QuoteReference;
                        _log.LogInformation("[Backfill] SR {Id}  -- body extraction quoteRef={QuoteRef}", srSpId, quoteRef ?? "(null)");
                    }
                    else
                    {
                        _log.LogInformation("[Backfill] SR {Id}  -- no email body stored, nothing to extract", srSpId);
                    }
                }
            }
            catch (Exception ex)
            {
                _log.LogWarning(ex, "[Backfill] SR {Id}  -- extraction threw unexpectedly (msg={MsgId})", srSpId, messageId);
                skipped++;
                continue;
            }

            if (string.IsNullOrWhiteSpace(quoteRef))
            {
                _log.LogInformation("[Backfill] SR {Id}  -- skipped (no QuoteReference extracted)", srSpId);
                skipped++;
                continue;
            }

            // Patch the SR
            await GetGraph().Sites[siteId].Lists[srListId].Items[srSpId].Fields
                .PatchAsync(new FieldValueSet
                {
                    AdditionalData = new Dictionary<string, object?> { ["QuoteReference"] = quoteRef }
                });

            // Patch all linked SLI rows
            string? sliNext = null;
            do
            {
                var sliPage = sliNext is null
                    ? await GetGraph().Sites[siteId].Lists[sliListId].Items.GetAsync(req =>
                    {
                        req.QueryParameters.Expand = ["fields($select=id,SupplierResponseId,QuoteReference)"];
                        req.QueryParameters.Top    = 500;
                    })
                    : await GetGraph().Sites[siteId].Lists[sliListId].Items.WithUrl(sliNext).GetAsync();

                foreach (var sliItem in sliPage?.Value ?? [])
                {
                    if (sliItem.Id is null || sliItem.Fields?.AdditionalData is not { } sd) continue;
                    var srId = GetStr(sd, "SupplierResponseId");
                    if (!string.Equals(srId, srSpId, StringComparison.OrdinalIgnoreCase)) continue;
                    if (!string.IsNullOrWhiteSpace(GetStr(sd, "QuoteReference"))) continue;

                    await GetGraph().Sites[siteId].Lists[sliListId].Items[sliItem.Id].Fields
                        .PatchAsync(new FieldValueSet
                        {
                            AdditionalData = new Dictionary<string, object?> { ["QuoteReference"] = quoteRef }
                        });
                }
                sliNext = sliPage?.OdataNextLink;
            } while (sliNext is not null);

            patched++;
            _log.LogInformation("[Backfill] Patched QuoteReference '{Ref}' on SR {Id}", quoteRef, srSpId);
        }

        _log.LogInformation("[Backfill] QuoteReference backfill complete  -- patched={Patched}, skipped={Skipped}", patched, skipped);
        return (patched, skipped);
    }

    // ??"?????"??? Deduplication ??"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"?????"???

    /// <summary>
    /// Removes duplicate SupplierResponse + SupplierLineItem rows:
    /// 1. SR rows with no MessageId (within <paramref name="days"/> days)  -- deleted with their SLIs.
    /// 2. SR rows sharing the same MessageId  -- keep the one with the most populated fields,
    ///    delete the rest with their SLIs.
    /// Returns counts of deleted SR and SLI rows.
    /// </summary>
    public async Task<(int SrDeleted, int SliDeleted)> DeduplicateSupplierResponsesAsync(
        int days = 7, CancellationToken ct = default)
    {
        var siteId    = await GetSiteIdAsync();
        var srListId  = await GetSupplierResponsesListIdAsync();
        var sliListId = await GetSupplierLineItemsListIdAsync();
        var cutoff    = DateTimeOffset.UtcNow.AddDays(-days);

        // Load all SR rows in window with key fields
        var srPage = await GetGraph().Sites[siteId].Lists[srListId].Items
            .GetAsync(req =>
            {
                req.QueryParameters.Expand = ["fields($select=id,EmailFrom,ReceivedAt,MessageId,RFQ_ID,SupplierName,ProcessingSource,QuoteReference)"];
                req.QueryParameters.Top    = 2000;
            });

        // srId ?????' (messageId, receivedAt, score)
        var allSr = new List<(string SpId, string? MessageId, DateTimeOffset ReceivedAt, int Score)>();
        while (srPage?.Value is not null)
        {
            foreach (var item in srPage.Value)
            {
                if (item.Id is null || item.Fields?.AdditionalData is not { } d) continue;
                var recStr = GetStr(d, "ReceivedAt");
                if (!DateTimeOffset.TryParse(recStr, out var rec)) continue;
                if (rec < cutoff) continue;

                var msgId  = GetStr(d, "MessageId");
                var source = GetStr(d, "ProcessingSource");
                var qref   = GetStr(d, "QuoteReference");
                // Score: prefer rows with MessageId, attachment source, and quote reference
                int score = (string.IsNullOrWhiteSpace(msgId)  ? 0 : 4)
                          + (source?.Equals("attachment", StringComparison.OrdinalIgnoreCase) == true ? 2 : 0)
                          + (string.IsNullOrWhiteSpace(qref)   ? 0 : 1);
                allSr.Add((item.Id, msgId, rec, score));
            }
            if (srPage.OdataNextLink is null) break;
            srPage = await GetGraph().Sites[siteId].Lists[srListId].Items
                .WithUrl(srPage.OdataNextLink).GetAsync();
        }

        var toDelete = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

        // 1. Mark no-MessageId rows for deletion
        foreach (var (spId, msgId, _, _) in allSr)
            if (string.IsNullOrWhiteSpace(msgId))
                toDelete.Add(spId);

        // 2. For rows sharing a MessageId, keep the highest-scoring one
        foreach (var grp in allSr
            .Where(r => !string.IsNullOrWhiteSpace(r.MessageId))
            .GroupBy(r => r.MessageId!, StringComparer.OrdinalIgnoreCase)
            .Where(g => g.Count() > 1))
        {
            var keep = grp.OrderByDescending(r => r.Score).First();
            foreach (var dup in grp.Where(r => r.SpId != keep.SpId))
                toDelete.Add(dup.SpId);
        }

        _log.LogInformation("[Dedup] {Count} SR row(s) marked for deletion", toDelete.Count);
        int srDeleted = 0, sliDeleted = 0;

        // Delete SLI rows linked to doomed SRs
        var sliPage = await GetGraph().Sites[siteId].Lists[sliListId].Items
            .GetAsync(req =>
            {
                req.QueryParameters.Expand = ["fields($select=id,SupplierResponseId)"];
                req.QueryParameters.Top    = 2000;
            });

        while (sliPage?.Value is not null)
        {
            foreach (var item in sliPage.Value)
            {
                ct.ThrowIfCancellationRequested();
                if (item.Id is null || item.Fields?.AdditionalData is not { } d) continue;
                var srId = GetStr(d, "SupplierResponseId");
                if (srId is null || !toDelete.Contains(srId)) continue;

                await GetGraph().Sites[siteId].Lists[sliListId].Items[item.Id]
                    .DeleteAsync();
                sliDeleted++;
            }
            if (sliPage.OdataNextLink is null) break;
            sliPage = await GetGraph().Sites[siteId].Lists[sliListId].Items
                .WithUrl(sliPage.OdataNextLink).GetAsync();
        }

        // Delete the SR rows themselves
        foreach (var spId in toDelete)
        {
            ct.ThrowIfCancellationRequested();
            await GetGraph().Sites[siteId].Lists[srListId].Items[spId].DeleteAsync();
            srDeleted++;
            _log.LogInformation("[Dedup] Deleted SR {Id}", spId);
        }

        _log.LogInformation("[Dedup] Complete  -- SR deleted={Sr}, SLI deleted={Sli}", srDeleted, sliDeleted);
        return (srDeleted, sliDeleted);
    }

    /// <summary>
    /// Reads every row from the PurchaseOrders list and runs <see cref="UpdateRliPurchaseStatusAsync"/>
    /// for each one that has a known RFQ ID.  Use this to backfill <c>IsPurchased</c> /
    /// <c>PurchaseRecordId</c> on SLI rows that were written before the purchase-status feature existed.
    /// Returns (processed, skipped) counts.
    /// </summary>
    public async Task<(int Processed, int Skipped)> BackfillRliPurchaseStatusAsync(
        int? days = null,
        CancellationToken ct = default)
    {
        var records = await ReadPurchaseOrdersAsync();

        if (days is > 0)
        {
            var cutoff = DateTimeOffset.UtcNow.AddDays(-days.Value);
            records = records.Where(r =>
                DateTimeOffset.TryParse(r.ReceivedAt, out var dt) ? dt >= cutoff : true).ToList();
            _log.LogInformation("[PO] Backfill scoped to last {Days} days  -- {Count} record(s) in window",
                days.Value, records.Count);
        }

        var jsonOpts = new JsonSerializerOptions { PropertyNameCaseInsensitive = true };
        int processed = 0, skipped = 0;

        foreach (var rec in records)
        {
            ct.ThrowIfCancellationRequested();

            if (string.IsNullOrWhiteSpace(rec.SpItemId))
            {
                skipped++;
                continue;
            }

            List<PoLineItem> lineItems;
            try { lineItems = JsonSerializer.Deserialize<List<PoLineItem>>(rec.LineItems, jsonOpts) ?? []; }
            catch { lineItems = []; }

            if (string.IsNullOrWhiteSpace(rec.RfqId) ||
                rec.RfqId.Equals("UNKNOWN", StringComparison.OrdinalIgnoreCase))
            {
                // No RFQ ID  -- match by MSPC directly against RFQ Line Items
                await MatchAndMarkRliByMspcAsync(rec.SupplierName, rec.PoNumber, lineItems);
            }
            else
            {
                await UpdateRliPurchaseStatusAsync(rec.RfqId, rec.SupplierName, rec.SpItemId, lineItems, rec.PoNumber);
            }
            processed++;
        }

        _log.LogInformation("[PO] Backfill complete  -- processed={Processed}, skipped={Skipped}",
            processed, skipped);
        return (processed, skipped);
    }

}

public record QcListResult(string[] Columns, string[][] Rows, string[] ItemIds, DateTime? LastModified = null);

public record LqUpdateResult(
    List<LqMatch> Updated,
    List<string>  Misses);

public record LqMatch(
    string Metal,
    string Shape,
    double PricePerPound,
    int    QuoteCount,
    double MinPricePerPound,
    double MaxPricePerPound);
