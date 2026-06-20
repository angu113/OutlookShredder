using System.Globalization;
using OutlookShredder.Proxy.Models;

namespace OutlookShredder.Proxy.Services;

// SalesOrderHistory list helpers: provisioning, bulk-load insert, and full/ids reads. Kept in their own
// partial so the 600KB SharePointService.cs doesn't grow. Reuses GetGraph()/GetSiteIdAsync/EnsureIndexedListAsync.
// Backs the Sales-Order-history feature (bulk load -> serve to the Raptor call card -> append on ERP-doc arrival).
public partial class SharePointService
{
    /// <summary>Raised after a NEW SalesOrderHistory row is written by <see cref="AppendSalesOrderIfAbsentAsync"/>
    /// (the FileWatcher append path), so this proxy's serving cache can merge it without an SP re-read. A
    /// plain event so SharePointService needn't reference the cache (avoids a DI cycle); cross-proxy
    /// freshness rides the existing EventType="ERP" bus event instead.</summary>
    public event Action<Models.SalesOrderRecord>? SalesOrderAppended;

    private async Task<string> GetSalesOrderHistoryListIdAsync() =>
        _salesOrderHistoryListId ??= await ResolveListIdAsync("SalesOrderHistory");

    /// <summary>Provisions the <c>SalesOrderHistory</c> list (one row per HSK-SO#). Indexes Title (OrderId
    /// dedup key), CustomerName (the call-card join key) and OrderDate (recency) AT construction — SP 500s
    /// on an unindexed $filter/$orderby even on a tiny list, and a brand-new column can't be observed first,
    /// so the known state must be "created indexed". Idempotent.</summary>
    public async Task<Dictionary<string, object>> EnsureSalesOrderHistoryListAsync(CancellationToken ct = default)
    {
        var siteId  = await GetSiteIdAsync();
        var results = new Dictionary<string, object>
        {
            ["SalesOrderHistory"] = await EnsureIndexedListAsync(siteId, "SalesOrderHistory",
            [
                ("CustomerName",    "text"),
                ("OrderDate",       "dateTime"),
                ("Status",          "text"),
                ("SecondaryStatus", "text"),
                ("CustomerPo",      "text"),
                ("NetAmount",       "number"),
                ("GrossAmount",     "number"),
                ("PctPaid",         "number"),
                ("DeliveryDate",    "dateTime"),
                ("Weight",          "number"),
                ("Source",          "text"),
            ], "Title", "CustomerName", "OrderDate"),
        };

        _salesOrderHistoryListId = await ResolveListIdAsync("SalesOrderHistory");
        return results;
    }

    /// <summary>Every OrderId (Title) already on the list — the dedup-vs-existing set for the resumable
    /// bulk load and the dry-run counts. Paginated (Top=999 + nextLink).</summary>
    public async Task<HashSet<string>> ReadAllSalesOrderIdsAsync(CancellationToken ct = default)
    {
        var siteId = await GetSiteIdAsync();
        var listId = await GetSalesOrderHistoryListIdAsync();
        var ids    = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

        var page = await GetGraph().Sites[siteId].Lists[listId].Items
            .GetAsync(r =>
            {
                r.QueryParameters.Expand = ["fields($select=Title)"];
                r.QueryParameters.Top    = 999;
            }, ct);

        while (page?.Value is not null)
        {
            foreach (var item in page.Value)
            {
                if (item.Fields?.AdditionalData is not { } d) continue;
                var id = GetStr(d, "Title");
                if (!string.IsNullOrWhiteSpace(id)) ids.Add(id);
            }
            if (page.OdataNextLink is null) break;
            page = await new Microsoft.Graph.Sites.Item.Lists.Item.Items.ItemsRequestBuilder(
                page.OdataNextLink, GetGraph().RequestAdapter).GetAsync(cancellationToken: ct);
        }
        return ids;
    }

    /// <summary>Reads EVERY SalesOrderHistory row (paginated) for the in-memory serving cache.</summary>
    public async Task<List<SalesOrderRecord>> ReadAllSalesOrdersAsync(CancellationToken ct = default)
    {
        var siteId = await GetSiteIdAsync();
        var listId = await GetSalesOrderHistoryListIdAsync();
        var result = new List<SalesOrderRecord>();

        var page = await GetGraph().Sites[siteId].Lists[listId].Items
            .GetAsync(r =>
            {
                r.QueryParameters.Expand =
                    ["fields($select=Title,CustomerName,OrderDate,Status,SecondaryStatus,CustomerPo,NetAmount,GrossAmount,PctPaid,DeliveryDate,Weight,Source)"];
                r.QueryParameters.Top = 999;
            }, ct);

        while (page?.Value is not null)
        {
            foreach (var item in page.Value)
            {
                if (item.Fields?.AdditionalData is not { } d) continue;
                var id = GetStr(d, "Title");
                if (string.IsNullOrWhiteSpace(id)) continue;
                result.Add(new SalesOrderRecord
                {
                    OrderId         = id,
                    CustomerName    = GetStr(d, "CustomerName") ?? "",
                    OrderDate       = ParseSpDate(GetStr(d, "OrderDate")),
                    Status          = GetStr(d, "Status"),
                    SecondaryStatus = GetStr(d, "SecondaryStatus"),
                    CustomerPo      = GetStr(d, "CustomerPo"),
                    NetAmount       = ParseSpNum(GetStr(d, "NetAmount")),
                    GrossAmount     = ParseSpNum(GetStr(d, "GrossAmount")),
                    PctPaid         = ParseSpNum(GetStr(d, "PctPaid")),
                    DeliveryDate    = ParseSpDate(GetStr(d, "DeliveryDate")),
                    Weight          = ParseSpNum(GetStr(d, "Weight")),
                    Source          = GetStr(d, "Source"),
                });
            }
            if (page.OdataNextLink is null) break;
            page = await new Microsoft.Graph.Sites.Item.Lists.Item.Items.ItemsRequestBuilder(
                page.OdataNextLink, GetGraph().RequestAdapter).GetAsync(cancellationToken: ct);
        }
        return result;
    }

    /// <summary>Inserts the parsed rows that are not already present (by OrderId), in parallel with a
    /// per-row timeout and throttle-retry. Resumable: it reads the existing OrderIds first and skips them,
    /// so a crash at 20k just re-runs and inserts the remainder. <paramref name="onAdded"/> is invoked with
    /// the running added-count for progress reporting.</summary>
    public async Task<(int Added, int AlreadyPresent, int Failed)> InsertSalesOrdersAsync(
        IReadOnlyList<CustomerImportService.SalesOrderRow> rows,
        Action<int>? onAdded = null, CancellationToken ct = default)
    {
        var siteId = await GetSiteIdAsync();
        var listId = await GetSalesOrderHistoryListIdAsync();

        var existing = await ReadAllSalesOrderIdsAsync(ct);
        var toInsert = rows.Where(r => !existing.Contains(r.OrderId)).ToList();
        int already  = rows.Count - toInsert.Count;
        _log.LogInformation("[SP] SalesOrderHistory load: {ToInsert} to insert, {Already} already present",
            toInsert.Count, already);

        int added = 0, failed = 0;
        await Parallel.ForEachAsync(toInsert,
            new ParallelOptions { MaxDegreeOfParallelism = 4, CancellationToken = ct },
            async (row, token) =>
            {
                using var rowCts = CancellationTokenSource.CreateLinkedTokenSource(token);
                rowCts.CancelAfter(TimeSpan.FromSeconds(45));
                try
                {
                    await PostSalesOrderWithRetryAsync(siteId, listId, BuildSalesOrderFields(row, "export"), rowCts.Token);
                    var n = Interlocked.Increment(ref added);
                    onAdded?.Invoke(n);
                    if (n % 1000 == 0)
                        _log.LogInformation("[SP] SalesOrderHistory load: {N}/{Total} inserted", n, toInsert.Count);
                }
                catch (Exception ex) when (!token.IsCancellationRequested)
                {
                    _log.LogWarning("[SP] SalesOrderHistory load: POST failed for {Order}: {Err}", row.OrderId, ex.Message);
                    Interlocked.Increment(ref failed);
                }
            });

        if (failed > 0)
            _log.LogWarning("[SP] SalesOrderHistory load: {Failed} row(s) failed to insert", failed);
        _log.LogInformation("[SP] SalesOrderHistory load complete: {Added} added, {Already} already present, {Failed} failed",
            added, already, failed);
        return (added, already, failed);
    }

    /// <summary>Appends one order as it arrives via an ERP SalesOrder doc (the FileWatcher path), insert-if-
    /// absent by OrderId so a row already backfilled from the bulk export is not duplicated. The ERP doc is
    /// sparse — only customer/date/total — so the total lands in <c>GrossAmount</c> (the card displays
    /// <c>NetAmount ?? GrossAmount</c>). Returns the new record (and raises <see cref="SalesOrderAppended"/>)
    /// when it inserted, or null when the order was already present or the inputs are unusable.</summary>
    public async Task<Models.SalesOrderRecord?> AppendSalesOrderIfAbsentAsync(
        string? orderId, string? customer, string? documentDate, string? totalAmount,
        string source = "erp-doc", CancellationToken ct = default)
    {
        if (string.IsNullOrWhiteSpace(orderId) || string.IsNullOrWhiteSpace(customer)) return null;
        orderId = orderId.Trim();
        if (!orderId.StartsWith("HSK-SO", StringComparison.OrdinalIgnoreCase)) return null;

        var siteId = await GetSiteIdAsync();
        var listId = await GetSalesOrderHistoryListIdAsync();

        // Already present? Title is indexed, so this single-item $filter is cheap (no full-list scan).
        var hit = await GetGraph().Sites[siteId].Lists[listId].Items
            .GetAsync(r =>
            {
                r.QueryParameters.Filter = $"fields/Title eq '{orderId.Replace("'", "''")}'";
                r.QueryParameters.Top    = 1;
            }, ct);
        if (hit?.Value is { Count: > 0 }) return null;

        var rec = new Models.SalesOrderRecord
        {
            OrderId      = orderId,
            CustomerName = customer.Trim(),
            OrderDate    = CustomerImportService.ParseErpDate(documentDate) ?? ParseSpDate(documentDate),
            GrossAmount  = CustomerImportService.ParseMoney(totalAmount),
            Source       = source,
        };

        await PostSalesOrderWithRetryAsync(siteId, listId, new Dictionary<string, object?>
        {
            ["Title"]        = rec.OrderId,
            ["CustomerName"] = rec.CustomerName,
            ["OrderDate"]    = rec.OrderDate?.ToString("o"),
            ["GrossAmount"]  = rec.GrossAmount,
            ["Source"]       = source,
        }, ct);

        _log.LogInformation("[SP] SalesOrderHistory appended {Order} ({Customer}) from {Source}",
            rec.OrderId, rec.CustomerName, source);
        SalesOrderAppended?.Invoke(rec);   // local write-through into this proxy's serving cache
        return rec;
    }

    /// <summary>POSTs a new SalesOrderHistory item, retrying on SharePoint throttling (429/503/504) with
    /// exponential backoff (mirrors <c>PatchListItemWithRetryAsync</c>).</summary>
    private async Task PostSalesOrderWithRetryAsync(
        string siteId, string listId, Dictionary<string, object?> fields, CancellationToken ct, int maxAttempts = 5)
    {
        for (int attempt = 1; ; attempt++)
        {
            try
            {
                await GetGraph().Sites[siteId].Lists[listId].Items
                    .PostAsync(new Microsoft.Graph.Models.ListItem
                    {
                        Fields = new Microsoft.Graph.Models.FieldValueSet { AdditionalData = fields }
                    }, cancellationToken: ct);
                return;
            }
            catch (Microsoft.Graph.Models.ODataErrors.ODataError ex)
                when (attempt < maxAttempts && ex.ResponseStatusCode is 429 or 503 or 504)
            {
                var delay = TimeSpan.FromSeconds(Math.Pow(2, attempt));   // 2s, 4s, 8s, 16s
                _log.LogWarning("[SP] SalesOrderHistory POST throttled (HTTP {Code}) — retry {A}/{M} in {S}s",
                    ex.ResponseStatusCode, attempt, maxAttempts, delay.TotalSeconds);
                await Task.Delay(delay, ct);
            }
        }
    }

    private static Dictionary<string, object?> BuildSalesOrderFields(
        CustomerImportService.SalesOrderRow row, string source) => new()
    {
        ["Title"]           = row.OrderId,
        ["CustomerName"]    = row.CustomerName,
        ["OrderDate"]       = row.OrderDate?.ToString("o"),
        ["Status"]          = row.Status,
        ["SecondaryStatus"] = row.SecondaryStatus,
        ["CustomerPo"]      = row.CustomerPo,
        ["NetAmount"]       = row.NetAmount,
        ["GrossAmount"]     = row.GrossAmount,
        ["PctPaid"]         = row.PctPaid,
        ["DeliveryDate"]    = row.DeliveryDate?.ToString("o"),
        ["Weight"]          = row.Weight,
        ["Source"]          = source,
    };

    private static DateTimeOffset? ParseSpDate(string? s) =>
        DateTimeOffset.TryParse(s, CultureInfo.InvariantCulture,
            DateTimeStyles.AssumeUniversal | DateTimeStyles.AdjustToUniversal, out var dt) ? dt : null;

    private static double? ParseSpNum(string? s) =>
        double.TryParse(s, NumberStyles.Any, CultureInfo.InvariantCulture, out var d) ? d : null;
}
