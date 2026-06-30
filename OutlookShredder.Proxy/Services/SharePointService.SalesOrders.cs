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
                r.QueryParameters.Top    = 5000;   // fewer serial nextLink round-trips on the ~35k-row dedup read
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
                // Max page size: ~35k rows at Top=999 meant ~35 SERIAL nextLink round-trips (~34s measured —
                // the slowest startup warm). 5000/page (proven by the SR-cache read) cuts that to ~7 pages.
                // Graph clamps to its own ceiling and the nextLink loop stays correct if it returns fewer.
                r.QueryParameters.Top = 5000;
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
                    SpItemId        = item.Id,
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

    /// <summary>POSTs a new SalesOrderHistory item (returns the new item id), retrying on SharePoint
    /// throttling (429/503/504) with exponential backoff (mirrors <c>PatchListItemWithRetryAsync</c>).</summary>
    private async Task<string?> PostSalesOrderWithRetryAsync(
        string siteId, string listId, Dictionary<string, object?> fields, CancellationToken ct, int maxAttempts = 5)
    {
        for (int attempt = 1; ; attempt++)
        {
            try
            {
                var created = await GetGraph().Sites[siteId].Lists[listId].Items
                    .PostAsync(new Microsoft.Graph.Models.ListItem
                    {
                        Fields = new Microsoft.Graph.Models.FieldValueSet { AdditionalData = fields }
                    }, cancellationToken: ct);
                return created?.Id;
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

    /// <summary>PATCHes the mutable fields of an existing SalesOrderHistory item (a delta load updating a
    /// changed order), retrying on throttling like the POST path.</summary>
    private async Task PatchSalesOrderWithRetryAsync(
        string siteId, string listId, string itemId, Dictionary<string, object?> fields, CancellationToken ct, int maxAttempts = 5)
    {
        for (int attempt = 1; ; attempt++)
        {
            try
            {
                await GetGraph().Sites[siteId].Lists[listId].Items[itemId].Fields
                    .PatchAsync(new Microsoft.Graph.Models.FieldValueSet { AdditionalData = fields }, cancellationToken: ct);
                return;
            }
            catch (Microsoft.Graph.Models.ODataErrors.ODataError ex)
                when (attempt < maxAttempts && ex.ResponseStatusCode is 429 or 503 or 504)
            {
                var delay = TimeSpan.FromSeconds(Math.Pow(2, attempt));
                _log.LogWarning("[SP] SalesOrderHistory PATCH throttled (HTTP {Code}) — retry {A}/{M} in {S}s",
                    ex.ResponseStatusCode, attempt, maxAttempts, delay.TotalSeconds);
                await Task.Delay(delay, ct);
            }
        }
    }

    /// <summary>Delta load: upserts the parsed export against what's already on the list. Reads the existing
    /// rows once (the authoritative new-vs-changed comparison + SP item ids for the PATCH), then in parallel
    /// INSERTs orders not yet present and PATCHes orders whose tracked fields changed since the last load;
    /// unchanged orders are skipped (no write). Returns the delta counts plus the full merged row set so the
    /// caller can rebuild its serving cache without a second SP read. This is the "delta marker" the manual
    /// Sales-Orders load uses — only new + changed rows hit SharePoint.</summary>
    public async Task<SalesOrderUpsertResult> UpsertSalesOrdersAsync(
        IReadOnlyList<CustomerImportService.SalesOrderRow> rows,
        Action<int>? onProgress = null, CancellationToken ct = default)
    {
        var siteId = await GetSiteIdAsync();
        var listId = await GetSalesOrderHistoryListIdAsync();

        var existing = await ReadAllSalesOrdersAsync(ct);
        var byId = new Dictionary<string, SalesOrderRecord>(StringComparer.OrdinalIgnoreCase);
        foreach (var r in existing) byId.TryAdd(r.OrderId, r);

        var toInsert = new List<CustomerImportService.SalesOrderRow>();
        var toPatch  = new List<(SalesOrderRecord Existing, CustomerImportService.SalesOrderRow Row)>();
        int unchanged = 0;
        foreach (var row in rows)
        {
            if (string.IsNullOrWhiteSpace(row.OrderId)) continue;
            if (!byId.TryGetValue(row.OrderId, out var ex)) toInsert.Add(row);
            else if (SalesOrderChanged(ex, row))           toPatch.Add((ex, row));
            else                                            unchanged++;
        }
        _log.LogInformation("[SP] SalesOrderHistory delta: {New} new, {Changed} changed, {Same} unchanged (of {Total} parsed, {Existing} on file)",
            toInsert.Count, toPatch.Count, unchanged, rows.Count, existing.Count);

        int added = 0, changed = 0, failed = 0, done = 0;
        var inserted = new System.Collections.Concurrent.ConcurrentBag<SalesOrderRecord>();
        var opts = new ParallelOptions { MaxDegreeOfParallelism = 4, CancellationToken = ct };

        await Parallel.ForEachAsync(toInsert, opts, async (row, token) =>
        {
            using var rowCts = CancellationTokenSource.CreateLinkedTokenSource(token);
            rowCts.CancelAfter(TimeSpan.FromSeconds(45));
            try
            {
                var id = await PostSalesOrderWithRetryAsync(siteId, listId, BuildSalesOrderFields(row, "export"), rowCts.Token);
                inserted.Add(ToRecord(row, id));
                Interlocked.Increment(ref added);
            }
            catch (Exception ex) when (!token.IsCancellationRequested)
            { _log.LogWarning("[SP] SalesOrderHistory insert failed for {Order}: {Err}", row.OrderId, ex.Message); Interlocked.Increment(ref failed); }
            finally { var n = Interlocked.Increment(ref done); if (n % 250 == 0) onProgress?.Invoke(n); }
        });

        await Parallel.ForEachAsync(toPatch, opts, async (pair, token) =>
        {
            using var rowCts = CancellationTokenSource.CreateLinkedTokenSource(token);
            rowCts.CancelAfter(TimeSpan.FromSeconds(45));
            try
            {
                await PatchSalesOrderWithRetryAsync(siteId, listId, pair.Existing.SpItemId!, BuildSalesOrderFields(pair.Row, pair.Existing.Source ?? "export"), rowCts.Token);
                ApplyRow(pair.Existing, pair.Row);   // reflect the change in the merged set (in-place on the byId record)
                Interlocked.Increment(ref changed);
            }
            catch (Exception ex) when (!token.IsCancellationRequested)
            { _log.LogWarning("[SP] SalesOrderHistory patch failed for {Order}: {Err}", pair.Row.OrderId, ex.Message); Interlocked.Increment(ref failed); }
            finally { var n = Interlocked.Increment(ref done); if (n % 250 == 0) onProgress?.Invoke(n); }
        });

        var merged = byId.Values.Concat(inserted).ToList();
        _log.LogInformation("[SP] SalesOrderHistory delta complete: {Added} added, {Changed} changed, {Same} unchanged, {Failed} failed",
            added, changed, unchanged, failed);
        return new SalesOrderUpsertResult(added, changed, unchanged, failed, merged);
    }

    // Tracked-field comparison for the delta: any difference in the export-carried fields = a changed order.
    private static bool SalesOrderChanged(SalesOrderRecord ex, CustomerImportService.SalesOrderRow row) =>
        !StrEq(ex.CustomerName, row.CustomerName) || ex.OrderDate != row.OrderDate ||
        !StrEq(ex.Status, row.Status) || !StrEq(ex.SecondaryStatus, row.SecondaryStatus) ||
        !StrEq(ex.CustomerPo, row.CustomerPo) || ex.DeliveryDate != row.DeliveryDate ||
        !NumEq(ex.NetAmount, row.NetAmount) || !NumEq(ex.GrossAmount, row.GrossAmount) ||
        !NumEq(ex.PctPaid, row.PctPaid) || !NumEq(ex.Weight, row.Weight);

    private static bool StrEq(string? a, string? b) => string.Equals(a ?? "", b ?? "", StringComparison.Ordinal);
    private static bool NumEq(double? a, double? b) =>
        (a is null && b is null) || (a is not null && b is not null && Math.Abs(a.Value - b.Value) < 0.005);

    private static void ApplyRow(SalesOrderRecord ex, CustomerImportService.SalesOrderRow row)
    {
        ex.CustomerName = row.CustomerName; ex.OrderDate = row.OrderDate;
        ex.Status = row.Status; ex.SecondaryStatus = row.SecondaryStatus; ex.CustomerPo = row.CustomerPo;
        ex.NetAmount = row.NetAmount; ex.GrossAmount = row.GrossAmount; ex.PctPaid = row.PctPaid;
        ex.DeliveryDate = row.DeliveryDate; ex.Weight = row.Weight;
    }

    private static SalesOrderRecord ToRecord(CustomerImportService.SalesOrderRow row, string? spItemId) => new()
    {
        SpItemId = spItemId, OrderId = row.OrderId, CustomerName = row.CustomerName, OrderDate = row.OrderDate,
        Status = row.Status, SecondaryStatus = row.SecondaryStatus, CustomerPo = row.CustomerPo,
        NetAmount = row.NetAmount, GrossAmount = row.GrossAmount, PctPaid = row.PctPaid,
        DeliveryDate = row.DeliveryDate, Weight = row.Weight, Source = "export",
    };

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
