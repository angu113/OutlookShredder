namespace OutlookShredder.Proxy.Models;

/// <summary>One persisted/served row of the <c>SalesOrderHistory</c> SP list — one ERP order
/// (<see cref="OrderId"/> = the <c>Doc #</c>, e.g. <c>HSK-SO1036200</c>). Built from the bulk Sales-Orders
/// export (<see cref="Source"/> = <c>export</c>) or appended as SalesOrder ERP docs arrive
/// (<see cref="Source"/> = <c>erp-doc</c>, sparser — only Gross/customer/date populated).</summary>
public sealed class SalesOrderRecord
{
    /// <summary>SharePoint list item id — set on rows read back from SP so a delta load can PATCH a changed
    /// order in place. Null on freshly-parsed / appended rows that haven't been read back yet.</summary>
    public string?         SpItemId        { get; set; }
    public string          OrderId         { get; set; } = "";
    public string          CustomerName    { get; set; } = "";
    public DateTimeOffset?  OrderDate       { get; set; }
    public string?          Status          { get; set; }
    public string?          SecondaryStatus { get; set; }
    public string?          CustomerPo      { get; set; }
    public double?          NetAmount       { get; set; }
    public double?          GrossAmount     { get; set; }
    public double?          PctPaid         { get; set; }
    public DateTimeOffset?  DeliveryDate    { get; set; }
    public double?          Weight          { get; set; }
    public string?          Source          { get; set; }

    /// <summary>The amount shown on the call card: Net, falling back to Gross for sparse erp-doc rows.</summary>
    public double? DisplayAmount => NetAmount ?? GrossAmount;
}

/// <summary>Outcome of a delta Sales-Orders load: how many rows were new (inserted), changed (patched in
/// place) or unchanged (skipped), plus failures and the full merged row set so the serving cache can be
/// rebuilt without re-reading SharePoint.</summary>
public sealed record SalesOrderUpsertResult(
    int Added, int Changed, int Unchanged, int Failed, List<SalesOrderRecord> Merged);

/// <summary>On-disk serving cache for SalesOrderHistory (so startup skips the ~29s SP read). Holds the full
/// row set plus the <see cref="LastLoadUtc"/> delta marker — when SharePoint was last fully (re)loaded.</summary>
public sealed class SalesOrderCacheFile
{
    public int                   SchemaVersion { get; set; }
    public DateTime?             LastLoadUtc   { get; set; }
    public List<SalesOrderRecord> Rows         { get; set; } = [];
}

/// <summary>Response for <c>GET /api/sales-orders/by-customer</c> — a customer's recent orders plus a
/// summary header for the Raptor call card.</summary>
public sealed record CustomerOrdersResponse(
    CustomerOrdersSummary           Summary,
    IReadOnlyList<CustomerOrderDto> Orders);

/// <summary>Card header: total order count, lifetime Net $ (Gross fallback per row), and the most recent
/// order date (ISO <c>yyyy-MM-dd</c>, or null when no dated orders).</summary>
public sealed record CustomerOrdersSummary(int OrderCount, double LifetimeNet, string? LastOrderDate);

/// <summary>One order line on the call card. <see cref="Date"/> is ISO <c>yyyy-MM-dd</c> (null when the
/// order has no date); <see cref="Amount"/> is the displayed Net-or-Gross figure.</summary>
public sealed record CustomerOrderDto(
    string  OrderId,
    string? Date,
    double? Amount,
    double? Net,
    double? Gross,
    string? Status);
