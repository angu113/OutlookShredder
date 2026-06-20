namespace OutlookShredder.Proxy.Models;

/// <summary>One persisted/served row of the <c>SalesOrderHistory</c> SP list — one ERP order
/// (<see cref="OrderId"/> = the <c>Doc #</c>, e.g. <c>HSK-SO1036200</c>). Built from the bulk Sales-Orders
/// export (<see cref="Source"/> = <c>export</c>) or appended as SalesOrder ERP docs arrive
/// (<see cref="Source"/> = <c>erp-doc</c>, sparser — only Gross/customer/date populated).</summary>
public sealed class SalesOrderRecord
{
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
