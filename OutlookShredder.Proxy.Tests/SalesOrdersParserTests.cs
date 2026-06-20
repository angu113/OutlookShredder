using System.Text;
using Microsoft.Extensions.Logging.Abstractions;
using OutlookShredder.Proxy.Services;
using Xunit;

namespace OutlookShredder.Proxy.Tests;

// Behaviour tests for the Sales-Orders history parser (CustomerImportService.ParseSalesOrders), which mines
// the orders export ("ExportedData (8).csv" style) into one SalesOrderRow per HSK-SO# Doc #. Pure in/out —
// no SharePoint (dedup-vs-existing and the insert live downstream in SharePointService).
public class SalesOrdersParserTests
{
    private static CustomerImportService NewSvc() => new(NullLogger<CustomerImportService>.Instance);

    // The real export header (column order matches ExportedData (8).csv); the parser resolves by name.
    private const string Header =
        "\"Order Date\",\"Doc #\",\"Customer\",\"Contact\",\"Status\",\"Secondary Status\",\"Customer PO Number\"," +
        "\"% Paid\",\"Net $\",\"Gross $\",\"Delivery Date\",\"Wt.\"";

    // (orderDate, doc#, customer, status, secondary, po, %paid, net, gross, deliveryDate, wt)
    private static string Row(string date, string doc, string cust, string status = "Booked",
        string secondary = "Open Order", string po = "", string paid = "0", string net = "0",
        string gross = "0", string delivery = "", string wt = "0") =>
        $"\"{date}\",\"{doc}\",\"{cust}\",\"973 000 0000 - X -\",\"{status}\",\"{secondary}\",\"{po}\"," +
        $"{paid},{net},{gross},\"{delivery}\",{wt}";

    private static CustomerImportService.ParseResult<CustomerImportService.SalesOrderRow> Parse(params string[] dataRows)
    {
        var sb = new StringBuilder();
        sb.AppendLine(Header);
        foreach (var r in dataRows) sb.AppendLine(r);
        return NewSvc().ParseSalesOrders(sb.ToString());
    }

    [Fact]
    public void Parses_core_fields_from_an_order_row()
    {
        var parsed = Parse(Row("06-19-2026", "HSK-SO1036200", "McEntee Construction",
            status: "Booked", po: "PO-77", paid: "0", net: "3998.54", gross: "4100.00",
            delivery: "06-16-2026", wt: "1241.95484"));

        var row = Assert.Single(parsed.Rows);
        Assert.Equal("HSK-SO1036200",        row.OrderId);
        Assert.Equal("McEntee Construction", row.CustomerName);
        Assert.Equal("Booked",               row.Status);
        Assert.Equal("Open Order",           row.SecondaryStatus);
        Assert.Equal("PO-77",                row.CustomerPo);
        Assert.Equal(3998.54,                row.NetAmount);
        Assert.Equal(4100.00,                row.GrossAmount);
        Assert.Equal(1241.9548,              row.Weight);   // rounded to 4dp
        Assert.NotNull(row.OrderDate);
        Assert.Equal(new DateOnly(2026, 6, 19), DateOnly.FromDateTime(row.OrderDate!.Value.UtcDateTime));
        Assert.Equal(new DateOnly(2026, 6, 16), DateOnly.FromDateTime(row.DeliveryDate!.Value.UtcDateTime));
    }

    [Fact]
    public void Blank_money_and_date_cells_become_null()
    {
        var parsed = Parse(Row("06-19-2026", "HSK-SO1", "Acme", po: "", paid: "", net: "", gross: "", delivery: "", wt: ""));

        var row = Assert.Single(parsed.Rows);
        Assert.Null(row.NetAmount);
        Assert.Null(row.GrossAmount);
        Assert.Null(row.PctPaid);
        Assert.Null(row.Weight);
        Assert.Null(row.CustomerPo);
        Assert.Null(row.DeliveryDate);
    }

    [Fact]
    public void Quotes_are_kept_with_their_status()
    {
        var parsed = Parse(Row("06-19-2026", "HSK-SO1036318", "Dimitri Burlacod", status: "Quoted"));

        var row = Assert.Single(parsed.Rows);
        Assert.Equal("Quoted", row.Status);
    }

    [Fact]
    public void Non_order_and_blank_doc_rows_are_dropped()
    {
        var parsed = Parse(
            Row("06-19-2026", "",            "Acme"),    // blank Doc #
            Row("06-19-2026", "HSK-PO5000",  "Acme"),    // not a sales order
            Row("06-19-2026", "TOTALS",      "Acme"),    // footer
            Row("06-19-2026", "HSK-SO9001",  "Acme"));   // the only keepable row

        var row = Assert.Single(parsed.Rows);
        Assert.Equal("HSK-SO9001", row.OrderId);
    }

    [Fact]
    public void Blank_customer_row_is_dropped()
    {
        var parsed = Parse(Row("06-19-2026", "HSK-SO9002", ""));
        Assert.Empty(parsed.Rows);
    }

    [Fact]
    public void Duplicate_doc_collapses_to_first_and_is_reported()
    {
        var parsed = Parse(
            Row("06-19-2026", "HSK-SO7", "Acme", net: "100"),
            Row("06-20-2026", "HSK-SO7", "Acme", net: "999"));   // same Doc # — dropped

        var row = Assert.Single(parsed.Rows);
        Assert.Equal(100, row.NetAmount);                        // first kept
        Assert.Contains(parsed.Warnings, w => w.Contains("duplicate", System.StringComparison.OrdinalIgnoreCase));
    }

    [Fact]
    public void Missing_required_columns_yields_a_warning_and_no_rows()
    {
        var csv = "\"Order Date\",\"Status\"\n\"06-19-2026\",\"Booked\"\n";

        var parsed = NewSvc().ParseSalesOrders(csv);

        Assert.Empty(parsed.Rows);
        Assert.Contains(parsed.Warnings, w => w.Contains("Doc #") || w.Contains("Customer"));
    }

    [Theory]
    [InlineData("3998.54", 3998.54)]
    [InlineData("1,234.50", 1234.50)]
    [InlineData("$1,200", 1200)]
    [InlineData("50%", 50)]
    [InlineData("0", 0)]
    public void ParseMoney_strips_symbols(string raw, double expected)
        => Assert.Equal(expected, CustomerImportService.ParseMoney(raw));

    [Theory]
    [InlineData("")]
    [InlineData("   ")]
    [InlineData("n/a")]
    public void ParseMoney_blank_or_garbage_is_null(string raw)
        => Assert.Null(CustomerImportService.ParseMoney(raw));

    [Fact]
    public void ParseErpDate_parses_mmddyyyy_anchored_at_noon_utc()
    {
        var d = CustomerImportService.ParseErpDate("06-19-2026");

        Assert.NotNull(d);
        Assert.Equal(new DateOnly(2026, 6, 19), DateOnly.FromDateTime(d!.Value.UtcDateTime));
        Assert.Equal(12, d.Value.UtcDateTime.Hour);           // noon UTC survives US-TZ round-trip
        Assert.Equal(System.TimeSpan.Zero, d.Value.Offset);
    }

    [Theory]
    [InlineData("")]
    [InlineData("not a date")]
    public void ParseErpDate_blank_or_garbage_is_null(string raw)
        => Assert.Null(CustomerImportService.ParseErpDate(raw));
}
