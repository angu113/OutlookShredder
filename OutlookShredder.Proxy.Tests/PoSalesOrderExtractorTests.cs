using OutlookShredder.Proxy.Services;
using Xunit;

namespace OutlookShredder.Proxy.Tests;

// Behaviour tests for the deterministic HSK-SO extractor that captures the customer sales-order
// numbers printed in a procure-to-order PO's "Sales Order #" column. Pure in/out (text -> SO list).
// Sample text mirrors PdfPig's concatenated output from real PO PDFs.
public class PoSalesOrderExtractorTests
{
    [Fact]
    public void Extracts_single_sales_order_from_column_text()
    {
        var text = "Promise DateCustomerSales Order #UomQty06/09/2026HSK-SO1035965-1040320.00IN Safeguard Tech LLC";
        Assert.Equal(new[] { "HSK-SO1035965" }, PoSalesOrderExtractor.FromText(text));
    }

    [Fact]
    public void Extracts_multiple_distinct_sales_orders_in_first_seen_order()
    {
        // One PO consolidating material for two customer orders (real Allegroni/Dobco case).
        var text = "...HSK-SO1034707-1020.00IN Allegroni 05/04/2026 HSK-SO1034940-30380.00IN Dobco";
        Assert.Equal(new[] { "HSK-SO1034707", "HSK-SO1034940" }, PoSalesOrderExtractor.FromText(text));
    }

    [Fact]
    public void Dedups_repeated_references()
    {
        var text = "HSK-SO1035002 line1 ... HSK-SO1035002 line2 ... HSK-SO1035002 line3";
        Assert.Equal(new[] { "HSK-SO1035002" }, PoSalesOrderExtractor.FromText(text));
    }

    [Fact]
    public void Normalizes_separator_and_spacing_to_canonical_form()
    {
        Assert.Equal(new[] { "HSK-SO1035965" }, PoSalesOrderExtractor.FromText("HSK SO 1035965"));
        Assert.Equal(new[] { "HSK-SO1035965" }, PoSalesOrderExtractor.FromText("HSKSO1035965"));
    }

    [Fact]
    public void Returns_empty_for_a_stock_po_with_no_sales_order_column()
    {
        var text = "PURCHASE ORDER - MATERIAL HSK-PO1004397 Vendor ... Aluminum Sheet 6061T6 ... TOTAL 1,059.00";
        Assert.Empty(PoSalesOrderExtractor.FromText(text));
    }

    [Fact]
    public void Does_not_mistake_a_po_number_for_a_sales_order()
        => Assert.Empty(PoSalesOrderExtractor.FromText("HSK-PO1004397"));

    [Fact]
    public void Empty_or_null_input_yields_empty()
    {
        Assert.Empty(PoSalesOrderExtractor.FromText(null));
        Assert.Empty(PoSalesOrderExtractor.FromText(""));
    }
}
