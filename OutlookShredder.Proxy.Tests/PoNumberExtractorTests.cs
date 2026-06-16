using OutlookShredder.Proxy.Services;
using Xunit;

namespace OutlookShredder.Proxy.Tests;

// Behaviour tests for reading the HSK-PO document number from a "Purchase Order - Material" body
// when the filename has none. Sample text mirrors PdfPig's concatenated output from a real PO PDF.
public class PoNumberExtractorTests
{
    [Fact]
    public void Reads_po_number_printed_under_the_title()
    {
        var text = "PURCHASE ORDER - MATERIALHSK-PO1006243PurchaserShip toRequested by:Attention: Purchasing";
        Assert.Equal("HSK-PO1006243", PoNumberExtractor.FromText(text));
    }

    [Fact]
    public void Normalizes_separator_and_spacing_to_canonical_form()
    {
        Assert.Equal("HSK-PO1006243", PoNumberExtractor.FromText("HSK PO 1006243"));
        Assert.Equal("HSK-PO1006243", PoNumberExtractor.FromText("HSKPO1006243"));
    }

    [Fact]
    public void Returns_the_po_number_even_when_a_sales_order_is_also_present()
    {
        // Real PO body carries both the PO# (header) and the customer SO# (line column).
        var text = "PURCHASE ORDER - MATERIALHSK-PO1006243 ... Sales Order #HSK-SO1036172 GWB Admin Building";
        Assert.Equal("HSK-PO1006243", PoNumberExtractor.FromText(text));
    }

    [Fact]
    public void Does_not_mistake_a_sales_order_for_a_po_number()
        => Assert.Null(PoNumberExtractor.FromText("Sales Order #HSK-SO1036172"));

    [Fact]
    public void Returns_null_when_no_po_number_present()
        => Assert.Null(PoNumberExtractor.FromText("PICKING SLIP HSK-SO1036172 some other text"));

    [Fact]
    public void Empty_or_null_input_yields_null()
    {
        Assert.Null(PoNumberExtractor.FromText(null));
        Assert.Null(PoNumberExtractor.FromText(""));
    }
}
