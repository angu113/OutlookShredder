using Microsoft.Extensions.Logging.Abstractions;
using OutlookShredder.Proxy.Models;
using OutlookShredder.Proxy.Services;
using Xunit;

namespace OutlookShredder.Proxy.Tests;

// Dedup contract: rows that are the same requested item (same MSPC + line) collapse to the
// single richest row; genuinely distinct items are all kept. Asserts observable in/out only.
public class ProductDeduplicatorTests
{
    private static ProductLine Line(string mspc, int lineNo, string name, string comments) => new()
    {
        ProductSearchKey        = mspc,
        LineNumber              = lineNo,
        ProductName             = name,
        UnitsQuoted             = 100,
        TotalPrice              = 500,
        SupplierProductComments = comments,
    };

    [Fact]
    public void Same_mspc_and_line_collapses_to_the_richest_row()
    {
        var rows = new List<ProductLine>
        {
            Line("AF6061/2502500", 1, "Aluminum Flat Bar 6061T6511 0.250 X 2.500", "Full temper and finish notes"),
            Line("AF6061/2502500", 1, "Aluminum Flat Bar", "T6511"),
            Line("AF6061/2502500", 1, "6061", ""),
        };
        var result = ProductDeduplicator.Deduplicate(rows, "test", dryRun: false, NullLogger.Instance);
        Assert.Single(result);
        Assert.Equal("Aluminum Flat Bar 6061T6511 0.250 X 2.500", result[0].ProductName);  // richest survives
    }

    [Fact]
    public void Distinct_mspcs_are_both_kept()
    {
        var rows = new List<ProductLine>
        {
            Line("AF6061/2502500", 1, "Aluminum Flat Bar", "a"),
            Line("RB304/1000000",  1, "Stainless Round Bar", "b"),
        };
        var result = ProductDeduplicator.Deduplicate(rows, "test", dryRun: false, NullLogger.Instance);
        Assert.Equal(2, result.Count);
    }

    [Fact]
    public void A_single_row_is_returned_unchanged()
    {
        var rows = new List<ProductLine> { Line("AF6061/2502500", 1, "Aluminum Flat Bar", "x") };
        Assert.Single(ProductDeduplicator.Deduplicate(rows, "test", dryRun: false, NullLogger.Instance));
    }
}
