using OutlookShredder.Proxy.Services;
using Xunit;

namespace OutlookShredder.Proxy.Tests;

// WeightCalculator is the single source of truth for $/lb. These assert the observable
// contracts (unit conversions, supplier-stated vs estimated weight) — not internal density
// constants, which are tunable. A behaviour change updates the matching test in the same commit.
public class WeightCalculatorTests
{
    [Theory]
    [InlineData(5, "lb", 5)]
    [InlineData(5, null, 5)]    // blank unit = pounds
    [InlineData(128, "oz", 8)]
    public void ToLb_converts_to_pounds(double v, string? unit, double expected)
        => Assert.Equal(expected, WeightCalculator.ToLb(v, unit), 3);

    [Fact]
    public void ToLb_kg_uses_the_standard_factor()
        => Assert.Equal(220.462, WeightCalculator.ToLb(100, "kg"), 3);

    [Theory]
    [InlineData(12, "in", 1)]
    [InlineData(5, "ft", 5)]
    public void ToFeet_converts_to_feet(double v, string unit, double expected)
        => Assert.Equal(expected, WeightCalculator.ToFeet(v, unit)!.Value, 3);

    [Fact]
    public void ToFeet_meter_uses_the_standard_factor()
        => Assert.Equal(3.28084, WeightCalculator.ToFeet(1, "m")!.Value, 4);

    [Fact]
    public void ToFeet_is_null_for_nonpositive_or_missing()
    {
        Assert.Null(WeightCalculator.ToFeet(0, "ft"));
        Assert.Null(WeightCalculator.ToFeet(null, "ft"));
    }

    [Fact]
    public void ResolveLineWeightLb_uses_supplier_weight_exactly_when_present()
    {
        var (lb, est) = WeightCalculator.ResolveLineWeightLb(
            qty: 4, supplierWeightPerUnit: 2.5, supplierWeightUnit: "lb",
            catalogProductName: null, supplierLabel: null, lengthPerUnit: null, lengthUnit: null);
        Assert.Equal(10.0, lb!.Value, 3);
        Assert.False(est);   // supplier-stated weight is exact, not an estimate
    }

    [Fact]
    public void ResolveLineWeightLb_converts_a_supplier_kg_weight()
    {
        var (lb, est) = WeightCalculator.ResolveLineWeightLb(
            qty: 2, supplierWeightPerUnit: 1, supplierWeightUnit: "kg",
            catalogProductName: null, supplierLabel: null, lengthPerUnit: null, lengthUnit: null);
        Assert.Equal(2 * 2.20462, lb!.Value, 3);
        Assert.False(est);
    }

    [Fact]
    public void ResolveLineWeightLb_clamps_nonpositive_qty_to_one()
    {
        var (lb, _) = WeightCalculator.ResolveLineWeightLb(
            qty: 0, supplierWeightPerUnit: 2.5, supplierWeightUnit: "lb",
            catalogProductName: null, supplierLabel: null, lengthPerUnit: null, lengthUnit: null);
        Assert.Equal(2.5, lb!.Value, 3);   // qty 0 -> treated as 1
    }

    [Fact]
    public void ResolveLineWeightLb_is_null_when_there_is_no_basis()
    {
        var (lb, est) = WeightCalculator.ResolveLineWeightLb(
            qty: 2, supplierWeightPerUnit: null, supplierWeightUnit: null,
            catalogProductName: null, supplierLabel: null, lengthPerUnit: null, lengthUnit: null);
        Assert.Null(lb);
        Assert.False(est);
    }

    [Fact]
    public void ResolveLineWeightLb_estimates_from_catalog_dims_and_flags_it()
    {
        // No supplier weight -> theoretical from the catalog product's clean dims, flagged Estimated.
        var (lb, est) = WeightCalculator.ResolveLineWeightLb(
            qty: 2, supplierWeightPerUnit: null, supplierWeightUnit: null,
            catalogProductName: "Aluminum Flat Bar 6061T6511 0.250 X 2.500",
            supplierLabel: null, lengthPerUnit: 10, lengthUnit: "ft");
        Assert.NotNull(lb);
        Assert.True(lb!.Value > 0);
        Assert.True(est);   // theoretical estimate
    }

    [Fact]
    public void Calculate_returns_a_positive_per_foot_weight_for_a_known_flat_bar()
    {
        var r = WeightCalculator.Calculate("Aluminum Flat Bar 6061T6511 0.250 X 2.500");
        Assert.NotNull(r.LbPerFoot);
        Assert.True(r.LbPerFoot!.Value > 0);
    }

    [Fact]
    public void Calculate_returns_no_weight_for_an_empty_name()
        => Assert.Null(WeightCalculator.Calculate("").LbPerFoot);
}
