using OutlookShredder.Proxy.Services;
using Xunit;

namespace OutlookShredder.Proxy.Tests;

// Dimension canonicalisation: fractions, mixed numbers, leading-dot decimals, and sheet gauges
// all resolve to decimal-inch strings so the same product written differently matches. Contract
// assertions only.
public class DimensionNormalizerTests
{
    [Theory]
    [InlineData("3/16", "0.188")]     // simple fraction -> 3-dp inches (0.1875 rounded)
    [InlineData("1 1/2", "1.5")]      // mixed number
    [InlineData(".5", "0.5")]         // leading-dot decimal
    [InlineData("null", "null")]      // non-numeric pass-through (Claude emits literal "null")
    public void NormalizeToken_converts_fractions_and_decimals(string token, string expected)
        => Assert.Equal(expected, DimensionNormalizer.NormalizeToken(token, "steel", tube: false));

    [Fact]
    public void NormalizeToken_resolves_a_sheet_gauge_to_inches()
        => Assert.Equal("0.125", DimensionNormalizer.NormalizeToken("11ga", "stainless", tube: false));

    [Fact]
    public void CanonicalizeDims_normalises_each_axis()
        => Assert.Equal("0.188x48x96",
            DimensionNormalizer.CanonicalizeDims("3/16x48x96", "steel", "sheet"));

    [Fact]
    public void CanonicalizeDims_leaves_already_canonical_dims_unchanged()
        => Assert.Equal("0.188x48x96",
            DimensionNormalizer.CanonicalizeDims("0.188x48x96", "steel", "sheet"));

    [Fact]
    public void CanonicalizeDims_passes_through_null_and_empty()
    {
        Assert.Null(DimensionNormalizer.CanonicalizeDims(null, "steel", "sheet"));
        Assert.Equal("", DimensionNormalizer.CanonicalizeDims("", "steel", "sheet"));
    }

    [Fact]
    public void GaugeToInches_resolves_standard_steel_sheet_gauge()
        => Assert.Equal(0.0598, DimensionNormalizer.GaugeToInches("steel", tube: false, 16)!.Value, 4);
}
