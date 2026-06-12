using OutlookShredder.Proxy.Services.Drawing;
using Xunit;

namespace OutlookShredder.Proxy.Tests;

// Shop-facing formatting: fractional inches (to 1/16, decimal below that), gauge thickness for
// steel/SS and decimal for aluminium, the descriptive gauge-aware title, and that the whole
// drawing renders without throwing across the shapes (incl. the angled + return cases the callout
// geometry rework targets).
public class DrawFormatTests
{
    [Theory]
    [InlineData(0.5, "1/2\"")]
    [InlineData(0.75, "3/4\"")]
    [InlineData(2.0, "2\"")]
    [InlineData(4.25, "4-1/4\"")]
    [InlineData(2.1875, "2-3/16\"")]
    [InlineData(0.0625, "1/16\"")]
    [InlineData(0.1, "0.1\"")]        // NOT a sixteenth -> decimal, never rounded to 1/8
    [InlineData(9.063, "9.063\"")]    // developed blank -> decimal
    [InlineData(0.05, "0.05\"")]      // below 1/16 -> decimal inches
    [InlineData(0.0, "0\"")]
    public void FracInch_is_a_fraction_only_when_exactly_a_sixteenth(double v, string expected)
        => Assert.Equal(expected, DrawFormat.FracInch(v));

    [Theory]
    [InlineData("CRS", 0.0598, "16 ga")]    // steel gauge (MSG 16ga)
    [InlineData("SS", 0.125, "11 ga")]      // stainless 11ga = 0.125 exactly
    [InlineData("CRS", 0.25, "1/4\"")]      // thick plate -> fraction fallback
    [InlineData("alum", 0.063, "0.063\"")]  // aluminium -> decimal inches
    public void ThicknessLabel_prefers_gauge_for_steel_decimal_for_alum(string token, double t, string expected)
        => Assert.Equal(expected, DrawFormat.ThicknessLabel(token, t));

    [Fact]
    public void NearestGauge_snaps_a_near_decimal_and_returns_null_above_the_range()
    {
        Assert.Equal(11, GaugeTables.NearestGauge(MaterialFamily.ColdRolled, 0.12));  // 0.12 -> 11ga (0.1196)
        Assert.Null(GaugeTables.NearestGauge(MaterialFamily.ColdRolled, 0.25));       // plate -> no gauge
    }

    [Fact]
    public void Title_is_descriptive_with_gauge_and_fractional_dims()
    {
        var title = FlatPattern.Develop(DrawingTextParser.Parse("U 4 x 2 x 36, 11ga CRS")).Title;
        Assert.Contains("4\" Web", title);
        Assert.Contains("2\" Flanges", title);
        Assert.Contains("11 ga", title);
        Assert.Contains("36\" long", title);
    }

    [Theory]
    [InlineData("U 4 x 2 x 36, 0.375 steel")]   // thick stock — where the old centreline approximation drifted
    [InlineData("U 4 x 2 x 36, 16ga CRS")]
    [InlineData("L 3 x 2 x 36, 0.25 steel")]
    [InlineData("Z flange 2 @ 90 up web 4 flange 2 @ 90 down length 36, 16ga CRS")]   // Z web (opposing bends)
    [InlineData("U flange 1 return 0.5 @ 90 up flange 1 web 2 length 36, 16ga CRS")]  // web with return flanges
    public void CrossSectionDims_anchors_span_the_dim_value(string text)
    {
        var fp = FlatPattern.Develop(DrawingTextParser.Parse(text));
        var dims = DrawingPdfRenderer.ComputeCrossSectionDims(fp);
        Assert.NotEmpty(dims);
        foreach (var d in dims)
        {
            double span = Math.Sqrt((d.X2 - d.X1) * (d.X2 - d.X1) + (d.Y2 - d.Y1) * (d.Y2 - d.Y1));
            Assert.Equal(d.Value, span, 1);   // the two corner anchors are exactly Value apart
        }
    }

    [Theory]
    [InlineData("U 4 x 2 x 36, 16ga CRS")]
    [InlineData("U flange 2 @ 75 up flange 3 @ 90 up web 4 length 36, 16ga CRS")]            // angled flanges
    [InlineData("Z flange 2 @ 120 up web 4 flange 2 @ 90 down length 36, 14ga CRS")]
    [InlineData("U flange 1 return 0.5 @ 90 up flange 1 web 2 length 36, 16ga CRS")]         // return lip
    [InlineData("L 3 x 2 x 36, 14ga CRS")]
    [InlineData("Pan 24 x 18 x 2, 2 long flanges 2 short flanges, 16ga")]
    [InlineData("Flitch 48 x 6, 0.25 steel")]
    [InlineData("U 4 x 2 x 36, 0.06 alum")]
    public void Render_produces_a_pdf_without_throwing(string text)
    {
        var fp = FlatPattern.Develop(DrawingTextParser.Parse(text));
        var pdf = DrawingPdfRenderer.Render(fp);
        Assert.True(pdf.Length > 0);
    }
}
