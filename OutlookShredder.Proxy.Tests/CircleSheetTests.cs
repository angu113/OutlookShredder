using System.Linq;
using OutlookShredder.Proxy.Models.Drawing;
using OutlookShredder.Proxy.Services.Drawing;
using Xunit;

namespace OutlookShredder.Proxy.Tests;

// Circle/disc (+ donut) and plain Sheet — the two flat (no-bend) shapes added 2026-06.
// Circle: diameter (+ optional inner diameter = annulus), NO polish (a round blank has no grain axis).
// Sheet: width x height, carries polish. Both develop to a flat-pattern cut + render a PDF.
// Changing any of this behaviour must update this test in the same commit.
public class CircleSheetTests
{
    private static int Circles(FlatPatternResult fp) =>
        fp.Cut.Entities.Count(e => e.Type == "circle" && e.Layer != PartLabel.LayerName);

    // ── Circle parsing ─────────────────────────────────────────────────────────
    [Fact]
    public void Circle_parses_diameter()
    {
        var s = DrawingTextParser.Parse("Circle dia 12, 16ga steel");
        Assert.Equal(PartType.Circle, s.Type);
        Assert.Equal(12, s.Diameter, 3);
        Assert.Equal(0, s.InnerDiameter, 3);
    }

    [Fact]
    public void Circle_shorthand_without_the_dia_keyword()
        => Assert.Equal(12, DrawingTextParser.Parse("Circle 12, 16ga steel").Diameter, 3);

    [Fact]
    public void Donut_parses_inner_diameter()
    {
        var s = DrawingTextParser.Parse("Circle dia 12 id 8, 16ga steel");
        Assert.Equal(12, s.Diameter, 3);
        Assert.Equal(8, s.InnerDiameter, 3);
    }

    [Fact]
    public void Donut_inner_must_be_smaller_than_outer()
        => Assert.Throws<DrawingParseException>(() => DrawingTextParser.Parse("Circle dia 8 id 8, 16ga steel"));

    [Fact]
    public void Circle_carries_no_polish_even_if_a_token_leaks_in()   // a disc has no single grain axis
        => Assert.Equal(PolishDirection.None,
            DrawingTextParser.Parse("Circle dia 12, 16ga steel, polish vertical").PolishDirection);

    // ── Sheet parsing ──────────────────────────────────────────────────────────
    [Fact]
    public void Sheet_parses_width_and_height()
    {
        var s = DrawingTextParser.Parse("Sheet 24 x 12, 16ga steel");
        Assert.Equal(PartType.Sheet, s.Type);
        Assert.Equal(24, s.Width, 3);
        Assert.Equal(12, s.Height, 3);
    }

    [Fact]
    public void Sheet_carries_polish()
        => Assert.Equal(PolishDirection.Vertical,
            DrawingTextParser.Parse("Sheet 24 x 12, 16ga steel, polish vertical").PolishDirection);

    // ── Develop (flat pattern + cut geometry) ──────────────────────────────────
    [Fact]
    public void Circle_develops_to_a_single_disc()
    {
        var fp = FlatPattern.Develop(DrawingTextParser.Parse("Circle dia 12, 16ga steel"));
        Assert.True(fp.IsCircle);
        Assert.Equal(12, fp.FlatWidth, 3);
        Assert.Equal(12, fp.FlatHeight, 3);
        Assert.Equal(1, Circles(fp));
        Assert.Contains(fp.Cut.Entities, e => e.Type == "circle" && System.Math.Abs(e.R - 6) < 1e-6);
    }

    [Fact]
    public void Donut_cuts_two_concentric_circles()
    {
        var fp = FlatPattern.Develop(DrawingTextParser.Parse("Circle dia 12 id 8, 16ga steel"));
        Assert.Equal(2, Circles(fp));
        Assert.Contains(fp.Cut.Entities, e => e.Type == "circle" && System.Math.Abs(e.R - 6) < 1e-6);   // OD
        Assert.Contains(fp.Cut.Entities, e => e.Type == "circle" && System.Math.Abs(e.R - 4) < 1e-6);   // ID bore
    }

    [Fact]
    public void Sheet_develops_to_a_flat_rectangle()
    {
        var fp = FlatPattern.Develop(DrawingTextParser.Parse("Sheet 24 x 12, 16ga steel"));
        Assert.True(fp.IsPlate);                 // reuses the flat-plate top-view renderer
        Assert.Equal(24, fp.FlatWidth, 3);
        Assert.Equal(12, fp.FlatHeight, 3);
        Assert.Empty(fp.Holes);
        Assert.Contains(fp.Cut.Entities, e => e.Type == "polyline" && e.Vertices is { Count: 4 });
    }

    // ── End-to-end: a real PDF renders for each (incl. polish on the sheet) ────
    [Theory]
    [InlineData("Circle dia 12, 16ga steel")]
    [InlineData("Circle dia 12 id 8, 16ga steel")]
    [InlineData("Sheet 24 x 12, 16ga steel")]
    [InlineData("Sheet 24 x 12, 16ga steel, polish horizontal")]
    public void Pdf_renders(string text)
    {
        var pdf = DrawingPdfRenderer.Render(FlatPattern.Develop(DrawingTextParser.Parse(text)));
        Assert.NotNull(pdf);
        Assert.True(pdf.Length > 1000, $"expected a real PDF, got {pdf.Length} bytes");
    }
}
