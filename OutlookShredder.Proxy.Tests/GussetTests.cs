using System;
using System.Linq;
using OutlookShredder.Proxy.Models.Drawing;
using OutlookShredder.Proxy.Services.Drawing;
using Xunit;

namespace OutlookShredder.Proxy.Tests;

// Gusset — a flat right-triangle bracket: width (base) x height legs at a right angle, the hypotenuse
// joining their free ends. No bends; carries polish (a triangle has a grain axis). Develops to a
// 3-vertex closed cut polyline and renders a single face view (base + height + hyp callout).
// Changing any of this behaviour must update this test in the same commit.
public class GussetTests
{
    [Fact]
    public void Gusset_parses_width_and_height()
    {
        var s = DrawingTextParser.Parse("Gusset 12 x 8, 0.25 HRS");
        Assert.Equal(PartType.Gusset, s.Type);
        Assert.Equal(12, s.Width, 3);
        Assert.Equal(8, s.Height, 3);
    }

    [Fact]
    public void Triangle_keyword_also_parses_as_a_gusset()
        => Assert.Equal(PartType.Gusset, DrawingTextParser.Parse("Triangle 10 x 10, 0.25 HRS").Type);

    [Fact]
    public void Gusset_carries_polish()
        => Assert.Equal(PolishDirection.Vertical,
            DrawingTextParser.Parse("Gusset 12 x 8, 0.25 HRS, polish vertical").PolishDirection);

    [Fact]
    public void Gusset_rejects_missing_dimensions()
        => Assert.Throws<DrawingParseException>(() => DrawingTextParser.Parse("Gusset, 0.25 HRS"));

    [Fact]
    public void Gusset_develops_to_a_right_triangle()
    {
        var fp = FlatPattern.Develop(DrawingTextParser.Parse("Gusset 12 x 8, 0.25 HRS"));
        Assert.True(fp.IsGusset);
        Assert.Equal(12, fp.FlatWidth, 3);
        Assert.Equal(8, fp.FlatHeight, 3);
        // One closed cut polyline with the three triangle vertices (right angle at the origin).
        var tri = Assert.Single(fp.Cut.Entities, e => e.Type == "polyline" && e.Layer != PartLabel.LayerName);
        Assert.Equal(3, tri.Vertices!.Count);
        Assert.Contains(tri.Vertices, v => Math.Abs(v.X) < 1e-6 && Math.Abs(v.Y) < 1e-6);          // (0,0)
        Assert.Contains(tri.Vertices, v => Math.Abs(v.X - 12) < 1e-6 && Math.Abs(v.Y) < 1e-6);     // (W,0)
        Assert.Contains(tri.Vertices, v => Math.Abs(v.X) < 1e-6 && Math.Abs(v.Y - 8) < 1e-6);      // (0,H)
    }

    [Theory]
    [InlineData("Gusset 12 x 8, 0.25 HRS")]
    [InlineData("Gusset 10 x 10, 0.25 HRS, polish horizontal")]
    public void Pdf_renders(string text)
    {
        var pdf = DrawingPdfRenderer.Render(FlatPattern.Develop(DrawingTextParser.Parse(text)));
        Assert.NotNull(pdf);
        Assert.True(pdf.Length > 1000, $"expected a real PDF, got {pdf.Length} bytes");
    }
}
