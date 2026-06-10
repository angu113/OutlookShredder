using OutlookShredder.Proxy.Models.Drawing;
using OutlookShredder.Proxy.Services.Drawing;
using Xunit;

namespace OutlookShredder.Proxy.Tests;

// Returns = an extra lip/hem folded at a flange's free edge (90 or 180, length + ID/OD + up/down).
// They extend the flat blank by the lip's developed length and add a bend line; the web stays put.
public class ReturnsTests
{
    [Fact]
    public void Parses_a_per_flange_return()
    {
        var spec = DrawingTextParser.Parse("U flange 1 return 0.5 @ 90 up flange 1 web 2 length 36, 16ga CRS");
        Assert.NotNull(spec.ReturnLeft);
        Assert.Equal(0.5, spec.ReturnLeft!.Length, 3);
        Assert.Equal(90, spec.ReturnLeft.AngleDeg, 3);
        Assert.Equal(BendDir.Up, spec.ReturnLeft.Direction);
        Assert.Null(spec.ReturnRight);          // second flange has no return clause
    }

    [Fact]
    public void Parses_a_pan_shared_return_hem()
    {
        var spec = DrawingTextParser.Parse("Pan 24 x 18 x 2, 2 long flanges 2 short flanges, 16ga, returns 0.5 id @ 180 up");
        Assert.NotNull(spec.PanReturn);
        Assert.Equal(180, spec.PanReturn!.AngleDeg, 3);
        Assert.Equal(DimBasis.Inside, spec.PanReturn.Basis);
    }

    [Fact]
    public void Return_widens_the_blank_and_adds_a_bend_line()
    {
        var plain = FlatPattern.Develop(DrawingTextParser.Parse("U 2 x 1 x 36, 16ga CRS"));
        var withRet = FlatPattern.Develop(DrawingTextParser.Parse("U flange 1 return 0.5 @ 90 up flange 1 web 2 length 36, 16ga CRS"));
        Assert.True(withRet.FlatWidth > plain.FlatWidth);                 // lip adds developed length
        Assert.Equal(plain.BendLinesX.Length + 1, withRet.BendLinesX.Length); // one extra crease
    }

    [Fact]
    public void Hem_grows_the_blank_not_shrinks_it()
    {
        var plain = FlatPattern.Develop(DrawingTextParser.Parse("U 4 x 2 x 36, 16ga CRS"));
        var hem = FlatPattern.Develop(DrawingTextParser.Parse("U flange 2 return 0.5 @ 180 up flange 2 web 4 length 36, 16ga CRS"));
        // A 0.5" hem ADDS the lip + the 180° arc to the flat blank — it must be wider than the plain U,
        // and by roughly the lip length (not nearly zero, the old over-deduction bug).
        Assert.True(hem.FlatWidth > plain.FlatWidth, $"hem {hem.FlatWidth} should exceed plain {plain.FlatWidth}");
        Assert.True(hem.FlatWidth - plain.FlatWidth > 0.4, $"hem added only {hem.FlatWidth - plain.FlatWidth:0.###}");
    }

    [Fact]
    public void NoReturn_is_unchanged()
    {
        var spec = DrawingTextParser.Parse("U 4 x 2 x 36, 16ga CRS");
        Assert.Null(spec.ReturnLeft);
        Assert.Null(spec.ReturnRight);
    }

    [Fact]
    public void Three_side_pan_return_adds_three_creases()
    {
        // "2 long flanges 1 short flange" → bottom + top + left walls (3 sides).
        var fp = FlatPattern.Develop(DrawingTextParser.Parse(
            "Pan 24 x 18 x 2, 2 long flanges 1 short flange, 16ga, returns 0.5 @ 90 up"));
        // 3 wall bend lines + 3 return creases (one per present wall).
        Assert.Equal(6, fp.Cut.Entities.Count(e => e.Type == "line"));
    }

    [Fact]
    public void Pan_return_adds_four_crease_bend_lines()
    {
        var plain = FlatPattern.Develop(DrawingTextParser.Parse("Pan 24 x 18 x 2, 2 long flanges 2 short flanges, 16ga"));
        var withRet = FlatPattern.Develop(DrawingTextParser.Parse("Pan 24 x 18 x 2, 2 long flanges 2 short flanges, 16ga, returns 0.5 @ 90 up"));
        int Lines(FlatPatternResult fp) => fp.Cut.Entities.Count(e => e.Type == "line");
        Assert.Equal(Lines(plain) + 4, Lines(withRet));   // one return crease per wall
        // The lip pushes the outline beyond the base — at least one vertex goes negative.
        var verts = withRet.Cut.Entities.Where(e => e.Vertices is not null).SelectMany(e => e.Vertices!);
        Assert.Contains(verts, v => v.X < 0 || v.Y < 0);
    }
}
