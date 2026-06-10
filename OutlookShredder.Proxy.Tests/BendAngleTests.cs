using OutlookShredder.Proxy.Models.Drawing;
using OutlookShredder.Proxy.Services.Drawing;
using Xunit;

namespace OutlookShredder.Proxy.Tests;

// Per-bend angle + fold direction for U / L / Z: the parser captures inline "flange N @ A° up/down"
// and FlatPattern deducts each bend independently. Non-annotated input keeps the old 90°/shape-default
// behaviour (so existing fab notes render unchanged). Contract assertions only.
public class BendAngleTests
{
    [Fact]
    public void Parser_captures_per_flange_angle_and_direction()
    {
        var spec = DrawingTextParser.Parse("U flange 2 @ 75 up flange 3 @ 90 up web 4 length 36, 16ga CRS");
        Assert.True(spec.AnglesAnnotated);
        Assert.NotNull(spec.Bends);
        Assert.Equal(2, spec.Bends!.Count);
        Assert.Equal(75, spec.Bends[0].AngleDeg, 3);
        Assert.Equal(BendDir.Up, spec.Bends[0].Direction);
        Assert.Equal(90, spec.Bends[1].AngleDeg, 3);
        Assert.Equal(BendDir.Up, spec.Bends[1].Direction);
    }

    [Fact]
    public void Parser_reads_opposing_z_directions()
    {
        var spec = DrawingTextParser.Parse("Z flange 2 @ 120 up web 4 flange 2 @ 90 down length 36, 14ga CRS");
        Assert.True(spec.AnglesAnnotated);
        Assert.Equal(BendDir.Up, spec.Bends![0].Direction);
        Assert.Equal(120, spec.Bends[0].AngleDeg, 3);
        Assert.Equal(BendDir.Down, spec.Bends[1].Direction);
    }

    [Fact]
    public void Shorthand_has_no_bend_annotation()
    {
        var spec = DrawingTextParser.Parse("U 4 x 2 x 36, 16ga CRS");
        Assert.False(spec.AnglesAnnotated);
        Assert.Null(spec.Bends);
    }

    [Fact]
    public void FlatWidth_deducts_each_bend_independently()
    {
        var spec = DrawingTextParser.Parse("U flange 2 @ 75 up flange 3 @ 90 up web 4 length 36, 16ga CRS");
        var fp = FlatPattern.Develop(spec);

        double t = spec.Thickness, ri = spec.InsideRadius, k = spec.KFactor;
        double bd75 = BendMath.BendDeduction(ri, t, k, 75);
        double bd90 = BendMath.BendDeduction(ri, t, k, 90);
        // Outside-basis flanges 2 + 3 + web 4, minus the two distinct deductions.
        Assert.Equal(2 + 3 + 4 - bd75 - bd90, fp.FlatWidth, 3);
        // A 75° bend deducts less than 90°, so this is wider than an all-90 blank.
        Assert.True(fp.FlatWidth > 2 + 3 + 4 - 2 * bd90);
    }

    [Fact]
    public void NonAnnotated_U_keeps_symmetric_two_bend_blank()
    {
        var spec = DrawingTextParser.Parse("U 4 x 2 x 36, 16ga CRS");   // web 4, flanges 2/2
        var fp = FlatPattern.Develop(spec);
        double bd = BendMath.BendDeduction(spec.InsideRadius, spec.Thickness, spec.KFactor, 90);
        Assert.Equal(2 + 4 + 2 - 2 * bd, fp.FlatWidth, 3);
    }
}
