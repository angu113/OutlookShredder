using System.Linq;
using OutlookShredder.Proxy.Services.Drawing;
using Xunit;

namespace OutlookShredder.Proxy.Tests;

public class StrokeFontTests
{
    [Fact]
    public void Known_characters_produce_strokes_within_the_cell()
    {
        var strokes = StrokeFont.Strokes("HRS 0.625\"").ToList();
        Assert.NotEmpty(strokes);
        var pts = strokes.SelectMany(s => s).ToList();
        Assert.All(pts, p => Assert.InRange(p.Y, 0.0, StrokeFont.CapH));   // baseline..cap
        Assert.All(pts, p => Assert.True(p.X >= 0));
    }

    [Fact]
    public void Lower_case_is_drawn_as_upper_case()
    {
        Assert.Equal(StrokeFont.Strokes("HRS").Count(), StrokeFont.Strokes("hrs").Count());
    }

    [Fact]
    public void Spaces_and_unknown_characters_draw_nothing_but_still_advance()
    {
        Assert.Empty(StrokeFont.Strokes("   ").ToList());          // spaces: pen advances, no strokes
        Assert.Empty(StrokeFont.Strokes("§").ToList());            // unknown glyph: skipped
        Assert.True(StrokeFont.WidthUnits("A A") > StrokeFont.WidthUnits("AA"));   // the space takes width
    }

    [Fact]
    public void Width_grows_with_length()
    {
        Assert.True(StrokeFont.WidthUnits("HRS 0.625\"") > StrokeFont.WidthUnits("X2"));
    }
}
