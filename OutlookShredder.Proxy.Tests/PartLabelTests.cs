using System.Linq;
using OutlookShredder.Proxy.Services.Drawing;
using Xunit;

namespace OutlookShredder.Proxy.Tests;

// PartLabel draws the shop cutting-aid (qty / material / thickness) as single-stroke polylines on the
// no-cut white "Notes" layer, centered above the part, starting 1" tall and shrinking to fit the width.
public class PartLabelTests
{
    private static CutGeometry Rect(double w, double h)
    {
        var geo = new CutGeometry();
        geo.Layers.Add(new CutLayer { Name = FlatPattern.CutLayer, Color = FlatPattern.CutColor });
        geo.Entities.Add(CutEntity.Polyline(FlatPattern.CutLayer, true, new[]
        {
            new CutVertex(0, 0), new CutVertex(w, 0), new CutVertex(w, h), new CutVertex(0, h),
        }));
        return geo;
    }

    private static List<CutEntity> Label(CutGeometry g) =>
        g.Entities.Where(e => e.Layer == PartLabel.LayerName).ToList();

    // ── Content (BuildLines) ──────────────────────────────────────────────────

    [Fact]
    public void Quantity_is_shown_when_supplied_including_one()
    {
        Assert.Equal(new[] { "X2", "HRS 0.625\"" }, PartLabel.BuildLines(2, "HRS", 0.625));
        Assert.Equal(new[] { "X1", "HRS 0.625\"" }, PartLabel.BuildLines(1, "HRS", 0.625));
    }

    [Fact]
    public void No_quantity_omits_the_xN_line_and_material_is_upper_cased()
    {
        Assert.Equal(new[] { "GALV 0.075\"" }, PartLabel.BuildLines(null, "galv", 0.075));
    }

    [Fact]
    public void Nothing_to_say_is_empty()
    {
        Assert.Empty(PartLabel.BuildLines(null, "", 0));
    }

    // ── Font fit ──────────────────────────────────────────────────────────────

    [Fact]
    public void Wide_part_keeps_the_full_three_quarter_inch_font()
    {
        Assert.Equal(0.75, PartLabel.ChooseHeight(100, new[] { "X2", "HRS 0.625\"" }), 6);
    }

    [Fact]
    public void Narrow_part_shrinks_the_font_to_fit_its_width()
    {
        var lines = new[] { "X2", "HRS 0.625\"" };
        double h = PartLabel.ChooseHeight(3.0, lines);
        Assert.True(h < 0.75, $"narrow part should shrink the font, got {h}");

        // The widest line, at the chosen height, stays within the 3" part width.
        double widestIn = lines.Max(StrokeFontWidth) * h / StrokeFont.CapH;
        Assert.True(widestIn <= 3.0 + 1e-6, $"label width {widestIn} exceeds the 3\" part");
    }

    private static double StrokeFontWidth(string s) => StrokeFont.WidthUnits(s);

    // ── Geometry placement ────────────────────────────────────────────────────

    [Fact]
    public void Label_is_stroked_polylines_on_the_L1_layer_above_and_centered()
    {
        var geo = Rect(20, 6);
        PartLabel.AddTo(geo, 4, "CRS", 0.075);

        var label = Label(geo);
        Assert.NotEmpty(label);
        Assert.All(label, e => Assert.Equal("polyline", e.Type));          // geometry, not text
        Assert.Contains(geo.Layers, l => l.Name == PartLabel.LayerName && l.Color == PartLabel.LayerColor);

        var pts = label.SelectMany(e => e.Vertices!).ToList();
        Assert.All(pts, p => Assert.True(p.Y >= 6.0 - 1e-6, "label sits above the part top (y=6)"));
        // Centered on the 20-wide part (within half a glyph cell — narrow trailing glyphs shift the ink
        // centre slightly off the advance-box centre, which is expected).
        double cx = (pts.Min(p => p.X) + pts.Max(p => p.X)) / 2.0;
        Assert.InRange(cx, 9.5, 10.5);
    }

    [Fact]
    public void Adding_the_quantity_line_adds_more_strokes()
    {
        var withQty = Rect(40, 6); PartLabel.AddTo(withQty, 3, "HRS", 0.625);
        var noQty   = Rect(40, 6); PartLabel.AddTo(noQty,  null, "HRS", 0.625);
        Assert.True(Label(withQty).Count > Label(noQty).Count, "the xN line should add strokes");
    }

    [Fact]
    public void Nothing_to_say_adds_no_geometry()
    {
        var geo = Rect(10, 5);
        PartLabel.AddTo(geo, null, "", 0);
        Assert.Empty(Label(geo));
    }
}
