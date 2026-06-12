using System.Linq;
using OutlookShredder.Proxy.Services.Drawing;
using Xunit;

namespace OutlookShredder.Proxy.Tests;

// PartLabel adds the shop cutting-aid text (qty / material / thickness) on a no-cut layer, centered
// above the part, starting at 1" tall and shrinking only so the widest line fits the part's width.
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

    private static double LabelHeight(CutGeometry g) => g.Entities.Where(e => e.Type == "text").Max(e => e.Height);

    [Fact]
    public void Wide_part_keeps_the_full_one_inch_font()
    {
        var geo = Rect(100, 10);
        PartLabel.AddTo(geo, 2, "HRS", 0.625);
        Assert.Equal(1.0, LabelHeight(geo), 6);   // never grows past the 1" start
    }

    [Fact]
    public void Narrow_part_shrinks_the_font_to_fit_its_width()
    {
        var geo = Rect(3.0, 4.0);
        PartLabel.AddTo(geo, 2, "HRS", 0.625);

        double h = LabelHeight(geo);
        Assert.True(h < 1.0, $"narrow part should shrink the font, got {h}");

        // The widest line, at the chosen height, stays within the 3" part width.
        int widest = geo.Entities.Where(e => e.Type == "text").Max(e => (e.Text ?? "").Length);
        Assert.True(widest * 0.72 * h <= 3.0 + 1e-6, "label wider than the part");
    }

    [Fact]
    public void Label_is_centered_above_the_part_on_the_notes_layer()
    {
        var geo = Rect(20, 6);
        PartLabel.AddTo(geo, 4, "CRS", 0.075);

        var texts = geo.Entities.Where(e => e.Type == "text").ToList();
        Assert.NotEmpty(texts);
        Assert.All(texts, t => Assert.Equal(PartLabel.LayerName, t.Layer));
        Assert.All(texts, t => Assert.Equal(10.0, t.Cx, 6));   // centered on the 20-wide part
        Assert.All(texts, t => Assert.True(t.Cy >= 6.0, "label sits above the part top (y=6)"));
        Assert.Contains(geo.Layers, l => l.Name == PartLabel.LayerName && l.Color == PartLabel.LayerColor);
    }

    [Fact]
    public void Supplied_quantity_of_one_still_prints()
    {
        var geo = Rect(20, 6);
        PartLabel.AddTo(geo, 1, "HRS", 0.625);
        Assert.Contains(geo.Entities.Where(e => e.Type == "text"), t => t.Text == "x1");
    }

    [Fact]
    public void Nothing_to_say_adds_nothing()
    {
        var geo = Rect(10, 5);
        PartLabel.AddTo(geo, null, "", 0);   // no quantity, no material, no thickness
        Assert.DoesNotContain(geo.Entities, e => e.Type == "text");
    }
}
