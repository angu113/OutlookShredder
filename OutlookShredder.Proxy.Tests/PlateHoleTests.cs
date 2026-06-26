using System;
using System.Linq;
using OutlookShredder.Proxy.Models.Drawing;
using OutlookShredder.Proxy.Services.Drawing;
using Xunit;

namespace OutlookShredder.Proxy.Tests;

// Flitch-plate hole edge distances + single-row, and base-plate radiused corners.
// The flitch edge distances MUST come from independent "topedge"/"botedge" keywords — NOT a "/"-combined
// "edge X/Y", which the number grammar parses as a fraction (e.g. "2/2" -> 1.0): the "always 1 inch,
// input ignored" bug. Changing any of this behaviour must update this test in the same commit.
public class PlateHoleTests
{
    private const string FlitchHead = "Flitch 48 x 6, 16ga steel, holes 0.75 staggered @ 16, lhs 2 rhs 2";

    private static System.Collections.Generic.List<double> HoleRowYs(FlatPatternResult fp) =>
        fp.Holes.Select(h => Math.Round(h.Item2, 3)).Distinct().OrderBy(y => y).ToList();

    private static CutEntity Outline(FlatPatternResult fp) =>
        fp.Cut.Entities.First(e => e.Type == "polyline" && e.Layer != PartLabel.LayerName);

    [Fact]
    public void Flitch_topedge_botedge_are_honored_independently()
    {
        var h = DrawingTextParser.Parse($"{FlitchHead}, topedge 2 botedge 3").Holes!;
        Assert.Equal(2, h.TopEdge, 3);
        Assert.Equal(3, h.BottomEdge, 3);
    }

    [Fact]
    public void Flitch_integer_edges_do_not_collapse_to_one_inch()   // regression: "2/2" used to parse as 1.0
    {
        var h = DrawingTextParser.Parse($"{FlitchHead}, topedge 2 botedge 2").Holes!;
        Assert.Equal(2, h.TopEdge, 3);
        Assert.Equal(2, h.BottomEdge, 3);
    }

    [Fact]
    public void Flitch_single_edge_keyword_sets_both_rows()
    {
        var h = DrawingTextParser.Parse($"{FlitchHead}, edge 1.5").Holes!;
        Assert.Equal(1.5, h.TopEdge, 3);
        Assert.Equal(1.5, h.BottomEdge, 3);
    }

    [Fact]
    public void Flitch_hole_rows_sit_at_the_specified_edge_distances()
    {
        var fp = FlatPattern.Develop(DrawingTextParser.Parse($"{FlitchHead}, topedge 2 botedge 1"));
        var ys = HoleRowYs(fp);
        Assert.Contains(1.0, ys);   // bottom row 1" from the bottom edge
        Assert.Contains(4.0, ys);   // top row 2" from the top edge (W=6 -> 6-2)
    }

    [Fact]
    public void Flitch_single_row_drills_only_the_top_row()
    {
        var spec = DrawingTextParser.Parse($"{FlitchHead}, topedge 2 botedge 1 single-row");
        Assert.True(spec.Holes!.SingleRow);
        var ys = HoleRowYs(FlatPattern.Develop(spec));
        Assert.Single(ys);            // one line of holes
        Assert.Equal(4.0, ys[0], 3);  // at the top-edge distance
    }

    [Fact]
    public void BasePlate_radius_corners_parse_and_round_the_outline()
    {
        var spec = DrawingTextParser.Parse("Base Plate 8 x 8, 0.5 steel, 4 holes 0.75 edge 1 radius-corners 0.5");
        Assert.Equal(0.5, spec.CornerRadius, 3);
        // A square plate outline has 4 vertices; a radiused one tessellates each corner into an arc.
        Assert.True(Outline(FlatPattern.Develop(spec)).Vertices!.Count > 4);
    }

    [Fact]
    public void BasePlate_without_radius_keeps_square_corners()
    {
        var spec = DrawingTextParser.Parse("Base Plate 8 x 8, 0.5 steel, 4 holes 0.75 edge 1");
        Assert.Equal(0, spec.CornerRadius, 3);
        Assert.Equal(4, Outline(FlatPattern.Develop(spec)).Vertices!.Count);
    }
}
