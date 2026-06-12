using System.Linq;
using OutlookShredder.Proxy.Services;
using OutlookShredder.Proxy.Services.Drawing;
using Xunit;

namespace OutlookShredder.Proxy.Tests;

// FabDxfBuilder develops a slip's deduped FAB notes into ONE DXF, laying the parts out left-to-right
// 1" apart and bottom-aligned to Y=0. HSK-SO1036124 carries two distinct flitches.
public class FabDxfBuilderTests
{
    private const string FlitchA = "Flitch 109 x 7.25, 0.625 HRS, holes 0.75 paired @ 16, lhs 2 rhs 2, edge 1.5/1.5";
    private const string FlitchB = "Flitch 79 x 7.25, 0.625 HRS, holes 0.75 paired @ 16, lhs 2 rhs 2, edge 1.5/1.5";

    // Default quantity 1 unless a note carries its own.
    private static FabNote[] Notes(params string[] descs) => descs.Select(d => new FabNote(1, d)).ToArray();

    private static (double MinX, double MinY, double MaxX, double MaxY) Bounds(CutGeometry geo)
    {
        double minX = double.MaxValue, minY = double.MaxValue, maxX = double.MinValue, maxY = double.MinValue;
        void Acc(double x, double y)
        {
            if (x < minX) minX = x; if (y < minY) minY = y;
            if (x > maxX) maxX = x; if (y > maxY) maxY = y;
        }
        foreach (var e in geo.Entities)
        {
            if (e.Layer == PartLabel.LayerName) continue;   // the no-cut label must not affect part bounds
            if (e.Type == "polyline") foreach (var v in e.Vertices!) Acc(v.X, v.Y);
            else if (e.Type == "line") { Acc(e.X1, e.Y1); Acc(e.X2, e.Y2); }
            else if (e.Type == "circle") { Acc(e.Cx - e.R, e.Cy - e.R); Acc(e.Cx + e.R, e.Cy + e.R); }
        }
        return (minX, minY, maxX, maxY);
    }

    [Fact]
    public void Single_part_is_anchored_at_origin()
    {
        var (geo, parts) = FabDxfBuilder.Combine(Notes(FlitchA));
        Assert.NotNull(geo);
        Assert.Single(parts);
        var b = Bounds(geo!);
        Assert.Equal(0, b.MinX, 3);   // left edge at X=0
        Assert.Equal(0, b.MinY, 3);   // bottom edge at Y=0
    }

    [Fact]
    public void Two_parts_combine_into_one_geometry_offset_to_the_right()
    {
        var oneWidth = Bounds(FabDxfBuilder.Combine(Notes(FlitchA)).Geo!).MaxX;

        var (geo, parts) = FabDxfBuilder.Combine(Notes(FlitchA, FlitchB));
        Assert.NotNull(geo);
        Assert.Equal(2, parts.Count);

        // Layers are merged by name (cut / bend / notes), not duplicated per part.
        Assert.Equal(geo!.Layers.Select(l => l.Name).Distinct().Count(), geo.Layers.Count);

        // The second part sits to the right of the first with a >= 1" gap: total width exceeds a single
        // part by at least the gap, proving the parts don't overlap.
        var total = Bounds(geo).MaxX;
        Assert.True(total >= oneWidth + 1.0, $"expected combined width >= {oneWidth + 1.0}, got {total}");
    }

    [Fact]
    public void Both_distinct_flitch_slugs_are_reported()
    {
        var (_, parts) = FabDxfBuilder.Combine(Notes(FlitchA, FlitchB));
        Assert.Contains("flitch_109x7.25", parts);
        Assert.Contains("flitch_79x7.25", parts);
    }

    [Fact]
    public void Each_part_gets_a_no_cut_stroked_label()
    {
        var (geo, _) = FabDxfBuilder.Combine(new[] { new FabNote(2, FlitchA) });
        Assert.NotNull(geo);

        // The label is GEOMETRY (single-stroke polylines) on the no-cut white Notes layer — not DXF
        // TEXT (NcStudio drops text on import).
        Assert.Contains(geo!.Layers, l => l.Name == PartLabel.LayerName && l.Color == PartLabel.LayerColor);
        Assert.Contains(geo.Entities, e => e.Type == "polyline" && e.Layer == PartLabel.LayerName);
        Assert.DoesNotContain(geo.Entities, e => e.Type == "text");
    }

    [Fact]
    public void Undevelopable_notes_are_skipped_not_fatal()
    {
        var (geo, parts) = FabDxfBuilder.Combine(Notes("this is not a part", FlitchA));
        Assert.NotNull(geo);
        Assert.Single(parts);
        Assert.Equal("flitch_109x7.25", parts[0]);
    }

    [Fact]
    public void Nothing_developable_yields_null_geometry()
    {
        var (geo, _) = FabDxfBuilder.Combine(Notes("not a part", "also not a part"));
        Assert.Null(geo);
        Assert.Null(FabDxfBuilder.Build(Notes("not a part")));
    }

    [Fact]
    public void Build_output_loads_back_as_a_valid_dxf_with_both_parts_and_labels()
    {
        var result = FabDxfBuilder.Build(new[] { new FabNote(2, FlitchA), new FabNote(3, FlitchB) })!;
        Assert.Equal(2, result.Parts.Count);

        using var ms = new System.IO.MemoryStream(result.Dxf);
        var doc = netDxf.DxfDocument.Load(ms);
        Assert.NotNull(doc);

        // Two flitch outlines survive the round-trip on the cut layer.
        Assert.True(doc.Entities.Polylines2D.Count() >= 2,
            $"expected >= 2 outline polylines, got {doc.Entities.Polylines2D.Count()}");
        Assert.Contains(doc.Layers, l => l.Name == FlatPattern.CutLayer);
        // The stroked labels survive too, as polylines on the no-cut Notes layer (no DXF text).
        Assert.Contains(doc.Layers, l => l.Name == PartLabel.LayerName);
        Assert.True(doc.Entities.Polylines2D.Count(p => p.Layer.Name == PartLabel.LayerName) >= 2,
            "expected stroked label polylines on the Notes layer");
        Assert.Empty(doc.Entities.Texts);
    }
}
