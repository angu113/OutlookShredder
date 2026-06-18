using System.Linq;
using OutlookShredder.Proxy.Models.Drawing;
using OutlookShredder.Proxy.Services.Drawing;
using Xunit;

namespace OutlookShredder.Proxy.Tests;

// Polish / grain direction: a per-part 2-value axis (Vertical / Horizontal / unset) carried in the
// canonical fab-note text ("polish vertical|horizontal") so it round-trips with NO stored state — the
// same contract as the finish side. When set, the drawing gets a double-headed arrow along the axis +
// the bilingual "Dirección de pulido" label (ASCII "DIRECCION DE PULIDO" in the DXF). Changing any of
// this behaviour must update this test in the same commit.
public class PolishDirectionTests
{
    // ── Parsing / round-trip ──────────────────────────────────────────────────

    [Fact]
    public void Unset_by_default()
    {
        Assert.Equal(PolishDirection.None,
            DrawingTextParser.Parse("U 4 x 2 x 36, 16ga CRS").PolishDirection);
    }

    [Theory]
    [InlineData("vertical", PolishDirection.Vertical)]
    [InlineData("horizontal", PolishDirection.Horizontal)]
    public void Parses_the_polish_token(string word, PolishDirection expected)
    {
        Assert.Equal(expected,
            DrawingTextParser.Parse($"U 4 x 2 x 36, 16ga CRS, polish {word}").PolishDirection);
    }

    // One per sub-parser, proving the token is threaded into EVERY part type — including Column, which
    // returns before the finish parse (so polish must be parsed earlier than finish).
    [Theory]
    [InlineData("U 4 x 2 x 36, 16ga CRS, polish vertical")]
    [InlineData("L 2 x 3 x 36, 16ga CRS, polish vertical")]
    [InlineData("Flitch 48 x 6, 0.25 steel, polish vertical")]
    [InlineData("Pan 24 x 18 x 2, 2 long 2 short, 16ga, polish vertical")]
    [InlineData("Column height 96, base 10 x 10 x 0.75, bearing 8 x 8 x 0.5, column square 4 x 4 wall 0.25, HRS, polish vertical")]
    public void Polish_is_available_on_every_part_type(string text)
    {
        Assert.Equal(PolishDirection.Vertical, DrawingTextParser.Parse(text).PolishDirection);
    }

    [Fact]
    public void Polish_coexists_with_finish_and_does_not_disturb_it()
    {
        var spec = DrawingTextParser.Parse("U 4 x 2 x 36, 16ga CRS, finish outside, polish horizontal");
        Assert.Equal(FinishSide.Outside, spec.Finish);
        Assert.Equal(PolishDirection.Horizontal, spec.PolishDirection);
        Assert.Equal(PartType.UChannel, spec.Type);   // stripped tokens don't corrupt the core dims
    }

    // ── DXF geometry (PolishLabel) — same layer/style approach as the existing DXF heading ─────────

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

    private static List<CutEntity> L1(CutGeometry g) =>
        g.Entities.Where(e => e.Layer == PartLabel.LayerName).ToList();

    [Fact]
    public void Unset_adds_no_geometry()
    {
        var geo = Rect(20, 6);
        PolishLabel.AddTo(geo, PolishDirection.None);
        Assert.Empty(L1(geo));
    }

    [Fact]
    public void Arrow_and_label_land_on_the_L1_process_layer()
    {
        var geo = Rect(20, 6);
        PolishLabel.AddTo(geo, PolishDirection.Vertical);
        Assert.NotEmpty(L1(geo));
        Assert.Contains(geo.Layers, l => l.Name == PartLabel.LayerName && l.Color == PartLabel.LayerColor);
        Assert.Contains(L1(geo), e => e.Type == "polyline");   // stroked label glyphs (not DXF TEXT)
        Assert.Contains(L1(geo), e => e.Type == "line");       // arrow segments
    }

    [Fact]
    public void Label_sits_to_the_right_of_the_part()
    {
        var geo = Rect(20, 6);
        PolishLabel.AddTo(geo, PolishDirection.Horizontal);
        var labelPts = L1(geo).Where(e => e.Type == "polyline").SelectMany(e => e.Vertices!).ToList();
        Assert.NotEmpty(labelPts);
        Assert.True(labelPts.Min(p => p.X) >= 20.0 - 1e-6, "polish label is outside the part, on the right");
    }

    [Fact]
    public void Vertical_and_horizontal_shafts_run_along_their_axis()
    {
        var v = Rect(20, 6); PolishLabel.AddTo(v, PolishDirection.Vertical);
        var h = Rect(20, 6); PolishLabel.AddTo(h, PolishDirection.Horizontal);
        Assert.Contains(L1(v).Where(e => e.Type == "line"),
            e => System.Math.Abs(e.X1 - e.X2) < 1e-9 && System.Math.Abs(e.Y1 - e.Y2) > 1);   // vertical shaft
        Assert.Contains(L1(h).Where(e => e.Type == "line"),
            e => System.Math.Abs(e.Y1 - e.Y2) < 1e-9 && System.Math.Abs(e.X1 - e.X2) > 1);   // horizontal shaft
    }

    // ── End-to-end via the engine ──────────────────────────────────────────────

    [Fact]
    public void Develop_bakes_the_polish_arrow_into_the_cut_geometry()
    {
        int Arrows(FlatPatternResult fp) =>
            fp.Cut.Entities.Count(e => e.Layer == PartLabel.LayerName && e.Type == "line");

        var with    = FlatPattern.Develop(DrawingTextParser.Parse("Flitch 48 x 6, 0.25 steel, polish vertical"));
        var without = FlatPattern.Develop(DrawingTextParser.Parse("Flitch 48 x 6, 0.25 steel"));
        Assert.True(Arrows(with) >= 5, "vertical polish adds the 5 arrow segments");
        Assert.Equal(0, Arrows(without));
    }

    // ── Label language by context (Angus, CP3) ─────────────────────────────────
    // Pixar PDFs render the polish label bilingually; the auto picking-slip drawings render it
    // Spanish-only (the shop floor reads those). The renderer picks Bi.T vs Bi.Es accordingly.

    [Fact]
    public void Pixar_label_is_bilingual()
    {
        var s = Bi.T("polish.direction");
        Assert.StartsWith("Polish direction", s);   // English half present
        Assert.Contains("pulido", s);               // Spanish half present
    }

    [Fact]
    public void Picking_slip_label_is_spanish_only()
    {
        var s = Bi.Es("polish.direction");
        Assert.DoesNotContain("Polish direction", s);   // no English
        Assert.Contains("pulido", s);                   // Spanish term only
    }

    // ── PDF render smoke (Angus: DXF L1 labels must not leak onto the PDF; rotated label must not throw) ──
    // Exercises the cut-geometry render paths that skip the L1 layer (paddle outline, pan bounds+draw)
    // and the rotated/leadered polish callout, across vertical + horizontal axes. A real PDF is produced.
    [Theory]
    [InlineData("Frying Pan 6 #150, 0.25 steel, polish vertical")]
    [InlineData("Pan 24 x 18 x 2, 2 long 2 short, 16ga, polish horizontal")]
    [InlineData("U 4 x 2 x 36, 16ga CRS, polish vertical")]
    [InlineData("Flitch 48 x 6, 0.25 steel, polish horizontal")]
    public void Pdf_render_with_polish_produces_a_pdf(string text)
    {
        var fp = FlatPattern.Develop(DrawingTextParser.Parse(text));
        var pdf = DrawingPdfRenderer.Render(fp);
        Assert.NotNull(pdf);
        Assert.True(pdf.Length > 1000, $"expected a real PDF, got {pdf.Length} bytes");
    }
}
