using OutlookShredder.Proxy.Services;
using OutlookShredder.Proxy.Services.Drawing;
using OutlookShredder.Proxy.Models.Drawing;
using Xunit;

namespace OutlookShredder.Proxy.Tests;

// A FAB note can wrap in the narrow Product column of a picking slip, putting the tail (e.g. the
// finish side) on the next row. The appender must stitch those continuation rows back on so the
// drawing parser sees the whole note. Mirrors the HSK-SO1036091 case where "finish" / "outside" split.
public class FabNoteStitchTests
{
    // Product column left ≈ 150; far-left MSPC column ≈ 5.
    private static List<(string, double)> Slip() => new()
    {
        ("MSPC Product Qty", 5),
        ("AP3003H22TP/063 Aluminum 3003H22 Tread Plate", 5),
        ("FAB: (1) U flange 1 id flange 1 id web 1.625 id length 56.625, 14ga alum, finish", 150),
        ("outside", 150),                       // wrapped tail — same Product column
        ("SETUP CHARGE Setup Charge", 5),       // new line-item — reaches the MSPC column
    };

    [Fact]
    public void Stitches_wrapped_finish_tail_back_onto_the_note()
    {
        var descs = PickingSlipFabAppender.ExtractFabDescs(Slip());
        Assert.Single(descs);
        Assert.EndsWith("finish outside", descs[0]);
        Assert.DoesNotContain("SETUP", descs[0]);   // must not swallow the next line-item
    }

    // The FAB-note item letter (drawn on each appended drawing page AND shown in the ERP viewer's note
    // list) must follow deduped order identically on both sides: A…Z, then AA, AB, … This is the
    // client/proxy consistency contract — the letter on the page must match the letter in the list.
    [Theory]
    [InlineData(0,  "A")]
    [InlineData(25, "Z")]
    [InlineData(26, "AA")]
    [InlineData(27, "AB")]
    [InlineData(51, "AZ")]
    [InlineData(52, "BA")]
    public void FabLetter_follows_deduped_order(int index, string expected)
        => Assert.Equal(expected, PickingSlipFabAppender.FabLetter(index));

    [Fact]
    public void Rejoined_note_parses_with_finish_outside()
    {
        var desc = PickingSlipFabAppender.ExtractFabDescs(Slip())[0];
        var spec = DrawingTextParser.Parse(desc);
        Assert.Equal(PartType.UChannel, spec.Type);
        Assert.Equal(FinishSide.Outside, spec.Finish);
    }

    // HSK-SO1036124: OpenBravo prints each FAB note twice — inline in the line-items and again in the
    // special-instructions footer. The "FAB: (1) [ … ]" terminator bounds the capture, but the footer
    // box renders narrower and CLIPS the text ("edge" -> "edg"), so the two copies aren't textually
    // identical and the clipped one fails to parse the edge value (renders the default margin). Both
    // copies still develop to the same part slug (flitch_109x7.25), so they dedupe by slug — and keeping
    // the longest (least-clipped) capture keeps the copy with the correct, fully-specified geometry.
    private static List<(string, double)> ClippedDoublePrint() => new()
    {
        ("MSPC Product Qty on Hand", 5),
        ("HF/6257 Hot Rolled Flat Bar 0.625 X 7.000", 5),
        ("B: 1166.75 IN", 5),
        // inline copy (Product column ~190) — full text, ']' on the next row
        ("FAB: (1) [Flitch 109 x 7.25, 0.625 HRS, holes 0.75 paired @ 16, lhs 2 rhs 2, edge", 190),
        ("1.5/1.5]", 190),
        ("Laser # HSK-SO1036124 flitch_109x7.25", 190),
        ("Hole Punching HOLE PUNCHING", 5),
        // footer copy (special-instructions box ~30) — clipped: "edge" -> "edg"
        ("WRF Contracting LLC", 30),
        ("Pickup", 30),
        ("FAB: (1) [Flitch 109 x 7.25, 0.625 HRS, holes 0.75 paired @ 16, lhs 2 rhs 2, edg", 30),
        ("1.5/1.5]", 30),
        ("Laser # HSK-SO1036124 flitch_109x7.25", 30),
    };

    [Fact]
    public void Bracketed_note_is_bounded_and_excludes_trailing_lines()
    {
        var descs = PickingSlipFabAppender.ExtractFabDescs(ClippedDoublePrint());
        Assert.Equal(2, descs.Count);                       // the note appears twice on the slip
        Assert.All(descs, d => Assert.DoesNotContain("Laser", d));   // text after ']' excluded
        Assert.Contains(descs, d => d.EndsWith("edge 1.5/1.5"));     // inline copy captured in full
    }

    [Fact]
    public void Clipped_footer_echo_dedupes_by_slug_to_the_complete_copy()
    {
        var descs    = PickingSlipFabAppender.ExtractFabDescs(ClippedDoublePrint());
        var distinct = PickingSlipFabAppender.DedupeBySlug(PickingSlipFabAppender.DedupeDescs(descs));
        Assert.Single(distinct);                            // 2 drawing pages -> 1
        Assert.Contains("edge 1.5/1.5", distinct[0]);       // kept the full copy, not the clipped "edg"
        Assert.Equal(PartType.FlitchPlate, DrawingTextParser.Parse(distinct[0]).Type);
    }

    [Fact]
    public void Different_sized_flitches_are_not_collapsed()
    {
        var two = PickingSlipFabAppender.DedupeBySlug(new()
        {
            "Flitch 109 x 7.25, 0.625 HRS, holes 0.75 paired @ 16, lhs 2 rhs 2, edge 1.5/1.5",
            "Flitch 79 x 7.25, 0.625 HRS, holes 0.75 paired @ 16, lhs 2 rhs 2, edge 1.5/1.5",
        });
        Assert.Equal(2, two.Count);                         // distinct slugs -> kept separate
    }
}
