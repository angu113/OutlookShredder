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

    [Fact]
    public void Rejoined_note_parses_with_finish_outside()
    {
        var desc = PickingSlipFabAppender.ExtractFabDescs(Slip())[0];
        var spec = DrawingTextParser.Parse(desc);
        Assert.Equal(PartType.UChannel, spec.Type);
        Assert.Equal(FinishSide.Outside, spec.Finish);
    }
}
