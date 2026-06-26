using OutlookShredder.Proxy.Models.Drawing;
using OutlookShredder.Proxy.Services.Drawing;
using Xunit;

namespace OutlookShredder.Proxy.Tests;

// Numeric entry must accept a no-leading-zero decimal (".5" == "0.5") in every dimension field.
public class DecimalEntryTests
{
    [Fact]
    public void Flitch_hole_dia_no_leading_zero()
        => Assert.Equal(0.5, DrawingTextParser.Parse("Flitch 48 x 6, 16ga steel, holes .5 staggered @ 16").Holes!.Diameter, 3);

    [Fact]
    public void Flitch_topedge_botedge_no_leading_zero()
    {
        var h = DrawingTextParser.Parse("Flitch 48 x 6, 16ga steel, holes 0.75 staggered @ 16, topedge .5 botedge .75").Holes!;
        Assert.Equal(0.5, h.TopEdge, 3);
        Assert.Equal(0.75, h.BottomEdge, 3);
    }

    [Fact]
    public void BasePlate_radius_no_leading_zero()
        => Assert.Equal(0.5, DrawingTextParser.Parse("Base Plate 8 x 8, 0.5 steel, 4 holes 0.75 edge 1 radius-corners .5").CornerRadius, 3);

    [Fact]
    public void UChannel_dims_no_leading_zero()
    {
        var s = DrawingTextParser.Parse("U web .5 flange .75 length 36, 16ga steel");
        Assert.Equal(0.5, s.Web.Value, 3);
        Assert.Equal(0.75, s.FlangeLeft.Value, 3);
    }
}
