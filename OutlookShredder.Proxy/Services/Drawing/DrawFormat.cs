using System.Globalization;

namespace OutlookShredder.Proxy.Services.Drawing;

/// <summary>
/// Shop-facing number formatting for the drawing output:
/// dimensions in fractional inches (to 1/16", decimal below that), and thickness shown as a sheet
/// gauge for steel / stainless / galvanized or a decimal for aluminium and other metals.
/// </summary>
public static class DrawFormat
{
    /// <summary>
    /// A dimension with the inch mark. Shown as a reduced fraction ONLY when the value is exactly a
    /// 1/16 multiple (e.g. <c>2"</c>, <c>3/4"</c>, <c>2-3/16"</c>, <c>4-1/4"</c>); any other value is
    /// shown as decimal inches to 3dp — never rounded to a fraction. <c>0 → 0"</c>.
    /// </summary>
    public static string FracInch(double v)
    {
        if (v < 0) return "-" + FracInch(-v);
        if (v == 0) return "0\"";

        double sx = v * 16.0;
        double nearest = Math.Round(sx);
        if (nearest >= 1 && Math.Abs(sx - nearest) < 1e-3)   // exact sixteenth only — no rounding
        {
            int sixteenths = (int)nearest;
            int whole = sixteenths / 16, rem = sixteenths % 16;
            if (rem == 0) return $"{whole}\"";
            int g = Gcd(rem, 16);
            int n = rem / g, d = 16 / g;
            return whole == 0 ? $"{n}/{d}\"" : $"{whole}-{n}/{d}\"";
        }
        return DecInch(v);
    }

    /// <summary>A dimension as decimal inches (3dp) with the inch mark, e.g. <c>9.063"</c>.</summary>
    public static string DecInch(double v) => v.ToString("0.###", CultureInfo.InvariantCulture) + "\"";

    private static int Gcd(int a, int b) { while (b != 0) (a, b) = (b, a % b); return a; }

    /// <summary>Material family for a parsed material token (e.g. "CRS" → ColdRolled).</summary>
    public static MaterialFamily FamilyForToken(string token) => token switch
    {
        "CRS"    => MaterialFamily.ColdRolled,
        "HRS"    => MaterialFamily.HotRolled,
        "HRPO"   => MaterialFamily.HotRolled,
        "galv"   => MaterialFamily.Galvanized,
        "SS"     => MaterialFamily.Stainless,
        "alum"   => MaterialFamily.Aluminium,
        "Brass"  => MaterialFamily.Brass,
        "Copper" => MaterialFamily.Copper,
        _        => MaterialFamily.Unknown,
    };

    /// <summary>True for the families the shop calls out by gauge (steel / stainless / galvanized).</summary>
    public static bool UsesGauge(string token) => FamilyForToken(token) switch
    {
        MaterialFamily.ColdRolled or MaterialFamily.HotRolled or
        MaterialFamily.Galvanized or MaterialFamily.Stainless => true,
        _ => false,
    };

    /// <summary>
    /// Thickness for display: a gauge (e.g. <c>11 ga</c>) for steel/SS/galv — snapped to the nearest
    /// standard gauge — falling back to a fraction for thick stock with no standard gauge (e.g. plate);
    /// a decimal inch value for aluminium and other metals.
    /// </summary>
    public static string ThicknessLabel(string token, double thickness)
    {
        if (UsesGauge(token))
        {
            var g = GaugeTables.NearestGauge(FamilyForToken(token), thickness);
            return g is not null ? $"{g} ga" : FracInch(thickness);
        }
        return thickness.ToString("0.###", CultureInfo.InvariantCulture) + "\"";
    }
}
