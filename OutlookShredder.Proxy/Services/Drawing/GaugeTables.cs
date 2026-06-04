namespace OutlookShredder.Proxy.Services.Drawing;

/// <summary>
/// Standard sheet-gauge → decimal thickness (inches) lookups, by material family.
/// Steel families use the Manufacturers' Standard Gauge; galvanized uses the galvanized
/// standard (zinc coat makes it slightly thicker); stainless uses the stainless gauge;
/// aluminium / brass / copper use the Brown &amp; Sharpe (AWG) gauge. An explicit decimal
/// thickness always overrides a gauge.
/// </summary>
public static class GaugeTables
{
    // Manufacturers' Standard Gauge for uncoated sheet steel (CRS / HRS).
    private static readonly IReadOnlyDictionary<int, double> Steel = new Dictionary<int, double>
    {
        [7] = 0.1793, [8] = 0.1644, [9] = 0.1495, [10] = 0.1345, [11] = 0.1196, [12] = 0.1046,
        [13] = 0.0897, [14] = 0.0747, [15] = 0.0673, [16] = 0.0598, [17] = 0.0538, [18] = 0.0478,
        [19] = 0.0418, [20] = 0.0359, [21] = 0.0329, [22] = 0.0299, [23] = 0.0269, [24] = 0.0239,
        [25] = 0.0209, [26] = 0.0179, [27] = 0.0164, [28] = 0.0149,
    };

    // Manufacturers' Standard Gauge for galvanized sheet steel (includes zinc coating).
    private static readonly IReadOnlyDictionary<int, double> Galv = new Dictionary<int, double>
    {
        [8] = 0.1681, [9] = 0.1532, [10] = 0.1382, [11] = 0.1233, [12] = 0.1084, [13] = 0.0934,
        [14] = 0.0785, [15] = 0.0710, [16] = 0.0635, [17] = 0.0575, [18] = 0.0516, [19] = 0.0456,
        [20] = 0.0396, [21] = 0.0366, [22] = 0.0336, [23] = 0.0306, [24] = 0.0276, [25] = 0.0247,
        [26] = 0.0217, [27] = 0.0202, [28] = 0.0187,
    };

    // Stainless steel gauge.
    private static readonly IReadOnlyDictionary<int, double> Stainless = new Dictionary<int, double>
    {
        [7] = 0.1875, [8] = 0.1719, [9] = 0.1563, [10] = 0.1406, [11] = 0.1250, [12] = 0.1094,
        [13] = 0.0938, [14] = 0.0781, [15] = 0.0703, [16] = 0.0625, [17] = 0.0563, [18] = 0.0500,
        [19] = 0.0438, [20] = 0.0375, [21] = 0.0344, [22] = 0.0313, [23] = 0.0281, [24] = 0.0250,
        [25] = 0.0219, [26] = 0.0188, [27] = 0.0172, [28] = 0.0156,
    };

    // Brown & Sharpe (AWG) gauge — aluminium, brass, copper sheet.
    private static readonly IReadOnlyDictionary<int, double> BrownSharpe = new Dictionary<int, double>
    {
        [6] = 0.1620, [7] = 0.1443, [8] = 0.1285, [9] = 0.1144, [10] = 0.1019, [11] = 0.0907,
        [12] = 0.0808, [13] = 0.0720, [14] = 0.0641, [15] = 0.0571, [16] = 0.0508, [17] = 0.0453,
        [18] = 0.0403, [19] = 0.0359, [20] = 0.0320, [21] = 0.0285, [22] = 0.0253, [23] = 0.0226,
        [24] = 0.0201, [25] = 0.0179, [26] = 0.0159,
    };

    /// <summary>
    /// Resolves a gauge to a thickness for the given material family. Returns null when the
    /// gauge is not tabulated for that family.
    /// </summary>
    public static double? Thickness(MaterialFamily family, int gauge)
    {
        var table = TableFor(family);
        return table.TryGetValue(gauge, out var v) ? v : null;
    }

    /// <summary>Ordered (gauge, thickness) list for a material family — drives the thickness picker.</summary>
    public static IReadOnlyList<(int Gauge, double Thickness)> GaugesFor(MaterialFamily family)
        => TableFor(family).OrderBy(kv => kv.Key).Select(kv => (kv.Key, kv.Value)).ToList();

    private static IReadOnlyDictionary<int, double> TableFor(MaterialFamily family) => family switch
    {
        MaterialFamily.Galvanized => Galv,
        MaterialFamily.Stainless  => Stainless,
        MaterialFamily.Aluminium  => BrownSharpe,
        MaterialFamily.Brass      => BrownSharpe,
        MaterialFamily.Copper     => BrownSharpe,
        _                         => Steel,   // CRS / HRS / unknown
    };
}

/// <summary>Material family parsed from the user's text (drives gauge table, K-factor, display).</summary>
public enum MaterialFamily
{
    Unknown,
    Galvanized,
    ColdRolled,   // CRS
    HotRolled,    // HRS / HRPO
    Aluminium,
    Stainless,
    Brass,
    Copper,
}
