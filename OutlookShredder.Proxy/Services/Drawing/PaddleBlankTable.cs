namespace OutlookShredder.Proxy.Services.Drawing;

/// <summary>One ASME B16.48 spade (paddle blind / "frying pan") size, in inches.</summary>
public sealed record PaddleBlank(
    string Nps, double NpsValue, int PressureClass,
    double Od, double CenterToEnd, double HandleWidth, double Thickness);

/// <summary>
/// Hardcoded ASME B16.48 spade (paddle blind) dimensions, NPS 1/2"–20", Class 150 &amp; 300.
/// It is a fixed engineering standard, so the table lives in code (cached at process scope) rather
/// than SharePoint — add a class (e.g. #600) by extending this list.
///
/// Sources: spade outside diameter + thickness from the US-customary paddle-blind chart
/// (hydroblindsolutions.com); center-to-end and handle width converted to inches from the ASME
/// B16.48 metric table (wermac.org). The five sizes the US chart omits (3-1/2, 5, 14, 16, 18) take
/// the metric OD rounded to the nearest 1/8". Center-to-end is the disc centre to the handle tip;
/// thickness is the standard minimum (the wizard lets the user pick their actual plate).
/// </summary>
public static class PaddleBlankTable
{
    public static readonly IReadOnlyList<PaddleBlank> All = new[]
    {
        // ── Class 150 ────────────────────────────────────────────────────────
        //              NPS       value  cls   OD       C(c-to-end)  W(handle)  t
        new PaddleBlank("1/2",    0.5,   150,  1.75,    4.96,        1.25,      0.25),
        new PaddleBlank("3/4",    0.75,  150,  2.125,   5.16,        1.25,      0.25),
        new PaddleBlank("1",      1.0,   150,  2.5,     5.35,        1.25,      0.25),
        new PaddleBlank("1-1/4",  1.25,  150,  2.875,   5.71,        1.25,      0.25),
        new PaddleBlank("1-1/2",  1.5,   150,  3.25,    5.71,        1.25,      0.25),
        new PaddleBlank("2",      2.0,   150,  4.0,     6.10,        1.25,      0.25),
        new PaddleBlank("2-1/2",  2.5,   150,  4.75,    6.69,        1.25,      0.25),
        new PaddleBlank("3",      3.0,   150,  5.25,    6.69,        1.25,      0.25),
        new PaddleBlank("3-1/2",  3.5,   150,  6.25,    7.95,        1.5,       0.375),
        new PaddleBlank("4",      4.0,   150,  6.75,    7.95,        1.5,       0.375),
        new PaddleBlank("5",      5.0,   150,  7.625,   8.86,        1.5,       0.375),
        new PaddleBlank("6",      6.0,   150,  8.625,   8.86,        1.5,       0.5),
        new PaddleBlank("8",      8.0,   150,  10.875,  10.51,       1.5,       0.5),
        new PaddleBlank("10",     10.0,  150,  13.25,   12.68,       1.75,      0.625),
        new PaddleBlank("12",     12.0,  150,  16.0,    14.06,       1.75,      0.75),
        new PaddleBlank("14",     14.0,  150,  17.625,  14.88,       1.75,      0.75),
        new PaddleBlank("16",     16.0,  150,  20.125,  16.14,       1.75,      0.875),
        new PaddleBlank("18",     18.0,  150,  21.5,    16.81,       2.0,       1.0),
        new PaddleBlank("20",     20.0,  150,  23.75,   17.91,       2.0,       1.125),

        // ── Class 300 ────────────────────────────────────────────────────────
        new PaddleBlank("1/2",    0.5,   300,  2.0,     5.08,        1.25,      0.25),
        new PaddleBlank("3/4",    0.75,  300,  2.5,     5.35,        1.25,      0.25),
        new PaddleBlank("1",      1.0,   300,  2.75,    5.47,        1.25,      0.25),
        new PaddleBlank("1-1/4",  1.25,  300,  3.125,   5.91,        1.25,      0.25),
        new PaddleBlank("1-1/2",  1.5,   300,  3.625,   5.91,        1.25,      0.25),
        new PaddleBlank("2",      2.0,   300,  4.25,    6.22,        1.25,      0.375),
        new PaddleBlank("2-1/2",  2.5,   300,  5.0,     6.97,        1.25,      0.375),
        new PaddleBlank("3",      3.0,   300,  5.75,    6.97,        1.25,      0.375),
        new PaddleBlank("3-1/2",  3.5,   300,  6.375,   8.07,        1.5,       0.5),
        new PaddleBlank("4",      4.0,   300,  7.0,     8.07,        1.5,       0.5),
        new PaddleBlank("5",      5.0,   300,  8.375,   9.45,        1.5,       0.625),
        new PaddleBlank("6",      6.0,   300,  9.75,    9.45,        1.5,       0.625),
        new PaddleBlank("8",      8.0,   300,  12.0,    11.06,       1.5,       0.875),
        new PaddleBlank("10",     10.0,  300,  14.125,  13.11,       1.75,      1.0),
        new PaddleBlank("12",     12.0,  300,  16.5,    14.29,       1.75,      1.125),
        new PaddleBlank("14",     14.0,  300,  19.0,    15.55,       1.75,      1.25),
        new PaddleBlank("16",     16.0,  300,  21.125,  16.61,       1.75,      1.5),
        new PaddleBlank("18",     18.0,  300,  23.375,  17.72,       2.0,       1.625),
        new PaddleBlank("20",     20.0,  300,  25.625,  18.90,       2.0,       1.75),
    };

    /// <summary>Looks up a spade by NPS value (within tolerance) and pressure class.</summary>
    public static PaddleBlank? Find(double npsValue, int pressureClass) =>
        All.FirstOrDefault(p => p.PressureClass == pressureClass && Math.Abs(p.NpsValue - npsValue) < 0.01);
}
