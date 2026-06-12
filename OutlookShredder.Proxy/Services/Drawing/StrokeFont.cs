using System.Globalization;

namespace OutlookShredder.Proxy.Services.Drawing;

/// <summary>
/// A minimal single-stroke ("engraving") vector font: each glyph is a set of open polyline strokes on a
/// 4-wide x 6-tall cell (x 0..4, baseline y=0, cap y=6). Used by <see cref="PartLabel"/> to draw the
/// no-cut shop label as ACTUAL GEOMETRY rather than DXF TEXT — CAM importers (NcStudio's NcEditor) drop
/// TEXT/MTEXT on import, so a text label would be invisible there; stroked geometry on a non-cut colour
/// is visible AND ignorable. Covers A-Z, 0-9 and a few marks; lowercase is upper-cased by the caller.
/// </summary>
internal static class StrokeFont
{
    /// <summary>Cap height in cell units (a glyph spans y 0..CapH).</summary>
    public const double CapH    = 6.0;
    private const double GlyphW  = 4.0;   // nominal glyph width in cell units
    private const double Advance = 5.0;   // per-character advance (glyph + 1-unit gap)

    // Each glyph: strokes separated by ';', points within a stroke by space, "x,y". Grid x 0..4, y 0..6.
    private static readonly Dictionary<char, string> Raw = new()
    {
        ['0'] = "0,0 4,0 4,6 0,6 0,0; 0,0 4,6",
        ['1'] = "1,5 2,6 2,0; 1,0 3,0",
        ['2'] = "0,5 1,6 3,6 4,5 4,4 0,0 4,0",
        ['3'] = "0,5 1,6 3,6 4,5 3,3 4,1 3,0 1,0 0,1",
        ['4'] = "3,0 3,6 0,2 4,2",
        ['5'] = "4,6 0,6 0,3 3,3 4,2 4,1 3,0 0,0",
        ['6'] = "4,5 3,6 1,6 0,5 0,1 1,0 3,0 4,1 4,2 3,3 0,3",
        ['7'] = "0,6 4,6 1,0",
        ['8'] = "1,3 0,4 0,5 1,6 3,6 4,5 4,4 3,3 1,3 0,2 0,1 1,0 3,0 4,1 4,2 3,3",
        ['9'] = "0,1 1,0 3,0 4,1 4,5 3,6 1,6 0,5 0,4 1,3 4,3",

        ['A'] = "0,0 2,6 4,0; 1,2 3,2",
        ['B'] = "0,0 0,6 3,6 4,5 4,4 3,3 0,3; 3,3 4,2 4,1 3,0 0,0",
        ['C'] = "4,5 3,6 1,6 0,5 0,1 1,0 3,0 4,1",
        ['D'] = "0,0 0,6 2,6 4,4 4,2 2,0 0,0",
        ['E'] = "4,6 0,6 0,0 4,0; 0,3 3,3",
        ['F'] = "4,6 0,6 0,0; 0,3 3,3",
        ['G'] = "4,5 3,6 1,6 0,5 0,1 1,0 3,0 4,1 4,3 2,3",
        ['H'] = "0,0 0,6; 4,0 4,6; 0,3 4,3",
        ['I'] = "1,0 3,0; 2,0 2,6; 1,6 3,6",
        ['J'] = "3,6 3,1 2,0 1,0 0,1",
        ['K'] = "0,0 0,6; 4,6 0,3 4,0",
        ['L'] = "0,6 0,0 4,0",
        ['M'] = "0,0 0,6 2,3 4,6 4,0",
        ['N'] = "0,0 0,6 4,0 4,6",
        ['O'] = "1,0 0,1 0,5 1,6 3,6 4,5 4,1 3,0 1,0",
        ['P'] = "0,0 0,6 3,6 4,5 4,4 3,3 0,3",
        ['Q'] = "1,0 0,1 0,5 1,6 3,6 4,5 4,1 3,0 1,0; 2,2 4,0",
        ['R'] = "0,0 0,6 3,6 4,5 4,4 3,3 0,3; 2,3 4,0",
        ['S'] = "4,5 3,6 1,6 0,5 0,4 1,3 3,3 4,2 4,1 3,0 1,0 0,1",
        ['T'] = "0,6 4,6; 2,6 2,0",
        ['U'] = "0,6 0,1 1,0 3,0 4,1 4,6",
        ['V'] = "0,6 2,0 4,6",
        ['W'] = "0,6 1,0 2,3 3,0 4,6",
        ['X'] = "0,0 4,6; 0,6 4,0",
        ['Y'] = "0,6 2,3 4,6; 2,3 2,0",
        ['Z'] = "0,6 4,6 0,0 4,0",

        ['.'] = "1.5,0 2.5,0 2.5,0.8 1.5,0.8 1.5,0",
        ['"'] = "1,6 1,4.5; 3,6 3,4.5",
        ['/'] = "0,0 4,6",
        ['-'] = "0.5,3 3.5,3",
    };

    private static readonly Dictionary<char, List<List<(double X, double Y)>>> Glyphs =
        Raw.ToDictionary(kv => kv.Key, kv => Parse(kv.Value));

    private static List<List<(double, double)>> Parse(string s) =>
        s.Split(';', StringSplitOptions.RemoveEmptyEntries).Select(stroke =>
            stroke.Trim().Split(' ', StringSplitOptions.RemoveEmptyEntries).Select(pt =>
            {
                var xy = pt.Split(',');
                return (double.Parse(xy[0], CultureInfo.InvariantCulture),
                        double.Parse(xy[1], CultureInfo.InvariantCulture));
            }).ToList()
        ).ToList();

    /// <summary>Width of a rendered string in cell units (used for centering + font-fit).</summary>
    public static double WidthUnits(string s) => s.Length == 0 ? 0 : s.Length * Advance - (Advance - GlyphW);

    /// <summary>
    /// The string's strokes in cell units, laid left-to-right with the baseline at y=0 and the left edge
    /// at x=0. Unknown characters and spaces just advance the pen (no strokes).
    /// </summary>
    public static IEnumerable<List<(double X, double Y)>> Strokes(string s)
    {
        double x = 0;
        foreach (var ch in s)
        {
            if (Glyphs.TryGetValue(char.ToUpperInvariant(ch), out var glyph))
                foreach (var stroke in glyph)
                    yield return stroke.Select(p => (x + p.X, p.Y)).ToList();
            x += Advance;
        }
    }
}
