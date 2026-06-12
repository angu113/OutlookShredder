using System.Globalization;

namespace OutlookShredder.Proxy.Services.Drawing;

/// <summary>
/// Shop cutting-aid label baked into a part's cut geometry: quantity (when &gt; 1), material, and
/// thickness, as no-cut TEXT centered above the part. Added by <see cref="FlatPattern.Develop"/> so
/// EVERY part the engine develops — and therefore every DXF it emits, single or combined — carries the
/// label. The text lives on a dedicated <see cref="LayerName"/> layer in a colour outside NcStudio's
/// cut/mark/small map, so the CNC ignores it. The font starts <see cref="StartHeight"/>" tall and only
/// shrinks so the widest line fits within the part's width.
/// </summary>
public static class PartLabel
{
    /// <summary>No-cut annotation layer.</summary>
    public const string LayerName = "Notes";
    /// <summary>ACI 7 (white) — outside the yellow/blue/cyan cut/mark/small map, so NcStudio ignores it.</summary>
    public const short LayerColor = 7;

    private const double StartHeight = 1.0;   // inches — the starting (max) font height
    private const double MinHeight   = 0.1;   // floor so a tiny part still gets a (small) label
    private const double CharWidth   = 0.72;  // glyph width as a fraction of height (netDxf default txt style)
    private const double GapFactor   = 0.6;   // clear space above the part, as a fraction of the height
    private const double LineFactor  = 1.4;   // line-to-line spacing, as a fraction of the height

    /// <summary>
    /// Adds the label to <paramref name="geo"/> in place. <paramref name="quantity"/> is null when the
    /// caller has no order quantity (e.g. the design wizard) — then no "xN" line is shown; a supplied
    /// quantity (e.g. a picking-slip FAB note) always prints, including "x1". No-op when there is
    /// nothing to say.
    /// </summary>
    public static void AddTo(CutGeometry geo, int? quantity, string? material, double thickness)
    {
        var lines = BuildLines(quantity, material, thickness);
        if (lines.Count == 0) return;

        var (minX, minY, maxX, maxY) = Bounds(geo);
        double partW = maxX - minX;
        if (partW <= 0) return;
        double centerX = (minX + maxX) / 2.0;

        // Start at 1" tall; shrink only so the widest line fits within the part's width.
        int widest = lines.Max(l => l.Length);
        double height = StartHeight;
        double widthAtStart = widest * CharWidth * height;
        if (widthAtStart > partW) height = Math.Max(MinHeight, partW / (widest * CharWidth));

        if (!geo.Layers.Any(l => l.Name.Equals(LayerName, StringComparison.OrdinalIgnoreCase)))
            geo.Layers.Add(new CutLayer { Name = LayerName, Color = LayerColor });

        // Stack above the part: the LAST line sits nearest the part, earlier lines above it.
        double y = maxY + GapFactor * height;
        for (int i = lines.Count - 1; i >= 0; i--)
        {
            geo.Entities.Add(CutEntity.Label(LayerName, lines[i], centerX, y, height));
            y += LineFactor * height;
        }
    }

    private static List<string> BuildLines(int? quantity, string? material, double thickness)
    {
        var lines = new List<string>();
        if (quantity is { } q && q >= 1) lines.Add($"x{q}");
        var mat = (material ?? "").Trim();
        var thk = thickness > 0 ? thickness.ToString("0.####", CultureInfo.InvariantCulture) + "\"" : "";
        var matThk = string.Join(" ", new[] { mat, thk }.Where(s => s.Length > 0));
        if (matThk.Length > 0) lines.Add(matThk);
        return lines;
    }

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
            switch (e.Type)
            {
                case "polyline": if (e.Vertices != null) foreach (var v in e.Vertices) Acc(v.X, v.Y); break;
                case "line":     Acc(e.X1, e.Y1); Acc(e.X2, e.Y2); break;
                case "circle":   Acc(e.Cx - e.R, e.Cy - e.R); Acc(e.Cx + e.R, e.Cy + e.R); break;
                // "text" excluded — a label must not inflate the part's measured bounds.
            }
        }
        if (minX > maxX) return (0, 0, 0, 0);
        return (minX, minY, maxX, maxY);
    }
}
