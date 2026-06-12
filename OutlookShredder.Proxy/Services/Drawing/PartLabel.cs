using System.Globalization;

namespace OutlookShredder.Proxy.Services.Drawing;

/// <summary>
/// Shop cutting-aid label baked into a part's cut geometry: quantity (when supplied), material, and
/// thickness, centered above the part. Added by <see cref="FlatPattern.Develop"/> so EVERY part the
/// engine develops — and every DXF it emits, single or combined — carries it.
///
/// The label is drawn as SINGLE-STROKE GEOMETRY (<see cref="StrokeFont"/> polylines), NOT DXF TEXT:
/// NcStudio's NcEditor (and most CAM DXF importers) silently drop TEXT/MTEXT on import, so a text label
/// would be invisible to the operator. Geometry on the shop's dedicated text process layer
/// (<see cref="LayerName"/> = "L1", ACI 11 = pink/salmon) shows in the importer and is mapped to the
/// non-cut L1 graph — alongside cut = yellow/"Big Graph" and mark = blue/"Mid Graph". The font starts
/// <see cref="StartHeight"/>" tall and shrinks only so the widest line fits within the part's width.
/// </summary>
public static class PartLabel
{
    /// <summary>The shop's text process layer (NcStudio "L1" custom graph).</summary>
    public const string LayerName = "L1";
    /// <summary>ACI 11 (pink/salmon) — the colour the shop maps to the non-cut L1 graph in NcEditor.</summary>
    public const short LayerColor = 11;

    private const double StartHeight = 0.75;  // inches — the starting (max) cap height
    private const double MinHeight   = 0.1;   // floor so a tiny part still gets a (small) label
    private const double GapFactor   = 0.6;   // clear space above the part, as a fraction of the height
    private const double LineFactor  = 1.5;   // line-to-line pitch, as a fraction of the height

    /// <summary>Adds the label to <paramref name="geo"/> in place. No-op when there is nothing to say.</summary>
    public static void AddTo(CutGeometry geo, int? quantity, string? material, double thickness)
    {
        var lines = BuildLines(quantity, material, thickness);
        if (lines.Count == 0) return;

        var (minX, _, maxX, maxY) = Bounds(geo);
        double partW = maxX - minX;
        if (partW <= 0) return;
        double centerX = (minX + maxX) / 2.0;

        double height = ChooseHeight(partW, lines);
        double u      = height / StrokeFont.CapH;   // one cell unit in inches

        if (!geo.Layers.Any(l => l.Name.Equals(LayerName, StringComparison.OrdinalIgnoreCase)))
            geo.Layers.Add(new CutLayer { Name = LayerName, Color = LayerColor });

        // Stack above the part: the LAST line sits nearest the part, earlier lines above it.
        double baselineY = maxY + GapFactor * height;
        for (int i = lines.Count - 1; i >= 0; i--)
        {
            var line   = lines[i];
            double startX = centerX - StrokeFont.WidthUnits(line) * u / 2.0;
            foreach (var stroke in StrokeFont.Strokes(line))
                geo.Entities.Add(CutEntity.Polyline(LayerName, closed: false,
                    stroke.Select(p => new CutVertex(startX + p.X * u, baselineY + p.Y * u))));
            baselineY += LineFactor * height;
        }
    }

    /// <summary>Start at 1"; shrink only so the widest line fits within the part's width.</summary>
    internal static double ChooseHeight(double partW, IReadOnlyList<string> lines)
    {
        double widestUnits = lines.Count == 0 ? 0 : lines.Max(StrokeFont.WidthUnits);
        if (widestUnits <= 0) return StartHeight;
        // width(in) = widestUnits * (height / CapH) must be <= partW
        double maxByWidth = partW * StrokeFont.CapH / widestUnits;
        return Math.Max(MinHeight, Math.Min(StartHeight, maxByWidth));
    }

    /// <summary>The label lines (upper-cased for the stroke font): "XN" (quantity, when supplied) over
    /// "MATERIAL THICKNESS". Empty when there's nothing to say.</summary>
    internal static List<string> BuildLines(int? quantity, string? material, double thickness)
    {
        var lines = new List<string>();
        if (quantity is { } q && q >= 1) lines.Add($"X{q}");
        var mat = (material ?? "").Trim().ToUpperInvariant();
        var thk = thickness > 0 ? thickness.ToString("0.####", CultureInfo.InvariantCulture) + "\"" : "";
        var matThk = string.Join(" ", new[] { mat, thk }.Where(s => s.Length > 0));
        if (matThk.Length > 0) lines.Add(matThk);
        return lines;
    }

    /// <summary>Cut-geometry bounds, ignoring anything already on the no-cut Notes layer (so a label
    /// never measures itself).</summary>
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
            if (e.Layer.Equals(LayerName, StringComparison.OrdinalIgnoreCase)) continue;
            switch (e.Type)
            {
                case "polyline": if (e.Vertices != null) foreach (var v in e.Vertices) Acc(v.X, v.Y); break;
                case "line":     Acc(e.X1, e.Y1); Acc(e.X2, e.Y2); break;
                case "circle":   Acc(e.Cx - e.R, e.Cy - e.R); Acc(e.Cx + e.R, e.Cy + e.R); break;
            }
        }
        if (minX > maxX) return (0, 0, 0, 0);
        return (minX, minY, maxX, maxY);
    }
}
