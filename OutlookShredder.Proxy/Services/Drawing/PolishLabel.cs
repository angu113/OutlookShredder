using OutlookShredder.Proxy.Models.Drawing;

namespace OutlookShredder.Proxy.Services.Drawing;

/// <summary>
/// Polish / grain-direction annotation baked into a part's cut geometry: a double-headed arrow along
/// the chosen axis on the part, plus a "DIRECCION DE PULIDO" label to the right of the part. Like
/// <see cref="PartLabel"/>, it is drawn as SINGLE-STROKE GEOMETRY on the shop's non-cut process layer
/// (<see cref="PartLabel.LayerName"/> = "L1", ACI 11) rather than DXF TEXT — CAM importers drop
/// TEXT/MTEXT on import. The label is ASCII ("DIRECCION", no accent) because <see cref="StrokeFont"/>
/// has no accented glyphs; the accented "Dirección de pulido" appears on the PDF, which uses a real
/// font. No-op when the direction is unset.
///
/// TODO: items 2 (finish) and 4 (polish) should later share ONE right-side placement manager so the
/// polish label never collides with the finish callout. Do NOT refactor the finish label here.
/// </summary>
public static class PolishLabel
{
    private const string LabelText   = "DIRECCION DE PULIDO";
    private const double LabelHeight = 0.35;   // inches — cap height (small, like PartLabel's lines)

    /// <summary>Adds the arrow + label to <paramref name="geo"/> in place. No-op when unset/empty.</summary>
    public static void AddTo(CutGeometry geo, PolishDirection dir)
    {
        if (dir == PolishDirection.None) return;

        // Bounds of the real geometry only (the no-cut L1 layer — incl. PartLabel — is ignored), so the
        // arrow centres on the part and the label sits just past its true right edge.
        var (minX, minY, maxX, maxY) = Bounds(geo);
        double w = maxX - minX, h = maxY - minY;
        if (w <= 0 || h <= 0) return;

        if (!geo.Layers.Any(l => l.Name.Equals(PartLabel.LayerName, StringComparison.OrdinalIgnoreCase)))
            geo.Layers.Add(new CutLayer { Name = PartLabel.LayerName, Color = PartLabel.LayerColor });

        bool vertical = dir == PolishDirection.Vertical;
        double cx = (minX + maxX) / 2.0, cy = (minY + maxY) / 2.0;
        double half = (vertical ? h : w) * 0.30;            // arrow spans ~60% of the part along the axis
        double head = Math.Max(0.05, Math.Min(w, h) * 0.06); // arrowhead leg length

        if (vertical)
        {
            AddLine(geo, cx, cy - half, cx, cy + half);
            AddLine(geo, cx, cy + half, cx - head, cy + half - head);
            AddLine(geo, cx, cy + half, cx + head, cy + half - head);
            AddLine(geo, cx, cy - half, cx - head, cy - half + head);
            AddLine(geo, cx, cy - half, cx + head, cy - half + head);
        }
        else
        {
            AddLine(geo, cx - half, cy, cx + half, cy);
            AddLine(geo, cx + half, cy, cx + half - head, cy + head);
            AddLine(geo, cx + half, cy, cx + half - head, cy - head);
            AddLine(geo, cx - half, cy, cx - half + head, cy + head);
            AddLine(geo, cx - half, cy, cx - half + head, cy - head);
        }

        // Label to the RIGHT of the part bbox, vertically centred, stroked glyphs on L1.
        double u = LabelHeight / StrokeFont.CapH;
        double startX = maxX + Math.Max(0.25, w * 0.04);
        double baselineY = cy - LabelHeight / 2.0;
        foreach (var stroke in StrokeFont.Strokes(LabelText))
            geo.Entities.Add(CutEntity.Polyline(PartLabel.LayerName, closed: false,
                stroke.Select(p => new CutVertex(startX + p.X * u, baselineY + p.Y * u))));
    }

    private static void AddLine(CutGeometry geo, double x1, double y1, double x2, double y2) =>
        geo.Entities.Add(CutEntity.Line(PartLabel.LayerName, x1, y1, x2, y2));

    /// <summary>Cut-geometry bounds, ignoring the no-cut L1 layer (so the part — not its labels —
    /// drives placement). Mirrors <see cref="PartLabel"/>'s bounds.</summary>
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
            if (e.Layer.Equals(PartLabel.LayerName, StringComparison.OrdinalIgnoreCase)) continue;
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
