using Microsoft.Extensions.Logging;
using OutlookShredder.Proxy.Services;

namespace OutlookShredder.Proxy.Services.Drawing;

/// <summary>
/// Develops one or more FAB-note part descriptions and lays their flat patterns out left-to-right in a
/// SINGLE DXF: each part is placed <see cref="GapInches"/> to the RIGHT of the previous part's right
/// edge, and every part is bottom-aligned to Y=0. Used to auto-generate <c>{HSK#}.dxf</c> for a
/// picking slip's deduped FAB notes so the shop has one cut file per slip.
/// </summary>
public static class FabDxfBuilder
{
    /// <summary>Clear space left between adjacent parts (inches), per the picking-slip CAD spec.</summary>
    private const double GapInches = 1.0;

    public sealed record Result(byte[] Dxf, List<string> Parts);

    /// <summary>
    /// Builds the combined DXF. Notes that fail to parse/develop are skipped (logged) rather than
    /// failing the whole file. Returns null when nothing developable was produced.
    /// </summary>
    public static Result? Build(IEnumerable<FabNote> notes, ILogger? log = null)
    {
        var (geo, parts) = Combine(notes, log);
        return geo is null ? null : new Result(DrawingDxfWriter.Write(geo), parts);
    }

    /// <summary>
    /// Lays the developed parts out into one <see cref="CutGeometry"/> (left-to-right, 1" apart,
    /// bottom-aligned to Y=0; the first part's left edge sits at X=0). Returns a null geometry when no
    /// note developed. Separated from <see cref="Build"/> so the layout can be asserted without
    /// re-reading the serialized DXF.
    /// </summary>
    internal static (CutGeometry? Geo, List<string> Parts) Combine(IEnumerable<FabNote> notes, ILogger? log = null)
    {
        var combined  = new CutGeometry { Units = "in", Part = "fab" };
        var layerSeen = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        var parts     = new List<string>();
        double cursorX = 0;

        foreach (var note in notes)
        {
            CutGeometry geo;
            // Develop with the note's quantity so the engine bakes the "xN / material / thickness" label
            // (on the no-cut Notes layer) into each part; merged + translated below like any other entity.
            try { geo = FlatPattern.Develop(DrawingTextParser.Parse(note.Desc), note.Qty).Cut; }
            catch (Exception ex)
            {
                log?.LogWarning(ex, "[FAB-DXF] skipped un-developable note '{Desc}'", note.Desc);
                continue;
            }

            if (geo.Entities.Count == 0) continue;

            // Merge layers by name (cut="Big Graph"/yellow, bend="Mid Graph"/blue, notes="Notes") so all
            // parts share the same layers in the combined file.
            foreach (var ly in geo.Layers)
                if (layerSeen.Add(ly.Name))
                    combined.Layers.Add(new CutLayer { Name = ly.Name, Color = ly.Color });

            var (minX, minY, maxX, _) = Bounds(geo);
            double dx = cursorX - minX;   // slide left edge to the running cursor
            double dy = -minY;            // drop bottom edge to Y=0

            foreach (var e in geo.Entities)
                combined.Entities.Add(Translate(e, dx, dy));

            cursorX += (maxX - minX) + GapInches;
            parts.Add(string.IsNullOrEmpty(geo.Part) ? "part" : geo.Part);
        }

        return combined.Entities.Count == 0 ? (null, parts) : (combined, parts);
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
                case "polyline":
                    foreach (var v in e.Vertices ?? new()) Acc(v.X, v.Y);
                    break;
                case "line":
                    Acc(e.X1, e.Y1); Acc(e.X2, e.Y2);
                    break;
                case "circle":
                    Acc(e.Cx - e.R, e.Cy - e.R); Acc(e.Cx + e.R, e.Cy + e.R);
                    break;
            }
        }
        if (minX > maxX) return (0, 0, 0, 0);   // no measurable geometry
        return (minX, minY, maxX, maxY);
    }

    private static CutEntity Translate(CutEntity e, double dx, double dy) => e.Type switch
    {
        "polyline" => CutEntity.Polyline(e.Layer, e.Closed,
                          (e.Vertices ?? new()).Select(v => new CutVertex(v.X + dx, v.Y + dy, v.Bulge))),
        "line"     => CutEntity.Line(e.Layer, e.X1 + dx, e.Y1 + dy, e.X2 + dx, e.Y2 + dy),
        "circle"   => CutEntity.Circle(e.Layer, e.Cx + dx, e.Cy + dy, e.R),
        "text"     => CutEntity.Label(e.Layer, e.Text ?? "", e.Cx + dx, e.Cy + dy, e.Height),
        _          => e,
    };
}
