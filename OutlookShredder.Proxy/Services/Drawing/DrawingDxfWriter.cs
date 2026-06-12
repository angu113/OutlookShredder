using netDxf;
using netDxf.Entities;
using netDxf.Header;
using netDxf.Tables;
using netDxf.Units;

namespace OutlookShredder.Proxy.Services.Drawing;

/// <summary>
/// Emits a <see cref="CutGeometry"/> to DXF bytes via netDxf (3.x API: Polyline2D,
/// doc.Entities.Add). CUT layer carries the outer profile; BEND layer carries score lines.
/// </summary>
public static class DrawingDxfWriter
{
    public static byte[] Write(CutGeometry geo)
    {
        var doc = new DxfDocument(DxfVersion.AutoCad2010);
        doc.DrawingVariables.InsUnits = geo.Units.Equals("mm", StringComparison.OrdinalIgnoreCase)
            ? DrawingUnits.Millimeters
            : DrawingUnits.Inches;

        var layers = new Dictionary<string, Layer>(StringComparer.OrdinalIgnoreCase);
        foreach (var ld in geo.Layers)
        {
            var layer = new Layer(ld.Name) { Color = new AciColor(ld.Color) };
            doc.Layers.Add(layer);
            layers[ld.Name] = layer;
        }
        Layer LayerFor(string name) => layers.TryGetValue(name, out var l) ? l : Layer.Default;

        foreach (var e in geo.Entities)
        {
            var lyr = LayerFor(e.Layer);
            // Stamp the explicit ACI colour on the entity (not just the layer / ByLayer). Weihong
            // NcStudio sorts/imports by entity COLOUR — yellow = cut, blue = mark — so a ByLayer
            // entity (which carries no concrete colour) lands on the wrong process.
            var col = lyr.Color;
            switch (e.Type.ToLowerInvariant())
            {
                case "polyline":
                {
                    var verts = (e.Vertices ?? new())
                        .Select(v => new Polyline2DVertex(v.X, v.Y) { Bulge = v.Bulge })
                        .ToList();
                    doc.Entities.Add(new Polyline2D(verts, e.Closed) { Layer = lyr, Color = col });
                    break;
                }
                case "line":
                    doc.Entities.Add(new Line(new Vector2(e.X1, e.Y1), new Vector2(e.X2, e.Y2))
                    { Layer = lyr, Color = col });
                    break;
                case "circle":
                    doc.Entities.Add(new Circle(new Vector2(e.Cx, e.Cy), e.R) { Layer = lyr, Color = col });
                    break;
                case "text":
                    doc.Entities.Add(new Text(e.Text ?? "", new Vector2(e.Cx, e.Cy), e.Height)
                    { Layer = lyr, Color = col, Alignment = TextAlignment.BottomCenter });
                    break;
                default:
                    throw new NotSupportedException($"Unknown cut entity type '{e.Type}'.");
            }
        }

        using var ms = new MemoryStream();
        doc.Save(ms);
        return ms.ToArray();
    }
}
