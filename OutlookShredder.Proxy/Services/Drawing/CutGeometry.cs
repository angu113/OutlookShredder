namespace OutlookShredder.Proxy.Services.Drawing;

/// <summary>
/// Layered 2D cut geometry — origin lower-left, Y up, units in the part's units.
/// Schema matches the netDxf writer (CUT outer profile + BEND score lines + optional
/// CIRCLE reliefs) so it can be emitted to DXF unchanged and drawn to PDF for preview.
/// </summary>
public sealed class CutGeometry
{
    public string Units { get; set; } = "in";
    public string Part { get; set; } = "part";
    public List<CutLayer> Layers { get; set; } = new();
    public List<CutEntity> Entities { get; set; } = new();
}

public sealed class CutLayer
{
    public string Name { get; set; } = "CUT";
    public short Color { get; set; } = 7;   // AutoCAD ACI; 1=red, 5=blue
}

public readonly record struct CutVertex(double X, double Y, double Bulge = 0.0);

public sealed class CutEntity
{
    public string Type { get; set; } = "";   // polyline | line | circle | text
    public string Layer { get; set; } = "CUT";

    // polyline
    public bool Closed { get; set; }
    public List<CutVertex>? Vertices { get; set; }

    // line
    public double X1 { get; set; }
    public double Y1 { get; set; }
    public double X2 { get; set; }
    public double Y2 { get; set; }

    // circle (Cx/Cy also serve as the insertion point for text)
    public double Cx { get; set; }
    public double Cy { get; set; }
    public double R { get; set; }

    // text (annotation): content + height; position reuses Cx/Cy
    public string? Text { get; set; }
    public double Height { get; set; }

    public static CutEntity Polyline(string layer, bool closed, IEnumerable<CutVertex> verts)
        => new() { Type = "polyline", Layer = layer, Closed = closed, Vertices = verts.ToList() };

    public static CutEntity Line(string layer, double x1, double y1, double x2, double y2)
        => new() { Type = "line", Layer = layer, X1 = x1, Y1 = y1, X2 = x2, Y2 = y2 };

    public static CutEntity Circle(string layer, double cx, double cy, double r)
        => new() { Type = "circle", Layer = layer, Cx = cx, Cy = cy, R = r };

    /// <summary>A single line of annotation text, bottom-centered at (x, y).</summary>
    public static CutEntity Label(string layer, string text, double x, double y, double height)
        => new() { Type = "text", Layer = layer, Text = text, Cx = x, Cy = y, Height = height };
}
