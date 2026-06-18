using System.Globalization;
using OutlookShredder.Proxy.Models.Drawing;

namespace OutlookShredder.Proxy.Services.Drawing;

/// <summary>
/// Develops a <see cref="PartSpec"/> into a flat pattern: resolves inside/outside dimensions,
/// runs the bend math, and produces the cut geometry + a radiused cross-section profile.
/// Implements U-channel (2 same-direction bends), Z-channel (2 opposing bends) and L-angle
/// (1 bend). The flat blank is always a rectangle with one bend centreline per bend.
/// </summary>
public static class FlatPattern
{
    /// <summary>DXF layer for cut lines (the CNC's expected name).</summary>
    public const string CutLayer = "Big Graph";
    /// <summary>DXF layer for bend lines / sheet marking (the CNC's expected name).</summary>
    public const string BendLayer = "Mid Graph";
    /// <summary>ACI colour for the CUT layer — yellow (2), matching Weihong NcStudio's "Big Graph".</summary>
    public const short CutColor = 2;
    /// <summary>ACI colour for the BEND / MARK layer — blue (5), matching NcStudio's "Mid Graph".</summary>
    public const short BendColor = 5;

    /// <summary>
    /// Develops a part and bakes a shop cutting-aid label (quantity, material, thickness) onto its cut
    /// geometry on the no-cut <see cref="PartLabel.LayerName"/> layer — so every DXF the engine emits
    /// carries it. <paramref name="quantity"/> defaults to 1 (the "xN" line is shown only for N &gt; 1);
    /// callers with an order quantity (e.g. a picking-slip FAB note) pass it through (null = no "xN"
    /// line, for the design wizard). The PDF renderer ignores text entities, so the drawing is unaffected.
    /// </summary>
    public static FlatPatternResult Develop(PartSpec spec, int? quantity = null)
    {
        var fp = DevelopShape(spec);
        PartLabel.AddTo(fp.Cut, quantity, spec.Material, spec.Thickness);
        PolishLabel.AddTo(fp.Cut, spec.PolishDirection);
        return fp;
    }

    private static FlatPatternResult DevelopShape(PartSpec spec) => spec.Type switch
    {
        PartType.UChannel    => DevelopChannel(spec, isZ: false),
        PartType.ZChannel    => DevelopChannel(spec, isZ: true),
        PartType.LAngle      => DevelopLAngle(spec),
        PartType.FlitchPlate => DevelopPlate(spec),
        PartType.BasePlate   => DevelopPlate(spec),
        PartType.Pan         => DevelopPan(spec),
        PartType.PaddleBlind => DevelopPaddleBlind(spec),
        PartType.Column      => DevelopColumn(spec),
        _ => throw new NotSupportedException($"Part type {spec.Type} is not implemented yet."),
    };

    // ── Structural column — two plate flat patterns (one DXF) + a dimensioned elevation ──
    private static FlatPatternResult DevelopColumn(PartSpec spec)
    {
        double baseT = spec.BaseThickness, bearT = spec.BearingThickness;
        double tubeLen = spec.ColumnFullHeight - baseT - bearT;

        // Each plate reuses the plate developer; thickness only labels them (plates are flat-cut).
        var baseSpec = new PartSpec
        {
            Type = PartType.BasePlate, Length = spec.BaseLength, Width = spec.BaseWidth,
            Thickness = baseT, Holes = spec.BaseHoles, Material = spec.Material, Units = spec.Units,
        };
        var bearSpec = new PartSpec
        {
            Type = PartType.BasePlate, Length = spec.BearingLength, Width = spec.BearingWidth,
            Thickness = bearT, Holes = spec.BearingHoles, Material = spec.Material, Units = spec.Units,
        };
        var baseRes = DevelopPlate(baseSpec);
        var bearRes = DevelopPlate(bearSpec);

        // Merge both flat patterns into ONE cut drawing: base at origin, bearing offset to the right.
        double gap = Math.Max(2.0, spec.BaseWidth * 0.2);
        double xOff = baseRes.FlatWidth + gap;
        var entities = new List<CutEntity>(baseRes.Cut.Entities);
        foreach (var e in bearRes.Cut.Entities) entities.Add(OffsetEntity(e, xOff, 0));

        // Tube cross-section etched (marking layer — not cut) on the centre of each plate to
        // locate the weld. Centred since the plates are welded centred on the column. Square/rect
        // tube gets filleted corners (radius = wall thickness) for steel; aluminium extrusion is sharp.
        bool aluminium = spec.Material.IndexOf("alum", StringComparison.OrdinalIgnoreCase) >= 0;
        double cornerR = (spec.ColumnShape != "round" && !aluminium) ? spec.ColumnWall : 0;
        entities.AddRange(TubeWeldOutline(spec.ColumnShape, spec.BaseLength / 2, spec.BaseWidth / 2,
            spec.ColumnOuterWidth, spec.ColumnOuterDepth, spec.ColumnWall, cornerR));
        entities.AddRange(TubeWeldOutline(spec.ColumnShape, xOff + spec.BearingLength / 2, spec.BearingWidth / 2,
            spec.ColumnOuterWidth, spec.ColumnOuterDepth, spec.ColumnWall, cornerR));

        string slug = $"column_{Trim(spec.ColumnFullHeight)}h_base{Trim(spec.BaseLength)}x{Trim(spec.BaseWidth)}_brg{Trim(spec.BearingLength)}x{Trim(spec.BearingWidth)}";
        var cut = new CutGeometry
        {
            Units = spec.Units, Part = slug,
            Layers = { new CutLayer { Name = CutLayer, Color = CutColor }, new CutLayer { Name = BendLayer, Color = BendColor } },
            Entities = entities,
        };

        return new FlatPatternResult
        {
            Spec = spec,
            WebOutside = 0, FlangeLeftOutside = 0, FlangeRightOutside = 0,
            FlatWidth = xOff + bearRes.FlatWidth,
            FlatHeight = Math.Max(baseRes.FlatHeight, bearRes.FlatHeight),
            BendLinesX = Array.Empty<double>(),
            Cut = cut, Profile = new(),
            IsColumn = true,
            ColumnFullHeight = spec.ColumnFullHeight,
            ColumnTubeLength = tubeLen,
            ColumnBaseThickness = baseT,
            ColumnBearingThickness = bearT,
            ColumnBaseL = spec.BaseLength, ColumnBaseW = spec.BaseWidth,
            ColumnBearingL = spec.BearingLength, ColumnBearingW = spec.BearingWidth,
            ColumnBaseHoles = baseRes.Holes,
            ColumnBearingHoles = bearRes.Holes,
            ColumnOuterWidth = spec.ColumnOuterWidth,
            ColumnOuterDepth = spec.ColumnOuterDepth,
            ColumnWall = spec.ColumnWall,
            ColumnTubeCornerR = cornerR,
            ColumnShape = spec.ColumnShape,
            ColumnLabel = spec.ColumnLabel,
            Summary = ColumnSummary(spec, tubeLen),
            Title = ColumnTitle(spec),
        };
    }

    private static CutEntity OffsetEntity(CutEntity e, double dx, double dy) => e.Type switch
    {
        "polyline" => CutEntity.Polyline(e.Layer, e.Closed,
            (e.Vertices ?? new()).Select(v => new CutVertex(v.X + dx, v.Y + dy, v.Bulge))),
        "line"   => CutEntity.Line(e.Layer, e.X1 + dx, e.Y1 + dy, e.X2 + dx, e.Y2 + dy),
        "circle" => CutEntity.Circle(e.Layer, e.Cx + dx, e.Cy + dy, e.R),
        _ => e,
    };

    /// <summary>
    /// The tube/pipe cross-section as marking-layer geometry centred at (cx,cy): outer profile plus
    /// the inner (wall) profile, so the etched plate shows exactly where the column lands for welding.
    /// Square/rect → rectangles (w × d); round → circles (w = OD). On the BEND/marking layer (etch,
    /// not cut).
    /// </summary>
    private static IEnumerable<CutEntity> TubeWeldOutline(string shape, double cx, double cy, double w, double d, double wall, double cornerR)
    {
        var list = new List<CutEntity>();
        if (shape == "round")
        {
            double ro = w / 2.0;
            if (ro <= 0) return list;
            list.Add(CutEntity.Circle(BendLayer, cx, cy, ro));
            if (ro - wall > 0.01) list.Add(CutEntity.Circle(BendLayer, cx, cy, ro - wall));
        }
        else
        {
            if (w <= 0 || d <= 0) return list;
            // Outer AND inner (bore) corners filleted to the wall radius for any non-aluminium tube;
            // sharp for an aluminium extrusion (cornerR == 0).
            list.Add(CutEntity.Polyline(BendLayer, closed: true, RectVerts(cx, cy, w / 2.0, d / 2.0, cornerR)));
            if (w / 2.0 - wall > 0.01 && d / 2.0 - wall > 0.01)
                list.Add(CutEntity.Polyline(BendLayer, closed: true, RectVerts(cx, cy, w / 2.0 - wall, d / 2.0 - wall, cornerR)));
        }
        return list;
    }

    /// <summary>Rectangle outline centred at (cx,cy), half-extents (hw,hh); corners filleted to radius r
    /// (r ≤ 0 = sharp), tessellated CCW.</summary>
    private static List<CutVertex> RectVerts(double cx, double cy, double hw, double hh, double r)
    {
        r = Math.Min(r, Math.Min(hw, hh) - 1e-6);
        if (r <= 0.001)
            return new List<CutVertex>
            {
                new(cx - hw, cy - hh), new(cx + hw, cy - hh),
                new(cx + hw, cy + hh), new(cx - hw, cy + hh),
            };
        var pts = new List<CutVertex>();
        const int seg = 6;
        void Arc(double acx, double acy, double a0, double a1)
        {
            for (int i = 0; i <= seg; i++)
            {
                double a = (a0 + (a1 - a0) * i / seg) * Math.PI / 180.0;
                pts.Add(new CutVertex(acx + r * Math.Cos(a), acy + r * Math.Sin(a)));
            }
        }
        Arc(cx - hw + r, cy - hh + r, 180, 270);   // bottom-left
        Arc(cx + hw - r, cy - hh + r, 270, 360);   // bottom-right
        Arc(cx + hw - r, cy + hh - r, 0, 90);      // top-right
        Arc(cx - hw + r, cy + hh - r, 90, 180);    // top-left
        return pts;
    }

    private static string ColumnTitle(PartSpec s)
    {
        string col = s.ColumnLabel.Length > 0
            ? s.ColumnLabel
            : ColShapeName(s.ColumnShape) + " " + ColDimText(s, DrawFormat.FracInch);
        return $"{MaterialPlain(s.Material)} Column — {DrawFormat.FracInch(s.ColumnFullHeight)} overall  ({col})".Trim();
    }

    private static string ColShapeName(string shape) => shape switch
    {
        "round" => "Pipe",
        "rect"  => "Rect Tube",
        _       => "Square Tube",
    };

    private static string ColDimText(PartSpec s, Func<double, string>? fmt = null)
    {
        fmt ??= v => N(v) + "\"";
        return s.ColumnShape switch
        {
            "round" => $"{fmt(s.ColumnOuterWidth)} OD",
            "rect"  => $"{fmt(s.ColumnOuterWidth)} x {fmt(s.ColumnOuterDepth)}",
            _       => $"{fmt(s.ColumnOuterWidth)} x {fmt(s.ColumnOuterWidth)}",
        };
    }

    private static string ColumnSummary(PartSpec s, double tubeLen)
    {
        string u = U(s.Units);
        string shapeName = s.ColumnShape switch { "round" => "pipe", "rect" => "rectangular tube", _ => "square tube" };
        string HoleNote(HoleSpec? h) =>
            h is { Diameter: > 0 } ? $"  ({(h.Count <= 0 ? 4 : h.Count)} holes {F(h.Diameter)}{u} dia, edge {F(h.EdgeDistance)}{u})" : "";
        var lines = new List<string>
        {
            $"Structural column  {MaterialPlain(s.Material)}" + (s.ColumnLabel.Length > 0 ? $"   [{s.ColumnLabel}]" : ""),
            $"Full height {F(s.ColumnFullHeight)}{u}  =  base plate {F(s.BaseThickness)}{u} + {shapeName} {F(tubeLen)}{u} + bearing plate {F(s.BearingThickness)}{u}",
            $"Base plate    {F(s.BaseLength)}{u} x {F(s.BaseWidth)}{u} x {F(s.BaseThickness)}{u}" + HoleNote(s.BaseHoles),
            $"Bearing plate {F(s.BearingLength)}{u} x {F(s.BearingWidth)}{u} x {F(s.BearingThickness)}{u}" + HoleNote(s.BearingHoles),
            $"Column        {shapeName} {ColDimText(s)}, wall {F(s.ColumnWall)}{u}  →  cut to {F(tubeLen)}{u}",
            "Plates welded centred on the column.  DXF contains both plate flat patterns.",
        };
        return string.Join("\n", lines);
    }

    // ── Pan (base + up to 4 walls, 90° bends, mitered corners with bend-root relief) ──
    private static FlatPatternResult DevelopPan(PartSpec spec)
    {
        double t = spec.Thickness, ri = spec.InsideRadius, k = spec.KFactor, angle = 90;
        bool wB = spec.PanBottom, wT = spec.PanTop, wL = spec.PanLeft, wR = spec.PanRight;

        // Inside → outside: each bounding wall adds T to that base dimension; an inside depth adds T (the base).
        double Lo = spec.LengthBasis == DimBasis.Inside ? spec.Length + (wL ? t : 0) + (wR ? t : 0) : spec.Length;
        double Wo = spec.WidthBasis  == DimBasis.Inside ? spec.Width  + (wB ? t : 0) + (wT ? t : 0) : spec.Width;
        double Do = spec.DepthBasis  == DimBasis.Inside ? spec.Depth + t : spec.Depth;

        double ossb = BendMath.Ossb(ri, t, angle);
        double bd = BendMath.BendDeduction(ri, t, k, angle, spec.MeasuredBendDeduction);
        double wallDev = Math.Max(0.1, Do - bd);   // flange flat length (keeps the total blank correct)

        // Base inner rectangle (between bend lines); each present wall extends the blank by wallDev.
        double bx0 = wL ? wallDev : 0, bx1 = bx0 + Lo, xMax = bx1 + (wR ? wallDev : 0);
        double by0 = wB ? wallDev : 0, by1 = by0 + Wo, yMax = by1 + (wT ? wallDev : 0);

        // Outer cut outline: base + present-wall flanges with the corner squares notched out so adjacent
        // walls fold without colliding. Each notched corner carries the DEVELOPED bend relief, cut into
        // the outline itself: the two notch edges pull back along the bend lines by the bend radius and
        // join with a scallop arc bulging into the base (cylinder cut at 45°, quartered + unrolled), so
        // the relief follows the bend radius as both 90° bends form. Tessellated to render in PDF + DXF.
        bool blN = wB && wL, brN = wB && wR, trN = wT && wR, tlN = wT && wL;
        double R = Math.Min(Math.Max(ri, t), wallDev * 0.7);   // relief reach, kept within the flange
        const double rt2 = 0.70710678;

        // Relief: the developed bend-corner cut. The notch edges pull back along the bend lines and sweep
        // INTO the base to a peak just past the bend root (the unrolled 45°-cut quarter cylinder), so the
        // relief follows the bend radius. (rx,ry) = bend root; (bdx,bdy) = bisector into the base.
        void Scallop(List<CutVertex> o, double x1, double y1, double x2, double y2, double rx, double ry, double bdx, double bdy)
        {
            double apx = rx + bdx * 1.2 * R, apy = ry + bdy * 1.2 * R;    // peak past the bend root, into the base
            const int n = 8;
            // Two concave halves meeting at the peak, each bowed toward the bend root (the central axis).
            for (int s2 = 0; s2 <= n; s2++)              // first notch edge -> peak (control = root)
            {
                double u = (double)s2 / n, v = 1 - u;
                o.Add(new CutVertex(v * v * x1 + 2 * v * u * rx + u * u * apx,
                                    v * v * y1 + 2 * v * u * ry + u * u * apy));
            }
            for (int s2 = 1; s2 <= n; s2++)              // peak -> second notch edge (control = root)
            {
                double u = (double)s2 / n, v = 1 - u;
                o.Add(new CutVertex(v * v * apx + 2 * v * u * rx + u * u * x2,
                                    v * v * apy + 2 * v * u * ry + u * u * y2));
            }
        }

        // Return (lip/hem) on each present wall — works for 2-, 3- and 4-sided pans. Each present wall's
        // outer edge pushes out by the lip's developed length; the lip is inset by rd (a 45° miter) ONLY
        // at corners where both adjacent walls are present (blN/brN/trN/tlN) so the two lips meet when
        // bent. Ends with no neighbouring wall stay square.
        ReturnSpec? ret = spec.PanReturn;
        double rd = 0;
        if (ret != null)
        {
            double retOuter = ret.Basis == DimBasis.Inside ? ret.Length + t : ret.Length;
            double retBd = 2 * OssbClamped(ri, t, ret.AngleDeg) - BendMath.BendAllowance(ri, t, k, ret.AngleDeg);
            rd = Math.Max(0.05, retOuter - retBd);
        }

        List<CutVertex> ol;
        if (rd > 0)
        {
            // Per-edge lip detour: where a wall is present, dip the outer edge out by rd, mitering only the
            // ends adjacent to another present wall. Corners/reliefs are identical to the no-return case.
            double bl = blN ? rd : 0, br = brN ? rd : 0, tr = trN ? rd : 0, tl = tlN ? rd : 0;
            double xBL = blN ? bx0 : 0, xBR = brN ? bx1 : xMax;   // bottom edge span
            double xTL = tlN ? bx0 : 0, xTR = trN ? bx1 : xMax;   // top edge span
            double yRB = brN ? by0 : 0, yRT = trN ? by1 : yMax;   // right edge span
            double yLB = blN ? by0 : 0, yLT = tlN ? by1 : yMax;   // left edge span

            ol = new();
            // Bottom edge (left → right)
            if (wB) { ol.Add(new(xBL, 0)); ol.Add(new(xBL + bl, -rd)); ol.Add(new(xBR - br, -rd)); ol.Add(new(xBR, 0)); }
            else    { ol.Add(new(xBL, 0)); ol.Add(new(xBR, 0)); }
            if (brN) { Scallop(ol, bx1, by0 - R, bx1 + R, by0, bx1, by0, -rt2, rt2); ol.Add(new(xMax, by0)); }
            // Right edge (bottom → top)
            if (wR) { ol.Add(new(xMax, yRB)); ol.Add(new(xMax + rd, yRB + br)); ol.Add(new(xMax + rd, yRT - tr)); ol.Add(new(xMax, yRT)); }
            else    { ol.Add(new(xMax, yRT)); }
            if (trN) { Scallop(ol, bx1 + R, by1, bx1, by1 + R, bx1, by1, -rt2, -rt2); ol.Add(new(bx1, yMax)); }
            // Top edge (right → left)
            if (wT) { ol.Add(new(xTR, yMax)); ol.Add(new(xTR - tr, yMax + rd)); ol.Add(new(xTL + tl, yMax + rd)); ol.Add(new(xTL, yMax)); }
            else    { ol.Add(new(xTL, yMax)); }
            if (tlN) { Scallop(ol, bx0, by1 + R, bx0 - R, by1, bx0, by1, rt2, -rt2); ol.Add(new(0, by1)); }
            // Left edge (top → bottom)
            if (wL) { ol.Add(new(0, yLT)); ol.Add(new(-rd, yLT - tl)); ol.Add(new(-rd, yLB + bl)); ol.Add(new(0, yLB)); }
            else    { ol.Add(new(0, yLB)); }
            if (blN) Scallop(ol, bx0 - R, by0, bx0, by0 - R, bx0, by0, rt2, rt2);
        }
        else
        {
            ol = new() { new(blN ? bx0 : 0, 0), new(brN ? bx1 : xMax, 0) };
            if (brN) { Scallop(ol, bx1, by0 - R, bx1 + R, by0, bx1, by0, -rt2, rt2); ol.Add(new(xMax, by0)); }
            ol.Add(new(xMax, trN ? by1 : yMax));
            if (trN) { Scallop(ol, bx1 + R, by1, bx1, by1 + R, bx1, by1, -rt2, -rt2); ol.Add(new(bx1, yMax)); }
            ol.Add(new(tlN ? bx0 : 0, yMax));
            if (tlN) { Scallop(ol, bx0, by1 + R, bx0 - R, by1, bx0, by1, rt2, -rt2); ol.Add(new(0, by1)); }
            ol.Add(new(0, blN ? by0 : 0));
            if (blN) Scallop(ol, bx0 - R, by0, bx0, by0 - R, bx0, by0, rt2, rt2);
        }

        var entities = new List<CutEntity> { CutEntity.Polyline(CutLayer, closed: true, ol) };

        // Bend lines (only where a wall is present), spanning the base edge between the relieved corners.
        if (wB) entities.Add(CutEntity.Line(BendLayer, bx0, by0, bx1, by0));
        if (wT) entities.Add(CutEntity.Line(BendLayer, bx0, by1, bx1, by1));
        if (wL) entities.Add(CutEntity.Line(BendLayer, bx0, by0, bx0, by1));
        if (wR) entities.Add(CutEntity.Line(BendLayer, bx1, by0, bx1, by1));

        // Return crease bend lines at each present wall's outer edge (where the lip folds).
        if (rd > 0)
        {
            if (wB) entities.Add(CutEntity.Line(BendLayer, bx0, 0, bx1, 0));
            if (wT) entities.Add(CutEntity.Line(BendLayer, bx0, yMax, bx1, yMax));
            if (wL) entities.Add(CutEntity.Line(BendLayer, 0, by0, 0, by1));
            if (wR) entities.Add(CutEntity.Line(BendLayer, xMax, by0, xMax, by1));
        }

        // Cross-section profiles (radiused U, thickness shown like the channels). The flanges of each
        // section are the walls perpendicular to the cut: the width (side) section shows the long
        // (bottom/top) walls; the length (end) section shows the short (left/right) walls.
        var sideProfile = BuildPanSection(Wo, wB ? Do : t, wB && ret != null, wT ? Do : t, wT && ret != null, ret, t, ri);
        var endProfile  = BuildPanSection(Lo, wL ? Do : t, wL && ret != null, wR ? Do : t, wR && ret != null, ret, t, ri);

        string slug = $"pan_{Trim(Lo)}x{Trim(Wo)}x{Trim(Do)}";
        var cut = new CutGeometry
        {
            Units = spec.Units, Part = slug,
            Layers = { new CutLayer { Name = CutLayer, Color = CutColor }, new CutLayer { Name = BendLayer, Color = BendColor } },
            Entities = entities,
        };

        return new FlatPatternResult
        {
            Spec = spec, Ossb = ossb, BendDeduction = bd,
            WebOutside = 0, FlangeLeftOutside = 0, FlangeRightOutside = 0,
            // Blank extent includes the return lip on each present wall's outer edge.
            FlatWidth = xMax + (wL ? rd : 0) + (wR ? rd : 0),
            FlatHeight = yMax + (wB ? rd : 0) + (wT ? rd : 0),
            BendLinesX = Array.Empty<double>(),
            Cut = cut, Profile = new(),
            IsPan = true,
            PanBaseX0 = bx0, PanBaseX1 = bx1, PanBaseY0 = by0, PanBaseY1 = by1, PanWallDev = wallDev,
            PanDepth = Do,
            PanSideProfile = sideProfile,
            PanEndProfile = endProfile,
            Summary = PanSummary(spec, Lo, Wo, Do),
            Title = PlainTitle(spec),
        };
    }

    private static string PanSummary(PartSpec s, double Lo, double Wo, double Do)
    {
        string u = U(s.Units);
        int longN = (s.PanBottom ? 1 : 0) + (s.PanTop ? 1 : 0);
        int shortN = (s.PanLeft ? 1 : 0) + (s.PanRight ? 1 : 0);
        bool anyInside = s.LengthBasis == DimBasis.Inside || s.WidthBasis == DimBasis.Inside || s.DepthBasis == DimBasis.Inside;
        var lines = new List<string>
        {
            $"Pan  {s.Material}  (T={F(s.Thickness)}{u})",
            $"Base {F(Lo)}{u} x {F(Wo)}{u} OD, side flange {F(Do)}{u} deep OD",
            $"Side flanges: {longN} long + {shortN} short; mitered corners, bend relief {F(s.Thickness)}{u}",
        };
        if (anyInside)
        {
            string b(DimBasis db) => db == DimBasis.Inside ? "ID" : "OD";
            lines.Add($"Given: L {F(s.Length)} {b(s.LengthBasis)}, W {F(s.Width)} {b(s.WidthBasis)}, D {F(s.Depth)} {b(s.DepthBasis)}");
        }
        if (s.PanReturn is { } rp)
            lines.Add(rp.AngleDeg >= 170
                ? $"Return (all walls): hem {F(rp.Length)}{u} {rp.Direction.ToString().ToLowerInvariant()}, mitered corners"
                : $"Return (all walls): {F(rp.Length)}{u} @ {rp.AngleDeg:0.#}° {rp.Direction.ToString().ToLowerInvariant()}, mitered corners");
        if (s.Finish is FinishSide.Outside or FinishSide.Inside)
            lines.Add($"Finish on {s.Finish.ToString().ToLowerInvariant()} face");
        return string.Join("\n", lines);
    }

    // ── Paddle blind / spade ("frying pan") — solid disc + handle, no bends ──
    private static FlatPatternResult DevelopPaddleBlind(PartSpec spec)
    {
        double od = spec.PaddleOd, R = od / 2.0;
        double hw = Math.Min(spec.PaddleHandleWidth / 2.0, R * 0.95);   // handle half-width
        double C = spec.PaddleCenterToEnd;
        double cx = R, cy = R;                                          // disc centre (positive coords, origin lower-left)

        double xi = Math.Sqrt(Math.Max(1e-6, R * R - hw * hw));         // where the handle meets the disc
        double theta = Math.Atan2(hw, xi);
        double endCx = Math.Max(xi + hw * 0.5, C - hw);                 // centre of the rounded handle end

        var ol = new List<CutVertex>();
        const int discSeg = 64, endSeg = 16;
        // Disc major arc: from +theta CCW around the back to -theta (the part the handle doesn't take).
        for (int i = 0; i <= discSeg; i++)
        {
            double a = theta + (2 * Math.PI - 2 * theta) * i / discSeg;
            ol.Add(new CutVertex(cx + R * Math.Cos(a), cy + R * Math.Sin(a)));
        }
        ol.Add(new CutVertex(cx + endCx, cy - hw));                     // down the bottom handle edge
        // Rounded end: semicircle centred at (endCx,0), from -90° CCW through the tip to +90°.
        for (int i = 1; i < endSeg; i++)
        {
            double a = -Math.PI / 2 + Math.PI * i / endSeg;
            ol.Add(new CutVertex(cx + endCx + hw * Math.Cos(a), cy + hw * Math.Sin(a)));
        }
        ol.Add(new CutVertex(cx + endCx, cy + hw));                     // top of end (closes back to the first vertex)

        var entities = new List<CutEntity> { CutEntity.Polyline(CutLayer, closed: true, ol) };

        string slug = $"paddleblind_{spec.PaddleNps.Replace("/", "-")}in_cl{spec.PaddleClass}";
        var cut = new CutGeometry
        {
            Units = spec.Units, Part = slug,
            Layers = { new CutLayer { Name = CutLayer, Color = CutColor }, new CutLayer { Name = BendLayer, Color = BendColor } },
            Entities = entities,
        };

        return new FlatPatternResult
        {
            Spec = spec,
            WebOutside = 0, FlangeLeftOutside = 0, FlangeRightOutside = 0,
            FlatWidth = C + R, FlatHeight = od,
            BendLinesX = Array.Empty<double>(),
            Cut = cut, Profile = new(),
            IsPaddle = true,
            Summary = PaddleSummary(spec),
            Title = PlainTitle(spec),
        };
    }

    private static string PaddleSummary(PartSpec s)
    {
        string u = U(s.Units);
        return string.Join("\n", new[]
        {
            $"Paddle blind (spade / \"frying pan\")  {s.Material}  (T={F(s.Thickness)}{u})",
            $"NPS {s.PaddleNps}\" Class {s.PaddleClass} per ASME B16.48",
            $"Spade OD {F(s.PaddleOd)}{u}, handle {F(s.PaddleHandleWidth)}{u} wide x {F(s.PaddleCenterToEnd)}{u} centre-to-end",
        });
    }

    // ── Flat plate (Flitch / Base) — rectangle + bolt holes, no bends ────────
    private static FlatPatternResult DevelopPlate(PartSpec spec)
    {
        double L = spec.Length, W = spec.Width;
        var holes = ComputeHoles(spec, L, W);

        var entities = new List<CutEntity>
        {
            CutEntity.Polyline(CutLayer, closed: true, new[]
            {
                new CutVertex(0, 0), new CutVertex(L, 0), new CutVertex(L, W), new CutVertex(0, W),
            }),
        };
        foreach (var (hx, hy, dia) in holes)
            entities.Add(CutEntity.Circle(CutLayer, hx, hy, dia / 2.0));

        string slug = (spec.Type == PartType.FlitchPlate ? "flitch_" : "baseplate_") + $"{Trim(L)}x{Trim(W)}";
        var cut = new CutGeometry
        {
            Units = spec.Units,
            Part = slug,
            Layers = { new CutLayer { Name = CutLayer, Color = CutColor }, new CutLayer { Name = BendLayer, Color = BendColor } },
            Entities = entities,
        };

        return new FlatPatternResult
        {
            Spec = spec,
            WebOutside = 0, FlangeLeftOutside = 0, FlangeRightOutside = 0,
            FlatWidth = L, FlatHeight = W,
            BendLinesX = Array.Empty<double>(),
            Cut = cut,
            Profile = new(),
            IsPlate = true,
            Holes = holes,
            Summary = PlateSummary(spec, L, W, holes.Count),
            Title = PlainTitle(spec),
        };
    }

    private static List<(double x, double y, double dia)> ComputeHoles(PartSpec spec, double L, double W)
    {
        var list = new List<(double x, double y, double dia)>();
        var h = spec.Holes;
        if (h is null || h.Diameter <= 0) return list;

        if (h.Pattern == HolePattern.Corner)
        {
            double e = h.EdgeDistance > 0 ? h.EdgeDistance : 1.0;
            int n = h.Count <= 0 ? 4 : h.Count;
            var corners = new[] { (e, e), (L - e, e), (L - e, W - e), (e, W - e) };
            for (int i = 0; i < Math.Min(n, 4); i++)
                list.Add((corners[i].Item1, corners[i].Item2, h.Diameter));
        }
        else
        {
            // Flitch: two rows; first/last positions are paired (two holes), the middle follows
            // the pattern. Holes are evenly spaced between the captured end margins.
            double sp = h.Spacing > 0 ? h.Spacing : 16;
            double leftEnd  = h.LeftEndOffset  > 0 ? h.LeftEndOffset  : sp / 2.0;
            double rightEnd = h.RightEndOffset > 0 ? h.RightEndOffset : sp / 2.0;
            double topEdge = h.TopEdge > 0 ? h.TopEdge : W * 0.25;
            double botEdge = h.BottomEdge > 0 ? h.BottomEdge : W * 0.25;
            double rowTop = W - topEdge;   // top row, measured from the top edge
            double rowBot = botEdge;        // bottom row, measured from the bottom edge

            // First hole (LHS end pair) at the LHS offset, then EXACT `sp` spacing across the
            // length; the RHS end pair sits at the RHS offset (L - rightEnd) regardless, so the
            // final gap is just the leftover remainder. A hole still lands every `sp` even when
            // that leaves the end pair only an inch or two past the last staggered hole; we only
            // collapse the last interior hole into the end pair when the two would physically
            // overlap (centres closer than one hole diameter), never merely to tidy the gap.
            double first = leftEnd, last = L - rightEnd;
            var xs = new List<double>();
            if (last > first + 1e-6)
            {
                for (double x = first; x < last - 1e-6; x += sp) xs.Add(x);
                double mergeGap = h.Diameter > 0 ? h.Diameter : 1e-6;
                if (xs.Count > 0 && Math.Abs(xs[^1] - last) < mergeGap) xs[^1] = last;
                else xs.Add(last);
            }
            else
            {
                xs.Add(L / 2.0);   // plate too short for two end holes — one centred
            }

            for (int i = 0; i < xs.Count; i++)
            {
                double x = xs[i];
                bool ends = i == 0 || i == xs.Count - 1;
                if (h.Pattern == HolePattern.Paired || ends)
                {
                    list.Add((x, rowTop, h.Diameter));
                    list.Add((x, rowBot, h.Diameter));
                }
                else
                {
                    list.Add((x, i % 2 == 0 ? rowTop : rowBot, h.Diameter));   // staggered middle
                }
            }
        }
        return list;
    }

    private static string PlateSummary(PartSpec s, double L, double W, int holeCount)
    {
        string shape = s.Type == PartType.FlitchPlate ? "Flitch plate" : "Base plate";
        string u = U(s.Units);
        var lines = new List<string>
        {
            $"{shape}  {s.Material}  (T={F(s.Thickness)}{u})",
            $"Plate {F(L)}{u} x {F(W)}{u}",
        };
        if (s.Holes is { } h && holeCount > 0)
        {
            string pat = h.Pattern == HolePattern.Corner ? "corner" : h.Pattern.ToString().ToLowerInvariant();
            string detail = h.Pattern == HolePattern.Corner
                ? $"edge {F(h.EdgeDistance)}{u}"
                : $"{pat} @ {F(h.Spacing)}{u}";
            lines.Add($"Holes: {holeCount} x {F(h.Diameter)}{u} dia, {detail}");
        }
        return string.Join("\n", lines);
    }

    // ── Resolve per-bend angle + direction (from spec.Bends, else shape defaults) ──
    private static BendSpec[] ResolveBends(PartSpec spec, PartType type)
    {
        int n = type == PartType.LAngle ? 1 : 2;
        var arr = new BendSpec[n];
        for (int i = 0; i < n; i++)
        {
            if (spec.Bends is { } b && i < b.Count) arr[i] = b[i];
            else arr[i] = new BendSpec(spec.AngleDeg,
                (type == PartType.ZChannel && i == 1) ? BendDir.Down : BendDir.Up);
        }
        return arr;
    }

    // Outside dev-length of a return lip (inside basis runs to a free edge ⇒ +1T).
    private static double RetOuter(ReturnSpec r, double t) => r.Basis == DimBasis.Inside ? r.Length + t : r.Length;

    // Pan cross-section (a U whose two walls fold up) with optional return lips — developed through the
    // same centreline engine as the channels (no special radiused-U path). The base is a between-bends
    // segment, so it gets the same outer-face shrink as a channel web.
    private static List<(double x, double y)> BuildPanSection(
        double web, double wall1, bool ret1, double wall2, bool ret2, ReturnSpec? r, double t, double ri)
    {
        bool hasR1 = ret1 && r is not null;
        bool hasR2 = ret2 && r is not null;
        double ro  = r is not null ? RetOuter(r, t) : 0;
        var segs = new List<double>();
        var bends = new List<(double, BendDir, bool)>();
        void Seg(double len, (double, BendDir, bool)? b) { if (b is { } bb) bends.Add(bb); segs.Add(len); }
        if (hasR1) { Seg(ro, null); Seg(wall1, (r!.AngleDeg, r.Direction, true)); }
        else Seg(wall1, null);
        Seg(web, (90, BendDir.Up, false));
        int webIdx = segs.Count - 1;
        Seg(wall2, (90, BendDir.Up, false));
        if (hasR2) Seg(ro, (r!.AngleDeg, r.Direction, true));
        var sectionSegs = SectionSegs(segs.ToArray(), bends.Select(b => b.Item1).ToArray(), t);
        var (loop, _) = BuildOffsetProfile(sectionSegs, bends.ToArray(), t, ri, 0, sectionSegs[0], 0, -1, webIdx - 1);
        return loop;
    }

    // Tangent inset factor for a bend. tan(angle/2) is the fillet tangent for an acute/obtuse bend, but
    // it diverges as the faces fold parallel (→180°). A hem (≈180°) is a hairpin — a semicircle joining
    // two parallel faces — whose inset along the face is ~0 (the arc bulges perpendicular, faces stay
    // full length). So return 0 near 180°, and cap obtuse bends modestly.
    private static double TanHalf(double angleDeg)
        => angleDeg >= 170 ? 0.0 : Math.Min(2.5, Math.Tan(angleDeg * Math.PI / 360.0));

    // OSSB built on TanHalf so 90° returns are exact and 180° hems add (not consume) the lip in the flat.
    private static double OssbClamped(double ri, double t, double angleDeg) => (ri + t) * TanHalf(angleDeg);

    // Section-view segment spans. Every segment's OUTER faces sit (t/2)·tan(θ/2) PAST the centreline apex
    // at each BOUNDING bend (the sharp outer corner is outboard of the apex). Shrink each section span by
    // that per bounding bend so the FORMED outer dimensions equal the spec: a web (two bends) loses t, a
    // flange (one bend + free edge) loses t/2, a free edge loses nothing. The flat blank (UnrollChain)
    // keeps the full outer dev lengths — this only affects the drawn section + its dimension anchors.
    private static double[] SectionSegs(double[] segs, double[] bendAngles, double t)
    {
        var outv = (double[])segs.Clone();
        for (int i = 0; i < outv.Length; i++)
        {
            double shrink = 0;
            if (i >= 1)               shrink += 0.5 * t * TanHalf(bendAngles[i - 1]);   // bend before this segment
            if (i <= outv.Length - 2) shrink += 0.5 * t * TanHalf(bendAngles[i]);       // bend after this segment
            outv[i] = Math.Max(0.01, outv[i] - shrink);
        }
        return outv;
    }

    // Unrolls a segment chain to its flat width + each bend line's X (developed length).
    private static (double FlatWidth, double[] BendX) UnrollChain(double[] segOuter, double[] angles, double ri, double t, double k)
    {
        int m = angles.Length;
        var ossb = new double[m]; var ba = new double[m];
        for (int i = 0; i < m; i++) { ossb[i] = OssbClamped(ri, t, angles[i]); ba[i] = BendMath.BendAllowance(ri, t, k, angles[i]); }
        var bx = new double[m];
        double x = 0;
        for (int s = 0; s <= m; s++)
        {
            double straight = Math.Max(0, segOuter[s] - (s > 0 ? ossb[s - 1] : 0) - (s < m ? ossb[s] : 0));
            x += straight;
            if (s < m) { bx[s] = x + ba[s] / 2.0; x += ba[s]; }
        }
        return (x, bx);
    }

    // ── U-channel / Z-channel (+ optional returns on each flange) ────────────
    private static FlatPatternResult DevelopChannel(PartSpec spec, bool isZ)
    {
        double t = spec.Thickness, ri = spec.InsideRadius, k = spec.KFactor;

        // Inside → outside compensation: web spans both flanges (+2T); each flange runs to a free edge (+1T).
        double webO = spec.Web.Basis == DimBasis.Inside ? spec.Web.Value + 2 * t : spec.Web.Value;
        double flLO = spec.FlangeLeft.Basis == DimBasis.Inside ? spec.FlangeLeft.Value + t : spec.FlangeLeft.Value;
        double flRO = spec.FlangeRight.Basis == DimBasis.Inside ? spec.FlangeRight.Value + t : spec.FlangeRight.Value;
        var bs = ResolveBends(spec, isZ ? PartType.ZChannel : PartType.UChannel);

        // Build the left→right chain: [retL?, flL, web, flR, retR?].
        var segs = new List<double>();
        var bends = new List<(double Angle, BendDir Dir, bool IsReturn)>();
        void Seg(double len, (double, BendDir, bool)? bendBefore)
        { if (bendBefore is { } b) bends.Add(b); segs.Add(len); }

        if (spec.ReturnLeft is { } rl) { Seg(RetOuter(rl, t), null); Seg(flLO, (rl.AngleDeg, rl.Direction, true)); }
        else Seg(flLO, null);
        Seg(webO, (bs[0].AngleDeg, bs[0].Direction, false));
        int webIdx = segs.Count - 1;
        Seg(flRO, (bs[1].AngleDeg, bs[1].Direction, false));
        if (spec.ReturnRight is { } rr) Seg(RetOuter(rr, t), (rr.AngleDeg, rr.Direction, true));

        var angles = bends.Select(b => b.Angle).ToArray();
        var (flatWidth, bendsX) = UnrollChain(segs.ToArray(), angles, ri, t, k);
        double flatHeight = spec.Length;
        int anchorBendIdx = webIdx - 1;   // the flange↔web bend — keeps the web horizontal

        // One engine for every shape: develop the section by walking the bend centreline (offset ±t/2 to
        // the two faces), shrinking each span by its bounding-bend outer-corner offsets (see SectionSegs)
        // so every formed OUTER dimension equals the spec.
        var sectionSegs = SectionSegs(segs.ToArray(), angles, t);
        var (profile, sb) = BuildOffsetProfile(sectionSegs, bends.ToArray(), t, ri, 0, sectionSegs[0], 0, -1, anchorBendIdx);

        // Representative flange values for the summary.
        double bd0 = BendMath.BendDeduction(ri, t, k, bs[0].AngleDeg, spec.MeasuredBendDeduction);
        double bd1 = BendMath.BendDeduction(ri, t, k, bs[1].AngleDeg, spec.MeasuredBendDeduction);
        double ossb0 = BendMath.Ossb(ri, t, bs[0].AngleDeg), ba0 = BendMath.BendAllowance(ri, t, k, bs[0].AngleDeg);
        string slug = (isZ ? "zchannel_" : "uchannel_") + $"{Trim(webO)}x{Trim(flLO)}x{Trim(flatHeight)}";
        return Build(spec, ossb0, ba0, bd0, webO, flLO, flRO, flatWidth, flatHeight, bendsX, profile, slug, bs, new[] { bd0, bd1 }, sb);
    }

    // ── L-angle (1 bend, two legs; + optional returns on each leg) ───────────
    private static FlatPatternResult DevelopLAngle(PartSpec spec)
    {
        double t = spec.Thickness, ri = spec.InsideRadius, k = spec.KFactor;

        // legA = FlangeLeft, legB = FlangeRight. Each leg runs to a free edge (+1T inside).
        double legAo = spec.FlangeLeft.Basis == DimBasis.Inside ? spec.FlangeLeft.Value + t : spec.FlangeLeft.Value;
        double legBo = spec.FlangeRight.Basis == DimBasis.Inside ? spec.FlangeRight.Value + t : spec.FlangeRight.Value;
        var bs = ResolveBends(spec, PartType.LAngle);

        // Chain: [retB?, legB, legA, retA?]. legB = FlangeRight (ReturnRight); legA = FlangeLeft (ReturnLeft).
        var segs = new List<double>();
        var bends = new List<(double Angle, BendDir Dir, bool IsReturn)>();
        void Seg(double len, (double, BendDir, bool)? bendBefore)
        { if (bendBefore is { } b) bends.Add(b); segs.Add(len); }

        if (spec.ReturnRight is { } rb) { Seg(RetOuter(rb, t), null); Seg(legBo, (rb.AngleDeg, rb.Direction, true)); }
        else Seg(legBo, null);
        Seg(legAo, (bs[0].AngleDeg, bs[0].Direction, false));
        int baseIdx = segs.Count - 1;   // legA stays horizontal
        if (spec.ReturnLeft is { } ra) Seg(RetOuter(ra, t), (ra.AngleDeg, ra.Direction, true));

        var angles = bends.Select(b => b.Angle).ToArray();
        var (flatWidth, bendsX) = UnrollChain(segs.ToArray(), angles, ri, t, k);
        double flatHeight = spec.Length;
        int anchorBendIdx = baseIdx - 1;

        var sectionSegs = SectionSegs(segs.ToArray(), angles, t);
        var (profile, sb) = BuildOffsetProfile(sectionSegs, bends.ToArray(), t, ri, 0, sectionSegs[0], 0, -1, anchorBendIdx);

        double bd = BendMath.BendDeduction(ri, t, k, bs[0].AngleDeg, spec.MeasuredBendDeduction);
        double ossb = BendMath.Ossb(ri, t, bs[0].AngleDeg), ba = BendMath.BendAllowance(ri, t, k, bs[0].AngleDeg);
        string slug = $"angle_{Trim(legAo)}x{Trim(legBo)}x{Trim(flatHeight)}";
        // WebOutside unused for L; legA in FlangeLeftOutside, legB in FlangeRightOutside.
        return Build(spec, ossb, ba, bd, 0, legAo, legBo, flatWidth, flatHeight, bendsX, profile, slug, bs, new[] { bd }, sb);
    }

    // ── Shared result assembly (cut geometry rectangle + N bend lines) ───────
    private static FlatPatternResult Build(
        PartSpec spec, double ossb, double ba, double bd,
        double webO, double flLO, double flRO,
        double flatWidth, double flatHeight, double[] bends,
        List<(double x, double y)> profile, string slug,
        BendSpec[] resolvedBends, double[] bdEach, List<SectionBend> sectionBends)
    {
        var entities = new List<CutEntity>
        {
            CutEntity.Polyline(CutLayer, closed: true, new[]
            {
                new CutVertex(0, 0), new CutVertex(flatWidth, 0),
                new CutVertex(flatWidth, flatHeight), new CutVertex(0, flatHeight),
            }),
        };
        foreach (var bx in bends)
            entities.Add(CutEntity.Line(BendLayer, bx, 0, bx, flatHeight));

        var cut = new CutGeometry
        {
            Units = spec.Units,
            Part = slug,
            Layers = { new CutLayer { Name = CutLayer, Color = CutColor }, new CutLayer { Name = BendLayer, Color = BendColor } },
            Entities = entities,
        };

        return new FlatPatternResult
        {
            Spec = spec,
            Ossb = ossb,
            BendAllowance = ba,
            BendDeduction = bd,
            WebOutside = webO,
            FlangeLeftOutside = flLO,
            FlangeRightOutside = flRO,
            FlatWidth = flatWidth,
            FlatHeight = flatHeight,
            BendLinesX = bends,
            Cut = cut,
            Profile = profile,
            SectionBends = sectionBends,
            Summary = BuildSummary(spec, webO, flLO, flRO, resolvedBends, bdEach),
            Title = PlainTitle(spec),
        };
    }

    // ── Cross-section profiles ───────────────────────────────────────────────

    /// <summary>
    /// General radiused material loop for a segment chain. Walks the centreline (straight runs +
    /// 90° arcs of radius Ri+T/2 at each bend, turning +1 = left / -1 = right), then offsets ±T/2
    /// along the local normal to form the two faces. Handles same- and opposing-direction bends.
    /// </summary>
    private static (List<(double x, double y)> Loop, List<SectionBend> Bends) BuildOffsetProfile(
        double[] segs, (double AngleDeg, BendDir Dir, bool IsReturn)[] bends, double t, double ri,
        double sx, double sy, double hx, double hy, int anchorBendIdx = 0)
    {
        double rc = ri + t / 2.0;
        const int K = 6;
        var c = new List<(double x, double y, double nx, double ny)>();
        var marks = new List<SectionBend>();
        double px = sx, py = sy;
        double hn = Math.Sqrt(hx * hx + hy * hy); hx /= hn; hy /= hn;
        c.Add((px, py, -hy, hx));

        // A 180° hem is a hairpin whose outer arc bulges PAST the wall's straight end by (rc + t/2). If we
        // lay the wall to its full outer length and then add the hairpin on top, the formed flange/wall
        // reads TALLER than its OD in the section + iso (the fold pokes outside the OD dimension). Pull the
        // wall's straight run back by that bulge so the fold apex lands exactly on the OD — matching the
        // pan iso (which folds the hem down from the wall top). The lip is the chain-extremity segment; the
        // wall is the hem bend's interior neighbour. This is a section/iso-display correction only — the
        // flat blank (UnrollChain) is untouched, so the lip's developed length is still fully added there.
        var hemPull = new double[segs.Length];
        for (int b = 0; b < bends.Length; b++)
        {
            if (!bends[b].IsReturn || bends[b].AngleDeg < 170) continue;
            int wall = b == 0 ? b + 1 : (b + 1 == segs.Length - 1 ? b : b + 1);
            hemPull[wall] += rc + t / 2.0;
        }

        // Centreline tangent consumed by the fillet at each adjacent bend. A hem (≈180°) is a hairpin —
        // the arc bulges perpendicular, so the inset along the face is ~0 (TanHalf handles this).
        double TanCut(int bendIdx) => bendIdx >= 0 && bendIdx < bends.Length
            ? rc * TanHalf(bends[bendIdx].AngleDeg) : 0;
        for (int i = 0; i < segs.Length; i++)
        {
            double cut = TanCut(i - 1) + TanCut(i);
            double straight = Math.Max(0.001, segs[i] - cut - hemPull[i]);
            px += hx * straight; py += hy * straight;
            c.Add((px, py, -hy, hx));
            if (i < segs.Length - 1)
            {
                var (angDeg, dir, isRet) = bends[i];
                int d = dir == BendDir.Up ? +1 : -1;
                double sweep = angDeg * Math.PI / 180.0;
                double inHx = hx, inHy = hy;
                // Sharp-corner apex (where the two straight runs meet) — the bend's reference point.
                double vx = px + hx * TanCut(i), vy = py + hy * TanCut(i);
                double cx = px + d * (-hy) * rc, cy = py + d * hx * rc;
                double a0 = Math.Atan2(py - cy, px - cx);
                for (int kk = 1; kk <= K; kk++)
                {
                    double a = a0 + d * sweep * (kk / (double)K);
                    px = cx + rc * Math.Cos(a); py = cy + rc * Math.Sin(a);
                    hx = -d * Math.Sin(a); hy = d * Math.Cos(a);
                    c.Add((px, py, -hy, hx));
                }
                marks.Add(new SectionBend(vx, vy, inHx, inHy, hx, hy, angDeg, dir, isRet));
            }
        }
        var loop = new List<(double x, double y)>();
        foreach (var p in c) loop.Add((p.x + p.nx * t / 2, p.y + p.ny * t / 2));
        for (int i = c.Count - 1; i >= 0; i--) loop.Add((c[i].x - c[i].nx * t / 2, c[i].y - c[i].ny * t / 2));

        // Keep the web flat and horizontal: rotate the whole section so the segment leaving the
        // anchor bend (the flange↔web bend) runs along +x, and translate that bend's corner to the
        // origin. Flanges/returns then orient around the fixed horizontal web. (Standard 90° part with
        // no prefix return ⇒ anchor bend 0, already horizontal ⇒ no-op.)
        if (marks.Count > 0)
        {
            int ai = Math.Clamp(anchorBendIdx, 0, marks.Count - 1);
            double ax = marks[ai].X, ay = marks[ai].Y;
            double phi = Math.Atan2(marks[ai].OutHy, marks[ai].OutHx);
            double cs = Math.Cos(-phi), sn = Math.Sin(-phi);
            (double x, double y) Txp(double x, double y) { double dx = x - ax, dy = y - ay; return (dx * cs - dy * sn, dx * sn + dy * cs); }
            (double x, double y) Txv(double x, double y) => (x * cs - y * sn, x * sn + y * cs);
            for (int i = 0; i < loop.Count; i++) loop[i] = Txp(loop[i].x, loop[i].y);
            for (int i = 0; i < marks.Count; i++)
            {
                var m = marks[i];
                var (nx, ny) = Txp(m.X, m.Y);
                var (ih, iv) = Txv(m.InHx, m.InHy);
                var (oh, ov) = Txv(m.OutHx, m.OutHy);
                marks[i] = m with { X = nx, Y = ny, InHx = ih, InHy = iv, OutHx = oh, OutHy = ov };
            }
        }
        return (loop, marks);
    }

    // ── Title / summary ──────────────────────────────────────────────────────
    private static string PlainTitle(PartSpec s)
    {
        string mat = MaterialPlain(s.Material);
        string thk = DrawFormat.ThicknessLabel(s.Material, s.Thickness);   // gauge for steel/SS, decimal for alum
        static string D(double v) => DrawFormat.FracInch(v);              // dims in fractional inches

        if (s.Type is PartType.FlitchPlate or PartType.BasePlate)
        {
            string plate = s.Type == PartType.FlitchPlate ? "Flitch Plate" : "Base Plate";
            return $"{mat} {plate} — {thk} thickness x {D(s.Width)} wide x {D(s.Length)} long".Trim();
        }

        if (s.Type == PartType.Pan)
            return $"{mat} Pan — {D(s.Length)} x {D(s.Width)} x {D(s.Depth)} deep x {thk} thickness".Trim();

        if (s.Type == PartType.PaddleBlind)
            return $"{mat} Paddle Blind — {s.PaddleNps}\" NPS #{s.PaddleClass} x {thk} thickness".Trim();

        string shape = s.Type switch
        {
            PartType.UChannel => "U Channel",
            PartType.LAngle   => "Angle",
            PartType.ZChannel => "Z Channel",
            _ => s.Type.ToString(),
        };

        if (s.Type == PartType.LAngle)
            return $"{mat} {shape} — {D(s.FlangeLeft.Value)} x {D(s.FlangeRight.Value)} Legs x {thk} thickness x {D(s.Length)} long".Trim();

        // U / Z — web + flanges; collapse equal flanges to one value, else show both.
        bool sym = Math.Abs(s.FlangeLeft.Value - s.FlangeRight.Value) < 1e-6;
        string flanges = sym ? $"{D(s.FlangeLeft.Value)} Flanges"
                             : $"{D(s.FlangeLeft.Value)}/{D(s.FlangeRight.Value)} Flanges";
        return $"{mat} {shape} — {D(s.Web.Value)} Web x {flanges} x {thk} thickness x {D(s.Length)} long".Trim();
    }

    private static string MaterialPlain(string m) => Bi.MaterialEn(m);

    private static string N(double v) => v.ToString("0.###", CultureInfo.InvariantCulture);
    // Inch dims read as 2" in the footnote summary; non-inch units keep their literal suffix (e.g. 50mm).
    private static string U(string units) => units.Equals("in", StringComparison.OrdinalIgnoreCase) ? "\"" : units;

    private static string BuildSummary(PartSpec s, double webO, double flLO, double flRO, BendSpec[] bends, double[] bdEach)
    {
        string basis = (s.Web.Basis == s.FlangeLeft.Basis && s.FlangeLeft.Basis == s.FlangeRight.Basis)
            ? s.Web.Basis.ToString().ToLowerInvariant() : "mixed";
        string u = U(s.Units);
        string shape = s.Type switch
        {
            PartType.UChannel => "U-channel", PartType.LAngle => "L-angle",
            PartType.ZChannel => "Z-channel", _ => s.Type.ToString(),
        };
        string outside = s.Type == PartType.LAngle
            ? $"flange 1 {F(flLO)}{u}, flange 2 {F(flRO)}{u}"
            : $"flange 1 {F(flLO)}{u}, web {F(webO)}{u}, flange 2 {F(flRO)}{u}";
        var lines = new List<string>
        {
            $"{shape}  {s.Material}  (T={F(s.Thickness)}{u})",
            $"Basis: {basis}  →  outside {outside}",
        };
        if (s.AnglesAnnotated)
        {
            lines.Add($"Bend: Ri {F(s.InsideRadius)}{u}, K {s.KFactor:0.##}");
            for (int i = 0; i < bends.Length; i++)
                // Show the ACTUAL angle between the faces (180 − bend-from-flat), matching the drawn callout.
                lines.Add($"  bend {i + 1}: {(180.0 - bends[i].AngleDeg):0.#}° {bends[i].Direction.ToString().ToLowerInvariant()}, BD {F(bdEach[i])}{u}");
        }
        else
        {
            lines.Add($"Bend: Ri {F(s.InsideRadius)}{u}, K {s.KFactor:0.##}, {bends[0].AngleDeg:0.#}°,  BD {F(bdEach[0])}{u}/bend (×{bends.Length})");
        }
        string? finishNote = s.Finish switch
        {
            FinishSide.Inside  => "Finish on inside face",
            FinishSide.Outside => "Finish on outside face",
            FinishSide.Top     => "Finish on top face",
            FinishSide.Bottom  => "Finish on bottom face",
            _ => null,
        };
        if (finishNote != null) lines.Add(finishNote);

        // Returns (lip / hem).
        string RetNote(string where, ReturnSpec r) => r.AngleDeg >= 170
            ? $"Return ({where}): hem {F(r.Length)}{u} {r.Direction.ToString().ToLowerInvariant()}"
            : $"Return ({where}): {F(r.Length)}{u} @ {r.AngleDeg:0.#}° {r.Direction.ToString().ToLowerInvariant()}";
        if (s.ReturnLeft is { } rL)  lines.Add(RetNote(s.Type == PartType.LAngle ? "leg A" : "flange 1", rL));
        if (s.ReturnRight is { } rR) lines.Add(RetNote(s.Type == PartType.LAngle ? "leg B" : "flange 2", rR));
        if (s.PanReturn is { } rP)   lines.Add(RetNote("all walls", rP));

        return string.Join("\n", lines);
    }

    private static string F(double v) => v.ToString("0.###", CultureInfo.InvariantCulture);
    private static string Trim(double v) => v.ToString("0.##", CultureInfo.InvariantCulture);
}

/// <summary>
/// A bend marker in cross-section (model) coords: the sharp-corner apex, the incoming/outgoing
/// centreline unit headings, and the bend's angle + direction. Drives the degree/arc callouts.
/// </summary>
public sealed record SectionBend(
    double X, double Y, double InHx, double InHy, double OutHx, double OutHy, double AngleDeg, BendDir Dir,
    bool IsReturn = false);

/// <summary>Result of developing a part — drives the DXF, the PDF render, and the UI summary.</summary>
public sealed class FlatPatternResult
{
    public required PartSpec Spec { get; init; }
    public double Ossb { get; init; }
    public double BendAllowance { get; init; }
    public double BendDeduction { get; init; }
    public double WebOutside { get; init; }
    public double FlangeLeftOutside { get; init; }
    public double FlangeRightOutside { get; init; }
    public double FlatWidth { get; init; }
    public double FlatHeight { get; init; }
    public required double[] BendLinesX { get; init; }
    public required CutGeometry Cut { get; init; }
    /// <summary>Radiused cross-section material loop (model coords) for the section + iso views.</summary>
    public List<(double x, double y)> Profile { get; init; } = new();
    /// <summary>Per-bend markers (centreline apex + headings + angle/dir) for drawing angle callouts.</summary>
    public List<SectionBend> SectionBends { get; init; } = new();
    /// <summary>True for flat plates (Flitch / Base) — drawn as a single top view, not 3 panels.</summary>
    public bool IsPlate { get; init; }
    /// <summary>True for pans — drawn as a single flat-pattern top view (cut + bend lines + reliefs).</summary>
    public bool IsPan { get; init; }
    /// <summary>True for paddle blinds / spades — drawn as a single disc-plus-handle face view.</summary>
    public bool IsPaddle { get; init; }
    // Pan base rectangle (between bend lines) + flange flat length, for dimensioning.
    public double PanBaseX0 { get; init; }
    public double PanBaseX1 { get; init; }
    public double PanBaseY0 { get; init; }
    public double PanBaseY1 { get; init; }
    public double PanWallDev { get; init; }
    public double PanDepth { get; init; }   // outside wall height
    public List<(double x, double y)> PanSideProfile { get; init; } = new();   // U across the width
    public List<(double x, double y)> PanEndProfile { get; init; } = new();    // U across the length
    /// <summary>Plate bolt holes (centre x, y, diameter) in plate coords.</summary>
    public List<(double x, double y, double dia)> Holes { get; init; } = new();

    // ── Structural column — two plate flats + a dimensioned side elevation ──
    /// <summary>True for structural columns — base + bearing plate flats and a column elevation.</summary>
    public bool IsColumn { get; init; }
    public double ColumnFullHeight { get; init; }
    public double ColumnTubeLength { get; init; }
    public double ColumnBaseThickness { get; init; }
    public double ColumnBearingThickness { get; init; }
    public double ColumnBaseL { get; init; }
    public double ColumnBaseW { get; init; }
    public double ColumnBearingL { get; init; }
    public double ColumnBearingW { get; init; }
    public List<(double x, double y, double dia)> ColumnBaseHoles { get; init; } = new();
    public List<(double x, double y, double dia)> ColumnBearingHoles { get; init; } = new();
    public double ColumnTubeCornerR { get; init; }
    public double ColumnOuterWidth { get; init; }
    public double ColumnOuterDepth { get; init; }
    public double ColumnWall { get; init; }
    public string ColumnShape { get; init; } = "square";
    public string ColumnLabel { get; init; } = "";

    public required string Summary { get; init; }
    public string Title { get; init; } = "";
}
