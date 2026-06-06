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

    public static FlatPatternResult Develop(PartSpec spec) => spec.Type switch
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

        string slug = $"column_{Trim(spec.ColumnFullHeight)}h_base{Trim(spec.BaseLength)}x{Trim(spec.BaseWidth)}_brg{Trim(spec.BearingLength)}x{Trim(spec.BearingWidth)}";
        var cut = new CutGeometry
        {
            Units = spec.Units, Part = slug,
            Layers = { new CutLayer { Name = CutLayer, Color = 1 }, new CutLayer { Name = BendLayer, Color = 5 } },
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

    private static string ColumnTitle(PartSpec s)
    {
        string col = s.ColumnLabel.Length > 0
            ? s.ColumnLabel
            : ColShapeName(s.ColumnShape) + " " + ColDimText(s);
        return $"{MaterialPlain(s.Material)} Column — {N(s.ColumnFullHeight)}\" overall  ({col})".Trim();
    }

    private static string ColShapeName(string shape) => shape switch
    {
        "round" => "Pipe",
        "rect"  => "Rect Tube",
        _       => "Square Tube",
    };

    private static string ColDimText(PartSpec s) => s.ColumnShape switch
    {
        "round" => $"{N(s.ColumnOuterWidth)}\" OD",
        "rect"  => $"{N(s.ColumnOuterWidth)}\" x {N(s.ColumnOuterDepth)}\"",
        _       => $"{N(s.ColumnOuterWidth)}\" x {N(s.ColumnOuterWidth)}\"",
    };

    private static string ColumnSummary(PartSpec s, double tubeLen)
    {
        string u = s.Units;
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

        var ol = new List<CutVertex> { new(blN ? bx0 : 0, 0), new(brN ? bx1 : xMax, 0) };
        if (brN) { Scallop(ol, bx1, by0 - R, bx1 + R, by0, bx1, by0, -rt2, rt2); ol.Add(new(xMax, by0)); }
        ol.Add(new(xMax, trN ? by1 : yMax));
        if (trN) { Scallop(ol, bx1 + R, by1, bx1, by1 + R, bx1, by1, -rt2, -rt2); ol.Add(new(bx1, yMax)); }
        ol.Add(new(tlN ? bx0 : 0, yMax));
        if (tlN) { Scallop(ol, bx0, by1 + R, bx0 - R, by1, bx0, by1, rt2, -rt2); ol.Add(new(0, by1)); }
        ol.Add(new(0, blN ? by0 : 0));
        if (blN) Scallop(ol, bx0 - R, by0, bx0, by0 - R, bx0, by0, rt2, rt2);

        var entities = new List<CutEntity> { CutEntity.Polyline(CutLayer, closed: true, ol) };

        // Bend lines (only where a wall is present), spanning the base edge between the relieved corners.
        if (wB) entities.Add(CutEntity.Line(BendLayer, bx0, by0, bx1, by0));
        if (wT) entities.Add(CutEntity.Line(BendLayer, bx0, by1, bx1, by1));
        if (wL) entities.Add(CutEntity.Line(BendLayer, bx0, by0, bx0, by1));
        if (wR) entities.Add(CutEntity.Line(BendLayer, bx1, by0, bx1, by1));

        // Cross-section profiles (radiused U, thickness shown like the channels). The flanges of each
        // section are the walls perpendicular to the cut: the width (side) section shows the long
        // (bottom/top) walls; the length (end) section shows the short (left/right) walls.
        var sideProfile = BuildRadiusedU(Wo, wB ? Do : t, wT ? Do : t, t, ri);
        var endProfile  = BuildRadiusedU(Lo, wL ? Do : t, wR ? Do : t, t, ri);

        string slug = $"pan_{Trim(Lo)}x{Trim(Wo)}x{Trim(Do)}";
        var cut = new CutGeometry
        {
            Units = spec.Units, Part = slug,
            Layers = { new CutLayer { Name = CutLayer, Color = 1 }, new CutLayer { Name = BendLayer, Color = 5 } },
            Entities = entities,
        };

        return new FlatPatternResult
        {
            Spec = spec, Ossb = ossb, BendDeduction = bd,
            WebOutside = 0, FlangeLeftOutside = 0, FlangeRightOutside = 0,
            FlatWidth = xMax, FlatHeight = yMax,
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
        string u = s.Units;
        int longN = (s.PanBottom ? 1 : 0) + (s.PanTop ? 1 : 0);
        int shortN = (s.PanLeft ? 1 : 0) + (s.PanRight ? 1 : 0);
        bool anyInside = s.LengthBasis == DimBasis.Inside || s.WidthBasis == DimBasis.Inside || s.DepthBasis == DimBasis.Inside;
        var lines = new List<string>
        {
            $"Pan  {s.Material}  (T={F(s.Thickness)}{u})",
            $"Base {F(Lo)}{u} x {F(Wo)}{u} OD, wall {F(Do)}{u} deep OD",
            $"Walls: {longN} long + {shortN} short; mitered corners, bend relief {F(s.Thickness)}{u}",
        };
        if (anyInside)
        {
            string b(DimBasis db) => db == DimBasis.Inside ? "ID" : "OD";
            lines.Add($"Given: L {F(s.Length)} {b(s.LengthBasis)}, W {F(s.Width)} {b(s.WidthBasis)}, D {F(s.Depth)} {b(s.DepthBasis)}");
        }
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
            Layers = { new CutLayer { Name = CutLayer, Color = 1 }, new CutLayer { Name = BendLayer, Color = 5 } },
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
        string u = s.Units;
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
            Layers = { new CutLayer { Name = CutLayer, Color = 1 }, new CutLayer { Name = BendLayer, Color = 5 } },
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
        string u = s.Units;
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

    // ── U-channel (2 bends same direction) / Z-channel (2 bends opposing) ────
    private static FlatPatternResult DevelopChannel(PartSpec spec, bool isZ)
    {
        double t = spec.Thickness, ri = spec.InsideRadius, k = spec.KFactor, angle = spec.AngleDeg;

        // Inside → outside compensation: web spans both flanges (+2T); each flange runs to a free edge (+1T).
        double webO = spec.Web.Basis == DimBasis.Inside ? spec.Web.Value + 2 * t : spec.Web.Value;
        double flLO = spec.FlangeLeft.Basis == DimBasis.Inside ? spec.FlangeLeft.Value + t : spec.FlangeLeft.Value;
        double flRO = spec.FlangeRight.Basis == DimBasis.Inside ? spec.FlangeRight.Value + t : spec.FlangeRight.Value;

        double ossb = BendMath.Ossb(ri, t, angle);
        double ba = BendMath.BendAllowance(ri, t, k, angle);
        double bd = BendMath.BendDeduction(ri, t, k, angle, spec.MeasuredBendDeduction);

        double flatWidth = flLO + webO + flRO - 2 * bd;
        double flatHeight = spec.Length;

        double leftFlangeFlat = flLO - ossb;
        double webFlat = webO - 2 * ossb;
        double bend1X = leftFlangeFlat + ba / 2.0;
        double bend2X = leftFlangeFlat + ba + webFlat + ba / 2.0;
        var bends = new[] { bend1X, bend2X };

        // Z folds its flanges in opposite directions; U folds them the same way.
        var profile = isZ
            ? BuildOffsetProfile(new[] { flLO, webO, flRO }, new[] { +1, -1 }, t, ri, 0, flLO, 0, -1)
            : BuildRadiusedU(webO, flLO, flRO, t, ri);

        string slug = (isZ ? "zchannel_" : "uchannel_") + $"{Trim(webO)}x{Trim(flLO)}x{Trim(flatHeight)}";
        return Build(spec, ossb, ba, bd, webO, flLO, flRO, flatWidth, flatHeight, bends, profile, slug);
    }

    // ── L-angle (1 bend, two legs) ───────────────────────────────────────────
    private static FlatPatternResult DevelopLAngle(PartSpec spec)
    {
        double t = spec.Thickness, ri = spec.InsideRadius, k = spec.KFactor, angle = spec.AngleDeg;

        // legA = FlangeLeft, legB = FlangeRight. Each leg runs to a free edge (+1T inside).
        double legAo = spec.FlangeLeft.Basis == DimBasis.Inside ? spec.FlangeLeft.Value + t : spec.FlangeLeft.Value;
        double legBo = spec.FlangeRight.Basis == DimBasis.Inside ? spec.FlangeRight.Value + t : spec.FlangeRight.Value;

        double ossb = BendMath.Ossb(ri, t, angle);
        double ba = BendMath.BendAllowance(ri, t, k, angle);
        double bd = BendMath.BendDeduction(ri, t, k, angle, spec.MeasuredBendDeduction);

        double flatWidth = legAo + legBo - bd;
        double flatHeight = spec.Length;
        double bend1X = (legAo - ossb) + ba / 2.0;
        var bends = new[] { bend1X };

        // Centreline: (0,legBo) down to (0,0), bend, across to (legAo,0).
        var profile = BuildOffsetProfile(new[] { legBo, legAo }, new[] { +1 }, t, ri, 0, legBo, 0, -1);

        string slug = $"angle_{Trim(legAo)}x{Trim(legBo)}x{Trim(flatHeight)}";
        // WebOutside unused for L; legA in FlangeLeftOutside, legB in FlangeRightOutside.
        return Build(spec, ossb, ba, bd, 0, legAo, legBo, flatWidth, flatHeight, bends, profile, slug);
    }

    // ── Shared result assembly (cut geometry rectangle + N bend lines) ───────
    private static FlatPatternResult Build(
        PartSpec spec, double ossb, double ba, double bd,
        double webO, double flLO, double flRO,
        double flatWidth, double flatHeight, double[] bends,
        List<(double x, double y)> profile, string slug)
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
            Layers = { new CutLayer { Name = CutLayer, Color = 1 }, new CutLayer { Name = BendLayer, Color = 5 } },
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
            Summary = BuildSummary(spec, webO, flLO, flRO, bd, bends.Length),
            Title = PlainTitle(spec),
        };
    }

    // ── Cross-section profiles ───────────────────────────────────────────────

    private static double RoOf(double ri, double t, double a, double b, double c)
    {
        double maxR = Math.Min(Math.Min(b, c) * 0.9, Math.Max(a, 0.001) / 2.0 * 0.9);
        return Math.Min(ri + t, Math.Max(t, maxR));
    }

    /// <summary>U-channel material loop (web at bottom, flanges up) with radiused bend corners.</summary>
    private static List<(double x, double y)> BuildRadiusedU(double webO, double flL, double flR, double t, double ri)
    {
        double ro = RoOf(ri, t, webO, flL, flR);
        double riA = Math.Max(0.0, ro - t);
        const int seg = 8;
        var pts = new List<(double x, double y)>();
        void Arc(double cx, double cy, double r, double a0, double a1)
        {
            for (int i = 1; i <= seg; i++)
            {
                double a = (a0 + (a1 - a0) * i / seg) * Math.PI / 180.0;
                pts.Add((cx + r * Math.Cos(a), cy + r * Math.Sin(a)));
            }
        }
        pts.Add((0, flL)); pts.Add((0, ro));
        Arc(ro, ro, ro, 180, 270);
        pts.Add((webO - ro, 0));
        Arc(webO - ro, ro, ro, 270, 360);
        pts.Add((webO, flR)); pts.Add((webO - t, flR)); pts.Add((webO - t, t + riA));
        Arc(webO - t - riA, t + riA, riA, 0, -90);
        pts.Add((t + riA, t));
        Arc(t + riA, t + riA, riA, 270, 180);
        pts.Add((t, flL));
        return pts;
    }

    /// <summary>
    /// General radiused material loop for a segment chain. Walks the centreline (straight runs +
    /// 90° arcs of radius Ri+T/2 at each bend, turning +1 = left / -1 = right), then offsets ±T/2
    /// along the local normal to form the two faces. Handles same- and opposing-direction bends.
    /// </summary>
    private static List<(double x, double y)> BuildOffsetProfile(
        double[] segs, int[] turns, double t, double ri, double sx, double sy, double hx, double hy)
    {
        double rc = ri + t / 2.0;
        const int K = 6;
        var c = new List<(double x, double y, double nx, double ny)>();
        double px = sx, py = sy;
        double hn = Math.Sqrt(hx * hx + hy * hy); hx /= hn; hy /= hn;
        c.Add((px, py, -hy, hx));
        for (int i = 0; i < segs.Length; i++)
        {
            // The bend arc consumes ~rc of length at each adjacent bend; subtract it so segment
            // ends land on their nominal coordinates (so section dimensions touch the real edges).
            double cut = (i > 0 ? rc : 0) + (i < segs.Length - 1 ? rc : 0);
            double straight = Math.Max(0.001, segs[i] - cut);
            px += hx * straight; py += hy * straight;
            c.Add((px, py, -hy, hx));
            if (i < segs.Length - 1)
            {
                int d = turns[i];
                double cx = px + d * (-hy) * rc, cy = py + d * hx * rc;
                double a0 = Math.Atan2(py - cy, px - cx);
                for (int kk = 1; kk <= K; kk++)
                {
                    double a = a0 + d * (Math.PI / 2.0) * (kk / (double)K);
                    px = cx + rc * Math.Cos(a); py = cy + rc * Math.Sin(a);
                    hx = -d * Math.Sin(a); hy = d * Math.Cos(a);
                    c.Add((px, py, -hy, hx));
                }
            }
        }
        var loop = new List<(double x, double y)>();
        foreach (var p in c) loop.Add((p.x + p.nx * t / 2, p.y + p.ny * t / 2));
        for (int i = c.Count - 1; i >= 0; i--) loop.Add((c[i].x - c[i].nx * t / 2, c[i].y - c[i].ny * t / 2));
        return loop;
    }

    // ── Title / summary ──────────────────────────────────────────────────────
    private static string PlainTitle(PartSpec s)
    {
        if (s.Type is PartType.FlitchPlate or PartType.BasePlate)
        {
            string plate = s.Type == PartType.FlitchPlate ? "Flitch Plate" : "Base Plate";
            // Product-standard order: [metal] [drawing of] [thickness] x [width] x [length].
            return $"{MaterialPlain(s.Material)} {plate} {N(s.Thickness)}\" x {N(s.Width)}\" x {N(s.Length)}\"".Trim();
        }

        if (s.Type == PartType.Pan)
            return $"{MaterialPlain(s.Material)} Pan {N(s.Length)}\" x {N(s.Width)}\" x {N(s.Depth)}\" deep x {N(s.Thickness)}\"".Trim();

        if (s.Type == PartType.PaddleBlind)
            return $"{MaterialPlain(s.Material)} Paddle Blind {s.PaddleNps}\" #{s.PaddleClass} x {N(s.Thickness)}\"".Trim();

        string shape = s.Type switch
        {
            PartType.UChannel => "U Channel",
            PartType.LAngle   => "Angle",
            PartType.ZChannel => "Z Channel",
            _ => s.Type.ToString(),
        };
        string profileDims = s.Type == PartType.LAngle
            ? $"{N(s.FlangeLeft.Value)}\" x {N(s.FlangeRight.Value)}\""
            : $"{N(s.Web.Value)}\" x {N(s.FlangeLeft.Value)}\"";
        return $"{MaterialPlain(s.Material)} {shape} {profileDims} x {N(s.Thickness)}\" x {N(s.Length)}\"".Trim();
    }

    private static string MaterialPlain(string m) => m switch
    {
        "alum"   => "Aluminum",
        "CRS"    => "Cold Rolled Steel",
        "HRS"    => "Hot Rolled Steel",
        "galv"   => "Galvanized Steel",
        "HRPO"   => "HRPO Steel",
        "SS"     => "Stainless Steel",
        "Brass"  => "Brass",
        "Copper" => "Copper",
        _ => m,
    };

    private static string N(double v) => v.ToString("0.###", CultureInfo.InvariantCulture);

    private static string BuildSummary(PartSpec s, double webO, double flLO, double flRO, double bd, int nBends)
    {
        string basis = (s.Web.Basis == s.FlangeLeft.Basis && s.FlangeLeft.Basis == s.FlangeRight.Basis)
            ? s.Web.Basis.ToString().ToLowerInvariant() : "mixed";
        string u = s.Units;
        string shape = s.Type switch
        {
            PartType.UChannel => "U-channel", PartType.LAngle => "L-angle",
            PartType.ZChannel => "Z-channel", _ => s.Type.ToString(),
        };
        string outside = s.Type == PartType.LAngle
            ? $"legs {F(flLO)}{u} x {F(flRO)}{u}"
            : $"web {F(webO)}{u}, flange {(Math.Abs(flLO - flRO) < 1e-6 ? F(flLO) : $"{F(flLO)}/{F(flRO)}")}{u}";
        var lines = new List<string>
        {
            $"{shape}  {s.Material}  (T={F(s.Thickness)}{u})",
            $"Basis: {basis}  ->  outside {outside}",
            $"Bend: Ri {F(s.InsideRadius)}{u}, K {s.KFactor:0.##}, {s.AngleDeg:0.#} deg,  BD {F(bd)}{u}/bend (x{nBends})",
        };
        string? finishNote = s.Finish switch
        {
            FinishSide.Inside  => "Finish on inside face",
            FinishSide.Outside => "Finish on outside face",
            FinishSide.Top     => "Finish on top face",
            FinishSide.Bottom  => "Finish on bottom face",
            _ => null,
        };
        if (finishNote != null) lines.Add(finishNote);
        return string.Join("\n", lines);
    }

    private static string F(double v) => v.ToString("0.###", CultureInfo.InvariantCulture);
    private static string Trim(double v) => v.ToString("0.##", CultureInfo.InvariantCulture);
}

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
    public double ColumnOuterWidth { get; init; }
    public double ColumnOuterDepth { get; init; }
    public double ColumnWall { get; init; }
    public string ColumnShape { get; init; } = "square";
    public string ColumnLabel { get; init; } = "";

    public required string Summary { get; init; }
    public string Title { get; init; } = "";
}
