using System.Globalization;
using System.Text.RegularExpressions;
using OutlookShredder.Proxy.Models.Drawing;

namespace OutlookShredder.Proxy.Services.Drawing;

/// <summary>Thrown when a drawing description cannot be parsed; message is user-facing.</summary>
public sealed class DrawingParseException : Exception
{
    public DrawingParseException(string message) : base(message) { }
}

/// <summary>
/// Deterministic parser: plain-text part description → <see cref="PartSpec"/>. No AI.
///
/// Grammar (case-insensitive, commas optional, inches assumed):
///   TYPE  shorthand "U w x f x len"  OR  keyworded "web W flange F length L"
///   material: "18ga galv" | "16ga CRS" | "0.060 alum" (alum/ss need a decimal thickness)
///   basis:   global word "inside" / "outside" (default outside);
///            per-dimension "id" / "od" right after a value (e.g. "web 4 id")
///   options: "Ri 1/16", "K 0.42", "90deg"
///   numbers: integer, decimal, fraction "3/8", or mixed "1-3/8"
/// </summary>
public static class DrawingTextParser
{
    // A number: mixed fraction, simple fraction, decimal, or integer.
    private const string Num = @"\d+-\d+/\d+|\d+/\d+|\d*\.\d+|\d+";

    public static PartSpec Parse(string input)
    {
        if (string.IsNullOrWhiteSpace(input))
            throw new DrawingParseException("Enter a part description, e.g. \"U 4 x 2 x 36, 16ga CRS\".");

        string text = input.Trim();
        string lower = text.ToLowerInvariant();

        var type = DetectType(lower);

        // ── Material + thickness ────────────────────────────────────────────
        var (family, materialLabel) = DetectMaterial(lower);

        // ── Structural column — its own field set (per-plate thickness + tube); parsed + returned here ──
        if (type == PartType.Column)
            return ParseColumn(text, lower, materialLabel);

        double thickness = ResolveThickness(lower, family, materialLabel);

        // ── Bend params ─────────────────────────────────────────────────────
        double ri = MatchNum(lower, $@"\b(?:ri|radius|rad)\s*[:=]?\s*({Num})") ?? thickness;
        double k = MatchNum(lower, $@"\bk\s*[:=]?\s*({Num})") ?? BendMath.DefaultK(family);
        double angle = MatchNum(lower, $@"(?:\bangle\s*[:=]?\s*)?({Num})\s*(?:deg|degrees|°)") ?? 90.0;
        double? measuredBd = MatchNum(lower, $@"\b(?:bd|deduction)\s*[:=]?\s*({Num})");

        // ── Finish side — strip it so the 'inside'/'outside' words here don't trip the global basis ──
        FinishSide finish = FinishSide.None;
        var fm = Regex.Match(lower, @"\bfinish\s+(outside|inside|top|bottom)\b");
        if (fm.Success)
        {
            finish = fm.Groups[1].Value switch
            {
                "outside" => FinishSide.Outside,
                "inside"  => FinishSide.Inside,
                "top"     => FinishSide.Top,
                "bottom"  => FinishSide.Bottom,
                _         => FinishSide.None,
            };
            lower = lower.Remove(fm.Index, fm.Length);
        }

        // ── Flat plates (Flitch / Base) — different field set; parsed + returned here ──
        if (type is PartType.FlitchPlate or PartType.BasePlate)
            return ParsePlate(type, lower, thickness, materialLabel, ri, k);

        // ── Pan — L x W x D plus which walls are present; parsed + returned here ──
        if (type == PartType.Pan)
            return ParsePan(lower, thickness, materialLabel, ri, k, finish);

        // ── Paddle blind / spade ("frying pan") — NPS + class lookup; parsed + returned here ──
        if (type == PartType.PaddleBlind)
            return ParsePaddleBlind(lower, thickness, materialLabel, ri);

        // ── Global basis (whole words only; id/od are reserved for per-dim) ──
        DimBasis globalBasis = Regex.IsMatch(lower, @"\binside\b") ? DimBasis.Inside : DimBasis.Outside;

        // ── Dimensions (per part type) ──────────────────────────────────────
        Dim web = new(0, globalBasis), flangeL, flangeR;
        double length;
        // Optional inline per-bend angle + direction + return captured from the flange tokens.
        double? bendAngleL = null, bendAngleR = null;
        string? bendDirL = null, bendDirR = null;
        ReturnSpec? retL = null, retR = null;
        double? lenKw() => MatchNum(lower, $@"\b(?:length|len|run|long)\s*[:=]?\s*({Num})");

        if (type == PartType.LAngle)
        {
            // L-angle: two legs + length.  Shorthand "L 2 x 3 x 36"; keyworded "legs 2/3, length 36".
            var sh = Regex.Match(lower, $@"\bl(?:-?\s?angle)?\b\s*({Num})\s*x\s*({Num})\s*x\s*({Num})");
            if (sh.Success)
            {
                flangeL = new Dim(NumOf(sh.Groups[1].Value), globalBasis);
                flangeR = new Dim(NumOf(sh.Groups[2].Value), globalBasis);
                length  = NumOf(sh.Groups[3].Value);
            }
            else
            {
                var fx = MatchFlangesEx(lower, globalBasis)
                    ?? throw new DrawingParseException("Missing leg dimensions. Try \"L 2 x 3 x 36\" or \"legs 2/3\".");
                flangeL = fx.Left.Dim; flangeR = fx.Right.Dim;
                bendAngleL = fx.Left.AngleDeg; bendDirL = fx.Left.Dir; retL = fx.Left.Return;
                bendAngleR = fx.Right.AngleDeg; bendDirR = fx.Right.Dir; retR = fx.Right.Return;
                length  = lenKw() ?? throw new DrawingParseException("Missing length. Try \"length 36\".");
            }
        }
        else if (type == PartType.ZChannel)
        {
            // Z-channel: flange x web x flange + length.  Shorthand "Z 2 x 4 x 2 x 36".
            var sh = Regex.Match(lower, $@"\bz(?:-?\s?channel)?\b\s*({Num})\s*x\s*({Num})\s*x\s*({Num})\s*x\s*({Num})");
            if (sh.Success)
            {
                flangeL = new Dim(NumOf(sh.Groups[1].Value), globalBasis);
                web     = new Dim(NumOf(sh.Groups[2].Value), globalBasis);
                flangeR = new Dim(NumOf(sh.Groups[3].Value), globalBasis);
                length  = NumOf(sh.Groups[4].Value);
            }
            else
            {
                web = MatchDim(lower, @"web|base|bottom", globalBasis)
                    ?? throw new DrawingParseException("Missing web dimension. Try \"web 4\" or \"Z 2 x 4 x 2 x 36\".");
                var fx = MatchFlangesEx(lower, globalBasis)
                    ?? throw new DrawingParseException("Missing flange dimension. Try \"flange 2\".");
                flangeL = fx.Left.Dim; flangeR = fx.Right.Dim;
                bendAngleL = fx.Left.AngleDeg; bendDirL = fx.Left.Dir; retL = fx.Left.Return;
                bendAngleR = fx.Right.AngleDeg; bendDirR = fx.Right.Dir; retR = fx.Right.Return;
                length  = lenKw() ?? throw new DrawingParseException("Missing length. Try \"length 36\".");
            }
        }
        else
        {
            // U-channel: web x flange + length.  Shorthand "U 4 x 2 x 36".
            var sh = Regex.Match(lower, $@"\bu(?:-?\s?channel)?\b\s*({Num})\s*x\s*({Num})\s*x\s*({Num})");
            if (sh.Success)
            {
                web     = new Dim(NumOf(sh.Groups[1].Value), globalBasis);
                flangeL = new Dim(NumOf(sh.Groups[2].Value), globalBasis);
                flangeR = flangeL;
                length  = NumOf(sh.Groups[3].Value);
            }
            else
            {
                web = MatchDim(lower, @"web|base|bottom", globalBasis)
                    ?? throw new DrawingParseException("Missing web dimension. Try \"web 4\" or shorthand \"U 4 x 2 x 36\".");
                var fx = MatchFlangesEx(lower, globalBasis)
                    ?? throw new DrawingParseException("Missing flange dimension. Try \"flange 2\" (or \"flange 2/1.5\" for unequal).");
                flangeL = fx.Left.Dim; flangeR = fx.Right.Dim;
                bendAngleL = fx.Left.AngleDeg; bendDirL = fx.Left.Dir; retL = fx.Left.Return;
                bendAngleR = fx.Right.AngleDeg; bendDirR = fx.Right.Dir; retR = fx.Right.Return;
                length  = lenKw() ?? throw new DrawingParseException("Missing length. Try \"length 36\".");
            }
        }

        bool needWeb = type != PartType.LAngle;
        if ((needWeb && web.Value <= 0) || flangeL.Value <= 0 || flangeR.Value <= 0 || length <= 0)
            throw new DrawingParseException("Dimensions must be positive numbers.");
        if (thickness <= 0)
            throw new DrawingParseException(
                "Could not determine material thickness. Add a gauge (e.g. \"16ga CRS\") or a decimal thickness (e.g. \"0.060 alum\").");

        // Per-bend angle + direction (only when the flange tokens carried explicit angle/direction).
        // Otherwise leave Bends null so FlatPattern keeps its exact default geometry (90° + shape dirs).
        bool annotated = bendAngleL.HasValue || bendDirL != null || bendAngleR.HasValue || bendDirR != null;
        IReadOnlyList<BendSpec>? bends = null;
        if (annotated)
        {
            var b0 = new BendSpec(bendAngleL ?? angle, DirOf(bendDirL, BendDir.Up));
            if (type == PartType.LAngle)
                bends = new[] { b0 };
            else
            {
                var defR = type == PartType.ZChannel ? BendDir.Down : BendDir.Up;
                bends = new[] { b0, new BendSpec(bendAngleR ?? angle, DirOf(bendDirR, defR)) };
            }
        }

        return new PartSpec
        {
            Type = type,
            Web = web,
            FlangeLeft = flangeL,
            FlangeRight = flangeR,
            Length = length,
            Thickness = thickness,
            InsideRadius = ri,
            KFactor = k,
            AngleDeg = angle,
            MeasuredBendDeduction = measuredBd,
            Bends = bends,
            AnglesAnnotated = annotated,
            ReturnLeft = retL,
            ReturnRight = retR,
            Material = materialLabel,
            Units = "in",
            Finish = finish,
        };
    }

    private static PartType DetectType(string lower)
    {
        // Column first — its clause contains "base"/"bearing" which must not trip the plate checks.
        if (Regex.IsMatch(lower, @"\bcolumn\b")) return PartType.Column;
        if (Regex.IsMatch(lower, @"\bu(?:-?\s?channel)?\b")) return PartType.UChannel;
        if (Regex.IsMatch(lower, @"\bz(?:-?\s?channel)?\b")) return PartType.ZChannel;
        if (Regex.IsMatch(lower, @"\bl(?:-?\s?angle)?\b"))   return PartType.LAngle;
        if (Regex.IsMatch(lower, @"\bflitch\b"))      return PartType.FlitchPlate;
        if (Regex.IsMatch(lower, @"\bbase\s*plate\b")) return PartType.BasePlate;
        // Paddle blind ("frying pan") must come before the plain "pan" check.
        if (Regex.IsMatch(lower, @"\b(frying\s*pan|paddle\s*(?:blind|blank)|spade)\b")) return PartType.PaddleBlind;
        if (Regex.IsMatch(lower, @"\b(pan|tray)\b")) return PartType.Pan;
        if (Regex.IsMatch(lower, @"\bhat\b"))        throw NotYet("Hat");
        throw new DrawingParseException(
            "Could not identify the part type. Start with U (U-channel), L (angle) or Z (Z-channel).");
    }

    private static DrawingParseException NotYet(string what) =>
        new($"{what} is coming soon — U-channel, L-angle and Z-channel are available now.");

    /// <summary>Parses a flat plate (Flitch / Base) — length x width, thickness, and bolt holes.</summary>
    private static PartSpec ParsePlate(PartType type, string lower, double thickness, string material, double ri, double k)
    {
        if (thickness <= 0)
            throw new DrawingParseException("Add a gauge or decimal thickness (e.g. \"0.25 steel\").");

        double L, W;
        HoleSpec? holes = null;

        if (type == PartType.FlitchPlate)
        {
            var m = Regex.Match(lower, $@"\bflitch\b\s*({Num})\s*x\s*({Num})");
            if (!m.Success) throw new DrawingParseException("Flitch needs length x width, e.g. \"Flitch 48 x 6, 0.25 steel\".");
            L = NumOf(m.Groups[1].Value); W = NumOf(m.Groups[2].Value);

            var hm = Regex.Match(lower, $@"holes?\s+({Num})\s*(staggered|paired)?\s*(?:@|at|spacing|spaced)?\s*({Num})?");
            if (hm.Success)
            {
                // End offsets: "end 2" sets both; "lhs 2 rhs 4" overrides each independently
                // (a "/" separator can't be used here — "2/4" parses as the fraction 0.5).
                double leftEnd = 0, rightEnd = 0;
                var endm = Regex.Match(lower, $@"\bend\s*[:=]?\s*({Num})");
                if (endm.Success) leftEnd = rightEnd = NumOf(endm.Groups[1].Value);
                var lm = Regex.Match(lower, $@"\blhs\s*[:=]?\s*({Num})");
                if (lm.Success) leftEnd = NumOf(lm.Groups[1].Value);
                var rm = Regex.Match(lower, $@"\brhs\s*[:=]?\s*({Num})");
                if (rm.Success) rightEnd = NumOf(rm.Groups[1].Value);
                double topEdge = 0, bottomEdge = 0;
                var em = Regex.Match(lower, $@"\bedge\s*[:=]?\s*({Num})(?:\s*/\s*({Num}))?");
                if (em.Success)
                {
                    topEdge = NumOf(em.Groups[1].Value);
                    bottomEdge = em.Groups[2].Success && em.Groups[2].Value.Length > 0 ? NumOf(em.Groups[2].Value) : topEdge;
                }
                holes = new HoleSpec
                {
                    Diameter = NumOf(hm.Groups[1].Value),
                    Pattern = hm.Groups[2].Value == "paired" ? HolePattern.Paired : HolePattern.Staggered,
                    Spacing = hm.Groups[3].Success && hm.Groups[3].Value.Length > 0 ? NumOf(hm.Groups[3].Value) : 16,
                    LeftEndOffset = leftEnd,
                    RightEndOffset = rightEnd,
                    TopEdge = topEdge,
                    BottomEdge = bottomEdge,
                };
            }
        }
        else // BasePlate
        {
            var m = Regex.Match(lower, $@"\bbase\s*plate\b\s*({Num})\s*x\s*({Num})");
            if (!m.Success) throw new DrawingParseException("Base plate needs length x width, e.g. \"Base Plate 8 x 8, 0.5 steel\".");
            L = NumOf(m.Groups[1].Value); W = NumOf(m.Groups[2].Value);

            var hm = Regex.Match(lower, $@"({Num})\s*holes?\s+({Num})\s*(?:edge|ed)?\s*({Num})?");
            if (hm.Success)
                holes = new HoleSpec
                {
                    Pattern = HolePattern.Corner,
                    Count = (int)Math.Round(NumOf(hm.Groups[1].Value)),
                    Diameter = NumOf(hm.Groups[2].Value),
                    EdgeDistance = hm.Groups[3].Success && hm.Groups[3].Value.Length > 0 ? NumOf(hm.Groups[3].Value) : 1.0,
                };
        }

        if (L <= 0 || W <= 0) throw new DrawingParseException("Plate dimensions must be positive.");

        return new PartSpec
        {
            Type = type, Length = L, Width = W, Thickness = thickness,
            InsideRadius = ri, KFactor = k, Material = material, Holes = holes, Units = "in",
        };
    }

    /// <summary>Parses a pan — "Pan L x W x D, {n} long {n} short, {material}".</summary>
    private static PartSpec ParsePan(string lower, double thickness, string material, double ri, double k, FinishSide finish)
    {
        if (thickness <= 0)
            throw new DrawingParseException("Add a gauge or decimal thickness (e.g. \"16ga\").");

        // Each of L/W/D carries an optional id/od basis (default outside).
        var m = Regex.Match(lower,
            $@"\b(?:pan|tray)\b\s*({Num})\s*(id|od)?\s*x\s*({Num})\s*(id|od)?\s*x\s*({Num})\s*(id|od)?");
        if (!m.Success)
            throw new DrawingParseException("Pan needs length x width x depth, e.g. \"Pan 24 x 18 x 2, 2 long 2 short, 16ga\".");
        double L = NumOf(m.Groups[1].Value), W = NumOf(m.Groups[3].Value), D = NumOf(m.Groups[5].Value);
        if (L <= 0 || W <= 0 || D <= 0) throw new DrawingParseException("Pan dimensions must be positive.");

        int longN  = (int)Math.Round(MatchNum(lower, $@"({Num})\s*long")  ?? 2);
        int shortN = (int)Math.Round(MatchNum(lower, $@"({Num})\s*short") ?? 2);

        // Optional shared return (lip/hem) applied to every present wall: "returns 0.5 [id|od] @ 90|180 [up|down]".
        ReturnSpec? panReturn = null;
        var rm = Regex.Match(lower, $@"\breturns?\s+({Num})\s*(id|od)?\s*(?:@?\s*(90|180)\s*(?:°|deg|degrees)?)?\s*(up|down)?");
        if (rm.Success && NumOf(rm.Groups[1].Value) > 0)
            panReturn = new ReturnSpec(
                NumOf(rm.Groups[1].Value),
                BasisFrom(rm.Groups[2].Value, DimBasis.Outside),
                rm.Groups[3].Success && rm.Groups[3].Value.Length > 0 ? NumOf(rm.Groups[3].Value) : 90.0,
                DirOf(rm.Groups[4].Success && rm.Groups[4].Value.Length > 0 ? rm.Groups[4].Value : null, BendDir.Up));

        return new PartSpec
        {
            Type = PartType.Pan,
            Length = L, Width = W, Depth = D,
            LengthBasis = BasisFrom(m.Groups[2].Value, DimBasis.Outside),
            WidthBasis = BasisFrom(m.Groups[4].Value, DimBasis.Outside),
            DepthBasis = BasisFrom(m.Groups[6].Value, DimBasis.Outside),
            Thickness = thickness, InsideRadius = ri, KFactor = k, AngleDeg = 90,
            Material = material, Units = "in", Finish = finish,
            PanBottom = longN >= 1, PanTop = longN >= 2,
            PanLeft = shortN >= 1, PanRight = shortN >= 2,
            PanReturn = panReturn,
        };
    }

    /// <summary>Parses a paddle blind / spade — "Frying Pan {nps} #{class}, {material/thickness}".</summary>
    private static PartSpec ParsePaddleBlind(string lower, double thickness, string material, double ri)
    {
        // NPS: the first number after the type keyword (fraction / mixed / decimal), optional inch mark.
        var nm = Regex.Match(lower, $@"\b(?:frying\s*pan|paddle\s*(?:blind|blank)|spade)\b\s*({Num})\s*(?:""|in|inch)?");
        if (!nm.Success)
            throw new DrawingParseException("Frying pan needs a pipe size, e.g. \"Frying Pan 6 #150, SS\".");
        double npsValue = NumOf(nm.Groups[1].Value);

        // Pressure class: "#150" / "150" / "class 300".
        int cls = (int)Math.Round(MatchNum(lower, @"(?:#|class\s*)?\b(150|300)\b") ?? 0);
        if (cls != 150 && cls != 300)
            throw new DrawingParseException("Pressure class must be #150 or #300, e.g. \"Frying Pan 6 #150\".");

        var pb = PaddleBlankTable.Find(npsValue, cls)
            ?? throw new DrawingParseException($"No ASME B16.48 spade for NPS {nm.Groups[1].Value}\" Class {cls}. Sizes are 1/2\" to 20\".");

        // Thickness: the user's pick wins; fall back to the standard minimum if none was given.
        double t = thickness > 0 ? thickness : pb.Thickness;

        return new PartSpec
        {
            Type = PartType.PaddleBlind,
            PaddleOd = pb.Od,
            PaddleHandleWidth = pb.HandleWidth,
            PaddleCenterToEnd = pb.CenterToEnd,
            PaddleNps = pb.Nps,
            PaddleClass = cls,
            Thickness = t,
            InsideRadius = ri > 0 ? ri : t,
            Material = material,
            Units = "in",
        };
    }

    /// <summary>
    /// Parses a structural column —
    ///   "Column height H, base L x W x T [holes N x D edge E], bearing L x W x T [holes …],
    ///    column {round|square|rect} D1 [x D2] wall W, {material} "Product label"".
    /// Tube cut length = H − base T − bearing T (computed in FlatPattern).
    /// </summary>
    private static PartSpec ParseColumn(string text, string lower, string material)
    {
        double height = MatchNum(lower, $@"\b(?:height|h|gap)\s*[:=]?\s*({Num})")
            ?? throw new DrawingParseException("Column needs a full height, e.g. \"Column height 96, base 10 x 10 x 0.75, …\".");

        var bm = Regex.Match(lower, $@"\bbase\b\s*({Num})\s*x\s*({Num})\s*x\s*({Num})");
        if (!bm.Success)
            throw new DrawingParseException("Column needs a base plate, e.g. \"base 10 x 10 x 0.75\".");
        double baseL = NumOf(bm.Groups[1].Value), baseW = NumOf(bm.Groups[2].Value), baseT = NumOf(bm.Groups[3].Value);

        var rm = Regex.Match(lower, $@"\bbearing\b\s*({Num})\s*x\s*({Num})\s*x\s*({Num})");
        if (!rm.Success)
            throw new DrawingParseException("Column needs a bearing plate, e.g. \"bearing 8 x 8 x 0.5\".");
        double bearL = NumOf(rm.Groups[1].Value), bearW = NumOf(rm.Groups[2].Value), bearT = NumOf(rm.Groups[3].Value);

        HoleSpec? baseHoles = ParseColumnPlateHoles(lower, "base");
        HoleSpec? bearHoles = ParseColumnPlateHoles(lower, "bearing");

        var cm = Regex.Match(lower,
            $@"\bcolumn\s+(round|square|rect(?:angle|angular)?)\s+({Num})(?:\s*x\s*({Num}))?\s+wall\s+({Num})");
        if (!cm.Success)
            throw new DrawingParseException("Column needs a tube/pipe, e.g. \"column square 4 x 4 wall 0.25\" or \"column round 6.625 wall 0.28\".");
        string shape = cm.Groups[1].Value.StartsWith("rect") ? "rect" : cm.Groups[1].Value;
        double d1 = NumOf(cm.Groups[2].Value);
        double d2 = cm.Groups[3].Success && cm.Groups[3].Value.Length > 0 ? NumOf(cm.Groups[3].Value) : d1;
        double wall = NumOf(cm.Groups[4].Value);

        // Product label: trailing quoted token, taken from the original-case text.
        string label = "";
        var lm = Regex.Match(text, "\"([^\"]+)\"");
        if (lm.Success) label = lm.Groups[1].Value.Trim();

        if (height <= 0 || baseL <= 0 || baseW <= 0 || baseT <= 0 || bearL <= 0 || bearW <= 0 || bearT <= 0 || d1 <= 0)
            throw new DrawingParseException("Column dimensions must be positive numbers.");
        if (height - baseT - bearT <= 0)
            throw new DrawingParseException(
                $"Full height {F0(height)}\" is too small for the plates (base {F0(baseT)}\" + bearing {F0(bearT)}\").");

        return new PartSpec
        {
            Type = PartType.Column,
            ColumnFullHeight = height,
            BaseLength = baseL, BaseWidth = baseW, BaseThickness = baseT, BaseHoles = baseHoles,
            BearingLength = bearL, BearingWidth = bearW, BearingThickness = bearT, BearingHoles = bearHoles,
            ColumnShape = shape, ColumnOuterWidth = d1, ColumnOuterDepth = d2, ColumnWall = wall,
            ColumnLabel = label,
            Thickness = baseT,   // representative only; geometry uses the per-plate thicknesses
            Material = material, Units = "in",
        };
    }

    /// <summary>Corner-hole spec for one column plate, scoped to that plate's clause in the text.</summary>
    private static HoleSpec? ParseColumnPlateHoles(string lower, string keyword)
    {
        // Isolate the clause from this plate keyword up to the next section keyword.
        var clauseM = Regex.Match(lower, $@"\b{keyword}\b(.*?)(?=\bbearing\b|\bcolumn\b|$)");
        if (!clauseM.Success) return null;
        string clause = clauseM.Groups[1].Value;
        var hm = Regex.Match(clause, $@"holes?\s+({Num})\s*(?:x\s*)?({Num})\s*(?:edge|ed)?\s*({Num})?");
        if (!hm.Success) return null;
        return new HoleSpec
        {
            Pattern = HolePattern.Corner,
            Count = (int)Math.Round(NumOf(hm.Groups[1].Value)),
            Diameter = NumOf(hm.Groups[2].Value),
            EdgeDistance = hm.Groups[3].Success && hm.Groups[3].Value.Length > 0 ? NumOf(hm.Groups[3].Value) : 1.0,
        };
    }

    private static string F0(double v) => v.ToString("0.###", CultureInfo.InvariantCulture);

    private static (MaterialFamily, string) DetectMaterial(string lower)
    {
        if (Regex.IsMatch(lower, @"\bgalv(?:anized|anised)?\b")) return (MaterialFamily.Galvanized, "galv");
        if (Regex.IsMatch(lower, @"\bhrpo\b") || Regex.IsMatch(lower, @"\bhot[\s-]?rolled\b"))
            return (MaterialFamily.HotRolled, "HRPO");
        if (Regex.IsMatch(lower, @"\b(crs|cr|cold[\s-]?rolled)\b")) return (MaterialFamily.ColdRolled, "CRS");
        if (Regex.IsMatch(lower, @"\b(hrs|hr)\b")) return (MaterialFamily.HotRolled, "HRS");
        if (Regex.IsMatch(lower, @"\b(alum(?:in(?:i?um)?)?|al)\b")) return (MaterialFamily.Aluminium, "alum");
        if (Regex.IsMatch(lower, @"\b(stainless|ss)\b")) return (MaterialFamily.Stainless, "SS");
        if (Regex.IsMatch(lower, @"\bbrass\b")) return (MaterialFamily.Brass, "Brass");
        if (Regex.IsMatch(lower, @"\b(copper|cu)\b")) return (MaterialFamily.Copper, "Copper");
        if (Regex.IsMatch(lower, @"\bsteel\b")) return (MaterialFamily.ColdRolled, "CRS");
        return (MaterialFamily.Unknown, "");
    }

    private static double ResolveThickness(string lower, MaterialFamily family, string materialLabel)
    {
        // Explicit thickness keyword always wins.
        var explicitT = MatchNum(lower, $@"\b(?:t|thk|thickness)\s*[:=]?\s*({Num})");
        if (explicitT is > 0) return explicitT.Value;

        // Gauge table.
        var gauge = MatchInt(lower, @"\b(\d{1,2})\s*ga(?:uge)?\b");
        if (gauge is not null)
        {
            var fromGauge = GaugeTables.Thickness(family == MaterialFamily.Unknown ? MaterialFamily.ColdRolled : family, gauge.Value);
            if (fromGauge is > 0) return fromGauge.Value;
        }

        // Decimal thickness for non-gauge picks. The wizard emits the value immediately before the
        // material word ("0.125 CRS", "0.060 alum"); free NL reads the same. ANCHOR the match to the
        // material token (or an explicit inch mark) so a decimal DIMENSION (flange 1.5, web 3.25, length
        // 0.5) is never grabbed as the thickness — that mis-parse was garbling U/Z/L whenever a
        // custom/decimal thickness met a decimal dimension (gauges were unaffected, matched above).
        const string matTok = @"(?:galv\w*|hrpo|hot[\s-]?rolled|hrs|hr|crs|cr|cold[\s-]?rolled|alum\w*|al|stainless|ss|brass|copper|cu|steel)";
        var dec = MatchNum(lower, $@"\b(\d*\.\d+)\s*(?:""|in|inch)?\s*{matTok}\b");
        if (dec is > 0) return dec.Value;

        // No material word — accept a decimal that explicitly carries an inch mark ("0.25in", "0.25\"").
        var decInch = MatchNum(lower, @"\b(\d*\.\d+)\s*(?:""|in|inch)\b");
        if (decInch is > 0) return decInch.Value;

        return 0; // signalled as an error upstream
    }

    /// <summary>Matches "KEYWORD value [id|od]" and applies a per-dimension basis override.</summary>
    private static Dim? MatchDim(string lower, string keywordAlternation, DimBasis fallback)
    {
        var m = Regex.Match(lower, $@"\b(?:{keywordAlternation})\s*[:=]?\s*({Num})\s*\(?\s*(id|od)?\s*\)?");
        if (!m.Success) return null;
        var basis = BasisFrom(m.Groups[2].Value, fallback);
        return new Dim(NumOf(m.Groups[1].Value), basis);
    }

    /// <summary>
    /// Resolves the two flanges/legs, each with its own inside/outside basis. Accepts:
    ///   one value (equal), "F/F2" (unequal, shared basis), or two separate "flange …"
    /// occurrences (unequal, independent basis — e.g. "flange 2 id flange 1.5 od" = J-channel).
    /// </summary>
    private static (Dim, Dim)? MatchFlanges(string lower, DimBasis fallback)
    {
        var ms = Regex.Matches(lower, $@"\b(?:flanges?|legs?|walls?|sides?)\s*[:=]?\s*({Num})(?:\s*/\s*({Num}))?\s*\(?\s*(id|od)?\s*\)?");
        if (ms.Count == 0) return null;

        var m1 = ms[0];
        var b1 = BasisFrom(m1.Groups[3].Value, fallback);
        var left = new Dim(NumOf(m1.Groups[1].Value), b1);

        Dim right;
        if (ms.Count >= 2)
        {
            var m2 = ms[1];
            right = new Dim(NumOf(m2.Groups[1].Value), BasisFrom(m2.Groups[3].Value, fallback));
        }
        else if (m1.Groups[2].Success && m1.Groups[2].Value.Length > 0)
        {
            right = new Dim(NumOf(m1.Groups[2].Value), b1);   // "F/F2" shared basis
        }
        else
        {
            right = left;
        }
        return (left, right);
    }

    private static DimBasis BasisFrom(string token, DimBasis fallback) => token switch
    {
        "id" => DimBasis.Inside,
        "od" => DimBasis.Outside,
        _ => fallback,
    };

    /// <summary>A flange/leg dimension plus an optional inline per-bend angle + fold direction and
    /// an optional return (lip/hem) clause.</summary>
    private readonly record struct FlangeInfo(Dim Dim, double? AngleDeg, string? Dir, ReturnSpec? Return);

    /// <summary>
    /// Like <see cref="MatchFlanges"/> but also captures an optional inline "[@] {angle}[deg|°]
    /// {up|down}" bend suffix and an optional "return {len} [id|od] @ {90|180} {up|down}" clause
    /// on each flange (e.g. "flange 2 @ 75° up return 0.5 id @ 90 up").
    /// </summary>
    private static (FlangeInfo Left, FlangeInfo Right)? MatchFlangesEx(string lower, DimBasis fallback)
    {
        var rx = $@"\b(?:flanges?|legs?|walls?|sides?)\s*[:=]?\s*({Num})(?:\s*/\s*({Num}))?\s*\(?\s*(id|od)?\s*\)?" +
                 $@"(?:\s*@?\s*({Num})\s*(?:°|deg|degrees)?)?(?:\s*(up|down))?" +
                 $@"(?:\s*return\s+({Num})\s*(id|od)?\s*(?:@?\s*(90|180)\s*(?:°|deg|degrees)?)?\s*(up|down)?)?";
        var ms = Regex.Matches(lower, rx);
        if (ms.Count == 0) return null;

        static double? Ang(Match m) => m.Groups[4].Success && m.Groups[4].Value.Length > 0 ? NumOf(m.Groups[4].Value) : (double?)null;
        static string? Dir(Match m) => m.Groups[5].Success && m.Groups[5].Value.Length > 0 ? m.Groups[5].Value : null;
        static ReturnSpec? Ret(Match m)
        {
            if (!m.Groups[6].Success || m.Groups[6].Value.Length == 0) return null;
            double len = NumOf(m.Groups[6].Value);
            if (len <= 0) return null;
            var basis = BasisFrom(m.Groups[7].Value, DimBasis.Outside);
            double ang = m.Groups[8].Success && m.Groups[8].Value.Length > 0 ? NumOf(m.Groups[8].Value) : 90.0;
            var dir = DirOf(m.Groups[9].Success && m.Groups[9].Value.Length > 0 ? m.Groups[9].Value : null, BendDir.Up);
            return new ReturnSpec(len, basis, ang, dir);
        }

        var m1 = ms[0];
        var b1 = BasisFrom(m1.Groups[3].Value, fallback);
        var left = new FlangeInfo(new Dim(NumOf(m1.Groups[1].Value), b1), Ang(m1), Dir(m1), Ret(m1));

        FlangeInfo right;
        if (ms.Count >= 2)
        {
            var m2 = ms[1];
            right = new FlangeInfo(new Dim(NumOf(m2.Groups[1].Value), BasisFrom(m2.Groups[3].Value, fallback)), Ang(m2), Dir(m2), Ret(m2));
        }
        else if (m1.Groups[2].Success && m1.Groups[2].Value.Length > 0)
        {
            right = new FlangeInfo(new Dim(NumOf(m1.Groups[2].Value), b1), Ang(m1), Dir(m1), Ret(m1));   // "F/F2" shares basis + bend + return
        }
        else
        {
            right = new FlangeInfo(left.Dim, Ang(m1), Dir(m1), Ret(m1));   // one value ⇒ equal flanges, equal bend + return
        }
        return (left, right);
    }

    private static BendDir DirOf(string? s, BendDir fallback) => s switch
    {
        "up" => BendDir.Up,
        "down" => BendDir.Down,
        _ => fallback,
    };

    // ── number helpers ──────────────────────────────────────────────────────

    private static double? MatchNum(string text, string pattern)
    {
        var m = Regex.Match(text, pattern);
        return m.Success ? NumOf(m.Groups[1].Value) : null;
    }

    private static int? MatchInt(string text, string pattern)
    {
        var m = Regex.Match(text, pattern);
        return m.Success && int.TryParse(m.Groups[1].Value, out var v) ? v : null;
    }

    /// <summary>Parses integer, decimal, simple fraction "3/8", or mixed fraction "1-3/8".</summary>
    private static double NumOf(string s)
    {
        s = s.Trim();
        var mixed = Regex.Match(s, @"^(\d+)-(\d+)/(\d+)$");
        if (mixed.Success)
            return int.Parse(mixed.Groups[1].Value) +
                   double.Parse(mixed.Groups[2].Value) / double.Parse(mixed.Groups[3].Value);

        var frac = Regex.Match(s, @"^(\d+)/(\d+)$");
        if (frac.Success)
            return double.Parse(frac.Groups[1].Value) / double.Parse(frac.Groups[2].Value);

        return double.Parse(s, CultureInfo.InvariantCulture);
    }
}
