using System.Text.RegularExpressions;

namespace OutlookShredder.Proxy.Services;

/// <summary>
/// Computes theoretical weight per linear foot (lb/ft) or per square foot (lb/sqft)
/// from a catalog product name using standard metal cross-section geometry.
/// All formulas are based on AISC / metals-industry published section properties.
/// </summary>
public static class WeightCalculator
{
    public record WeightResult(
        double? LbPerFoot,
        double? LbPerSqFt,
        string? Formula,
        string? Note);

    // ── Densities (lb/in³) ────────────────────────────────────────────────────

    private static double Density(string metal) => metal switch
    {
        "aluminum"  => 0.0975,
        "stainless" => 0.2888,
        "copper"    => 0.3240,
        "brass"     => 0.3060,
        "bronze"    => 0.3180,
        _           => 0.2833   // carbon steel: HR, CR, galvanized, A36, A500, A513, etc.
    };

    private static string DetectMetal(string name)
    {
        if (CI(name, "Aluminum") || CI(name, "Aluminium")) return "aluminum";
        if (CI(name, "Stainless"))                         return "stainless";
        if (CI(name, "Copper"))                            return "copper";
        if (CI(name, "Brass"))                             return "brass";
        if (CI(name, "Bronze"))                            return "bronze";
        return "steel";
    }

    private static string DetectShape(string name)
    {
        if (CI(name, "Rectangular Tube") || CI(name, "Rect Tube") || CI(name, "Rect. Tube")) return "rect_tube";
        if (CI(name, "Square Tube"))          return "sq_tube";
        if (CI(name, "Round Tube"))           return "round_tube";
        if (CI(name, "Pipe"))                 return "pipe";
        if (CI(name, "Tread Plate") || CI(name, "Treadplate")) return "tread_plate";
        if (CI(name, "Plate") || CI(name, "Sheet"))            return "sheet";
        if (CI(name, "Hexagon Bar") || CI(name, "Hex Bar"))    return "hex_bar";
        if (CI(name, "Flat Bar"))             return "flat_bar";
        if (CI(name, "Round Bar"))            return "round_bar";
        if (CI(name, "Square Bar"))           return "sq_bar";
        if (CI(name, "Angle"))                return "angle";
        if (CI(name, "Channel"))              return "channel";
        if (CI(name, "Wide Flange") || CI(name, "W Beam") || CI(name, "Wide Beam")) return "beam";
        if (CI(name, "Standard Beam") || CI(name, "Std Beam"))  return "beam";
        if (CI(name, "H Pile") || CI(name, "H-Pile"))           return "beam";
        if (CI(name, "Beam"))                 return "beam";
        if (CI(name, "Tee Bar") || CI(name, "Structural Tee")) return "tee";
        return "unknown";
    }

    private static bool CI(string s, string sub) =>
        s.Contains(sub, StringComparison.OrdinalIgnoreCase);

    // ── Dimension extraction ──────────────────────────────────────────────────

    private static readonly Regex _dim3 = new(
        @"(\d+\.?\d*)\s*[Xx×]\s*(\d+\.?\d*)\s*[Xx×]\s*(\d+\.?\d*)",
        RegexOptions.Compiled);

    private static readonly Regex _dim2 = new(
        @"(\d+\.?\d*)\s*[Xx×]\s*(\d+\.?\d*)",
        RegexOptions.Compiled);

    // Decimal number (not followed by another X, not part of a grade like 304/6061/A500).
    // Looks for N.NNN pattern — requires decimal point to avoid false-positive on grade numbers.
    private static readonly Regex _dim1 = new(
        @"(?<![A-Za-z\d/])(\d+\.\d+)(?!\s*[Xx×])",
        RegexOptions.Compiled);

    // Pipe: "1.660 OD X 0.140 Wall" (value before keyword)
    private static readonly Regex _pipeOdWallFmt1 = new(
        @"(\d+\.?\d*)\s+OD\s+[Xx×]\s+(\d+\.?\d*)\s+Wall",
        RegexOptions.Compiled | RegexOptions.IgnoreCase);

    // Pipe: "OD 8.625 - Wall 0.188" (value after both keywords)
    private static readonly Regex _pipeOdWallFmt2 = new(
        @"OD\s+(\d+\.?\d*)\s*[-,]\s*Wall\s+(\d+\.?\d*)",
        RegexOptions.Compiled | RegexOptions.IgnoreCase);

    // Pipe: "OD 1.315 - 0.133 Wall" (OD after keyword, wall value before keyword)
    private static readonly Regex _pipeOdWallFmt3 = new(
        @"OD\s+(\d+\.?\d*)\s*[-,]\s*(\d+\.?\d*)\s+Wall",
        RegexOptions.Compiled | RegexOptions.IgnoreCase);

    // Beam/channel: "(H10 x W4.661 x FT 0.491 x WT 0.311)" — full I-beam spec
    private static readonly Regex _beamHWFtWt = new(
        @"H\s*(\d+\.?\d*)\s*[Xx×]\s*W\s*(\d+\.?\d*)\s*[Xx×]\s*FT\s+(\d+\.?\d*)\s*[Xx×]\s*WT\s+(\d+\.?\d*)",
        RegexOptions.Compiled | RegexOptions.IgnoreCase);

    // Channel: "(H3 x W1.5 x WT 0.13)" — no flange thickness given
    private static readonly Regex _channelHWWt = new(
        @"H\s*(\d+\.?\d*)\s*[Xx×]\s*W\s*(\d+\.?\d*)\s*[Xx×]\s*WT\s+(\d+\.?\d*)",
        RegexOptions.Compiled | RegexOptions.IgnoreCase);

    // Fractions with metal-standard (power-of-two) denominators only — avoids corrupting grades like
    // "304/304L". Mixed ("1 1/2"), simple ("1/8"), plus leading-dot decimals (".125") and inch/foot marks
    // and unit/finish words that sit BETWEEN dimension numbers and the × separators (which the dim regexes
    // require to be whitespace-only). Without this, "1\" SQ x 1/8\" Wall x 24'" extracts NO dims.
    private static readonly Regex _mixedFrac = new(@"(?<![\d./])(\d+)\s+(\d+)\s*/\s*(2|4|8|16|32|64)\b", RegexOptions.Compiled);
    private static readonly Regex _simpleFrac = new(@"(?<![\d./])(\d+)\s*/\s*(2|4|8|16|32|64)\b", RegexOptions.Compiled);
    private static readonly Regex _leadDot   = new(@"(?<![\d.])\.(\d+)", RegexOptions.Compiled);
    private static readonly Regex _unitWords = new(
        @"\b(SQ|SQUARE|RECT|RECTANGULAR|RND|ROUND|WALL|WA|THICK|THK|MILL|FINISH|MF|LONG|LG|RANDOM|RDM|OD|ID|DIA|PCS?|EA|FT|FOOT|FEET|IN|INCH|INCHES)\.?\b",
        RegexOptions.Compiled | RegexOptions.IgnoreCase);
    // A period that is NOT part of a decimal (e.g. the "." left behind by stripping "SQ." / "WA.") — it would
    // otherwise sit between a dimension number and the × and re-break the dim regex.
    private static readonly Regex _strayDot = new(@"(?<!\d)\.(?!\d)", RegexOptions.Compiled);

    private static string NormalizeDimText(string s)
    {
        string Fmt(double d) => d.ToString("0.####", System.Globalization.CultureInfo.InvariantCulture);
        s = _mixedFrac.Replace(s, m => Fmt(double.Parse(m.Groups[1].Value) + double.Parse(m.Groups[2].Value) / double.Parse(m.Groups[3].Value)));
        s = _simpleFrac.Replace(s, m => Fmt(double.Parse(m.Groups[1].Value) / double.Parse(m.Groups[2].Value)));
        s = _leadDot.Replace(s, "0.$1");
        s = s.Replace("\"", " ").Replace("'", " ");
        s = _unitWords.Replace(s, " ");
        s = _strayDot.Replace(s, " ");
        return s;
    }

    private static double[]? ExtractDims(string name)
    {
        var cleaned = NormalizeDimText(Regex.Replace(name, @"\([^)]*\)", " "));  // strip parentheticals + normalize

        var m3 = _dim3.Match(cleaned);
        if (m3.Success) return [D(m3, 1), D(m3, 2), D(m3, 3)];

        var m2 = _dim2.Match(cleaned);
        if (m2.Success) return [D(m2, 1), D(m2, 2)];

        var m1 = _dim1.Match(cleaned);
        if (m1.Success) return [D(m1, 1)];

        return null;

        static double D(Match m, int g) =>
            double.Parse(m.Groups[g].Value, System.Globalization.CultureInfo.InvariantCulture);
    }

    private static (double OD, double Wall)? ExtractPipeOdWall(string name)
    {
        var m1 = _pipeOdWallFmt1.Match(name);
        if (m1.Success) return (Parse(m1.Groups[1].Value), Parse(m1.Groups[2].Value));

        var m2 = _pipeOdWallFmt2.Match(name);
        if (m2.Success) return (Parse(m2.Groups[1].Value), Parse(m2.Groups[2].Value));

        var m3 = _pipeOdWallFmt3.Match(name);
        if (m3.Success) return (Parse(m3.Groups[1].Value), Parse(m3.Groups[2].Value));

        return null;

        static double Parse(string s) =>
            double.Parse(s, System.Globalization.CultureInfo.InvariantCulture);
    }

    /// <summary>
    /// Extracts I-beam section properties from the parenthetical: (H× × W× × FT × × WT ×).
    /// Returns (H, W, FT, WT) when all four are present, null otherwise.
    /// </summary>
    private static (double H, double W, double FT, double WT)? ExtractBeamHWFtWt(string name)
    {
        var m = _beamHWFtWt.Match(name);
        if (!m.Success) return null;
        return (P(m, 1), P(m, 2), P(m, 3), P(m, 4));
        static double P(Match m, int g) =>
            double.Parse(m.Groups[g].Value, System.Globalization.CultureInfo.InvariantCulture);
    }

    /// <summary>
    /// Extracts C-channel section properties from the parenthetical: (H× × W× × WT ×).
    /// Returns (H, W, WT) when all three are present, null otherwise.
    /// </summary>
    private static (double H, double W, double WT)? ExtractChannelHWWt(string name)
    {
        // Prefer full I-beam spec if present (beam misidentified as channel)
        var beam = ExtractBeamHWFtWt(name);
        if (beam.HasValue)
            return (beam.Value.H, beam.Value.W, beam.Value.WT);

        var m = _channelHWWt.Match(name);
        if (!m.Success) return null;
        return (P(m, 1), P(m, 2), P(m, 3));
        static double P(Match m, int g) =>
            double.Parse(m.Groups[g].Value, System.Globalization.CultureInfo.InvariantCulture);
    }

    // ── Main calculation ──────────────────────────────────────────────────────

    public static WeightResult Calculate(string productName)
    {
        if (string.IsNullOrWhiteSpace(productName))
            return new(null, null, null, "Empty product name");

        var metal = DetectMetal(productName);
        var shape = DetectShape(productName);
        var rho   = Density(metal);
        var dims  = ExtractDims(productName);

        static double R(double v) => Math.Round(v, 4);

        return shape switch
        {
            "flat_bar" when dims?.Length >= 2 =>
                new(R(dims[0] * dims[1] * rho * 12), null,
                    $"Flat Bar {dims[0]}×{dims[1]} ({metal})", null),

            "round_bar" when dims?.Length >= 1 =>
                new(R(Math.PI / 4 * dims[0] * dims[0] * rho * 12), null,
                    $"Round Bar D={dims[0]} ({metal})", null),

            "sq_bar" when dims?.Length >= 1 =>
                new(R(dims[0] * dims[0] * rho * 12), null,
                    $"Square Bar {dims[0]}×{dims[0]} ({metal})", null),

            "hex_bar" when dims?.Length >= 1 =>
                // area of regular hexagon with across-flats width F: A = (√3/2)×F²
                new(R(Math.Sqrt(3) / 2 * dims[0] * dims[0] * rho * 12), null,
                    $"Hex Bar F={dims[0]} ({metal})", null),

            "rect_tube" when dims?.Length >= 3 =>
                new(R((dims[0] * dims[1] - (dims[0] - 2 * dims[2]) * (dims[1] - 2 * dims[2])) * rho * 12), null,
                    $"Rect Tube {dims[0]}×{dims[1]}×{dims[2]} ({metal})", null),

            "sq_tube" when dims is not null && dims.Length >= 2 =>
                SquareTubeWeight(dims, rho, metal),

            "round_tube" when dims?.Length >= 2 =>
                new(R(Math.PI * dims[1] * (dims[0] - dims[1]) * rho * 12), null,
                    $"Round Tube OD={dims[0]} wall={dims[1]} ({metal})", null),

            "pipe" => PipeWeight(productName, rho, metal),

            "angle" when dims?.Length >= 3 =>
                new(R((dims[0] + dims[1] - dims[2]) * dims[2] * rho * 12), null,
                    $"Angle {dims[0]}×{dims[1]}×{dims[2]} ({metal})", null),

            "angle" when dims?.Length == 2 =>
                new(R((2 * dims[0] - dims[1]) * dims[1] * rho * 12), null,
                    $"Equal Angle {dims[0]}×{dims[0]}×{dims[1]} ({metal})", null),

            // I-beam (Wide Flange, Standard, H-Pile, Alum Beam).
            // Prefers H/W/FT/WT from parenthetical; formula: 2*(W*FT) + (H-2*FT)*WT
            "beam" => BeamWeight(productName, rho, metal),

            // C-channel.  If (H,W,FT,WT) is present use full formula; otherwise approximate
            // with uniform wall t: Area ≈ t*(H + 2W - 2t).
            "channel" => ChannelWeight(productName, dims, rho, metal),

            "sheet" when dims?.Length >= 1 =>
                // lb per square foot: thickness(in) × density(lb/in³) × 144(in²/sqft)
                new(null, R(dims[0] * rho * 144),
                    $"Sheet T={dims[0]} ({metal})", null),

            "tread_plate" when dims?.Length >= 1 =>
                // ~12% heavier than base sheet due to raised tread pattern
                new(null, R(dims[0] * rho * 144 * 1.12),
                    $"Tread Plate T={dims[0]} ({metal})", null),

            _ => new(null, null, null, $"Shape '{shape}' — not computable or missing dimensions")
        };
    }

    // ── Line $/lb weight basis (SINGLE source of truth) ───────────────────────
    // Resolves a line's TOTAL weight in lb for $/lb derivation, using the SAME basis everywhere it is
    // needed (the state-of-play summary, the SLI write, and the backfill): the supplier's stated
    // WeightPerUnit when present (exact), else a THEORETICAL estimate from the matched catalog product's
    // clean dimensions (flagged Estimated=true), falling back to the supplier's own free-text label only
    // when no catalog match. Returns (null,false) when no weight can be resolved (e.g. a flat sheet, whose
    // weight is per-sqft not per-foot, or dimensions that don't parse) — the caller then keeps the AI value.
    public static (double? TotalLb, bool Estimated) ResolveLineWeightLb(
        double qty, double? supplierWeightPerUnit, string? supplierWeightUnit,
        string? catalogProductName, string? supplierLabel,
        double? lengthPerUnit, string? lengthUnit)
    {
        if (qty <= 0) qty = 1;
        if (supplierWeightPerUnit is > 0)
            return (ToLb(supplierWeightPerUnit.Value, supplierWeightUnit) * qty, false);

        var catName = catalogProductName?.Trim() ?? "";
        var label   = supplierLabel?.Trim() ?? "";
        var name    = catName.Length > 0 ? catName : label;
        if (name.Length == 0) return (null, false);

        var wc = Calculate(name);
        if (wc.LbPerFoot is not > 0 && catName.Length > 0 && label.Length > 0)
            wc = Calculate(label);   // catalog name didn't parse → fall back to the supplier's label
        if (wc.LbPerFoot is > 0)
        {
            var lenFt = ToFeet(lengthPerUnit, lengthUnit);
            if (lenFt is > 0) return (wc.LbPerFoot.Value * lenFt.Value * qty, true);
        }
        return (null, false);
    }

    public static double ToLb(double v, string? unit) => (unit ?? "").Trim().ToLowerInvariant() switch
    {
        "kg" => v * 2.20462, "g" => v / 453.592, "oz" => v / 16.0,
        _    => v,   // lb / blank
    };

    public static double? ToFeet(double? v, string? unit)
    {
        if (v is not > 0) return null;
        return (unit ?? "").Trim().ToLowerInvariant() switch
        {
            "in" or "inch" or "inches" or "\"" => v / 12.0,
            "mm" => v / 304.8, "cm" => v / 30.48, "m" => v * 3.28084,
            _    => v,   // ft / blank
        };
    }

    private static WeightResult SquareTubeWeight(double[] dims, double rho, string metal)
    {
        double s, t;
        if (dims.Length >= 3 && Math.Abs(dims[0] - dims[1]) < 0.0001)
            (s, t) = (dims[0], dims[2]);   // "S × S × wall" format
        else
            (s, t) = (dims[0], dims[1]);   // "S × wall" format (less common)

        double area = 4 * t * (s - t);
        return new(Math.Round(area * rho * 12, 4), null,
            $"Square Tube {s}×{s}×{t} ({metal})", null);
    }

    private static WeightResult PipeWeight(string name, double rho, string metal)
    {
        var odwall = ExtractPipeOdWall(name);
        if (!odwall.HasValue)
            return new(null, null, null, "Pipe OD/wall spec not found in product name");

        double area = Math.PI * odwall.Value.Wall * (odwall.Value.OD - odwall.Value.Wall);
        return new(Math.Round(area * rho * 12, 4), null,
            $"Pipe OD={odwall.Value.OD} wall={odwall.Value.Wall} ({metal})", null);
    }

    private static WeightResult BeamWeight(string name, double rho, string metal)
    {
        var spec = ExtractBeamHWFtWt(name);
        if (!spec.HasValue)
            return new(null, null, null, "Beam H/W/FT/WT spec not found in product name");

        var (h, w, ft, wt) = spec.Value;
        // I-beam area: two flanges + web
        double area = 2 * w * ft + (h - 2 * ft) * wt;
        if (area <= 0) return new(null, null, null, "Beam area calculation non-positive");
        return new(Math.Round(area * rho * 12, 4), null,
            $"Beam H={h} W={w} FT={ft} WT={wt} ({metal})", null);
    }

    private static WeightResult ChannelWeight(string name, double[]? dims, double rho, string metal)
    {
        // Try H/W/WT from parenthetical first (most accurate for standard channels)
        var spec = ExtractChannelHWWt(name);
        if (spec.HasValue)
        {
            var (h, w, wt) = spec.Value;
            // C-channel with uniform wall t: Area ≈ t*(H + 2W - 2t)
            double area = wt * (h + 2 * w - 2 * wt);
            if (area > 0)
                return new(Math.Round(area * rho * 12, 4), null,
                    $"Channel H={h} W={w} t={wt} ({metal})", null);
        }

        // Fall back to 3-dim (W, H, t) extracted from name (Sharp Inside Corner style)
        if (dims?.Length >= 3)
        {
            double w = dims[0], h = dims[1], t = dims[2];
            double area = t * (w + h - t);
            if (area > 0)
                return new(Math.Round(area * rho * 12, 4), null,
                    $"Channel {w}×{h}×{t} ({metal})", null);
        }

        return new(null, null, null, "Channel dimensions not extractable from product name");
    }
}
