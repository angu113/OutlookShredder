namespace OutlookShredder.Proxy.Models.Drawing;

/// <summary>Supported formed-part presets. v1 implements <see cref="UChannel"/>.</summary>
public enum PartType
{
    LAngle,
    UChannel,
    ZChannel,
    Hat,
    Pan,
    Custom,
}

/// <summary>Whether a stated dimension is measured to the outside or the inside of the form.</summary>
public enum DimBasis
{
    Outside,
    Inside,
}

/// <summary>A single dimension value plus how it was measured.</summary>
public readonly record struct Dim(double Value, DimBasis Basis)
{
    public static Dim Outside(double v) => new(v, DimBasis.Outside);
    public static Dim Inside(double v) => new(v, DimBasis.Inside);
}

/// <summary>
/// Parsed, fully-resolved description of a part to develop. Produced by
/// <c>DrawingTextParser</c>; consumed by <c>BendMath</c> / <c>FlatPattern</c>.
/// Dimensions keep their measured basis; thickness compensation happens in FlatPattern so
/// the conversion rule stays in one place and is testable.
/// </summary>
public sealed class PartSpec
{
    public PartType Type { get; init; } = PartType.UChannel;

    // ── U-channel profile dimensions ────────────────────────────────────────
    /// <summary>Bottom (web) width.</summary>
    public Dim Web { get; init; }
    /// <summary>Left flange (leg) height.</summary>
    public Dim FlangeLeft { get; init; }
    /// <summary>Right flange (leg) height. Equals <see cref="FlangeLeft"/> unless two were given.</summary>
    public Dim FlangeRight { get; init; }
    /// <summary>The part run / depth. No inside/outside basis — it is the blank's other extent.</summary>
    public double Length { get; init; }

    // ── Material + bend parameters ──────────────────────────────────────────
    /// <summary>Resolved material thickness T, in the spec's units.</summary>
    public double Thickness { get; init; }
    /// <summary>Inside bend radius Ri. Defaults to T when the user does not state one.</summary>
    public double InsideRadius { get; init; }
    /// <summary>K-factor for the bend-allowance estimate.</summary>
    public double KFactor { get; init; }
    /// <summary>Bend angle in degrees (90 = right angle).</summary>
    public double AngleDeg { get; init; } = 90.0;
    /// <summary>When set, used directly as the per-bend deduction; the K estimate is skipped.</summary>
    public double? MeasuredBendDeduction { get; init; }

    /// <summary>Human-readable material label for display/labelling, e.g. "16ga CRS".</summary>
    public string Material { get; init; } = "";
    /// <summary>"in" or "mm".</summary>
    public string Units { get; init; } = "in";
}
