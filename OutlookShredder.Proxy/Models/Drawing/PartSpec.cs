namespace OutlookShredder.Proxy.Models.Drawing;

/// <summary>Supported formed-part presets. v1 implements <see cref="UChannel"/>.</summary>
public enum PartType
{
    LAngle,
    UChannel,
    ZChannel,
    Hat,
    Pan,
    FlitchPlate,
    BasePlate,
    PaddleBlind,
    Column,
    Custom,
}

/// <summary>How holes are laid out on a flat plate.</summary>
public enum HolePattern { Staggered, Paired, Corner }

/// <summary>Which way a flange folds relative to the web plane: Up = above, Down = below.</summary>
public enum BendDir { Up, Down }

/// <summary>One bend: its angle from flat (90 = square corner) and fold direction.</summary>
public sealed record BendSpec(double AngleDeg, BendDir Direction);

/// <summary>
/// Which face carries the finish (brushed stainless, paint, etc.). Inside/Outside for
/// U / L / pan (relative to the bend direction); Top/Bottom for Z (relative to the first flange).
/// </summary>
public enum FinishSide { None, Inside, Outside, Top, Bottom }

/// <summary>Bolt-hole specification for a flat plate.</summary>
public sealed class HoleSpec
{
    public double Diameter { get; init; }
    public HolePattern Pattern { get; init; }
    public double Spacing { get; init; }       // flitch: hole spacing along the length
    public int Count { get; init; }            // base: number of holes
    public double EdgeDistance { get; init; }  // base: edge distance for corner holes
    public double LeftEndOffset { get; init; }   // flitch: LHS edge -> first hole
    public double RightEndOffset { get; init; }  // flitch: RHS edge -> last hole (may differ from left)
    public double TopEdge { get; init; }       // flitch: top material edge -> top row
    public double BottomEdge { get; init; }    // flitch: bottom material edge -> bottom row
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

    /// <summary>Plate width (for flat plates; Length is the plate length). Also the pan base width.</summary>
    public double Width { get; init; }

    /// <summary>Pan wall height.</summary>
    public double Depth { get; init; }
    // Inside/outside basis for the pan's base length, base width, and wall depth.
    public DimBasis LengthBasis { get; init; } = DimBasis.Outside;
    public DimBasis WidthBasis { get; init; } = DimBasis.Outside;
    public DimBasis DepthBasis { get; init; } = DimBasis.Outside;
    // Which pan walls are present. Bottom/Top run along the Length ("long" sides);
    // Left/Right run along the Width ("short" sides). A 3-sided pan omits one.
    public bool PanBottom { get; init; } = true;
    public bool PanTop { get; init; } = true;
    public bool PanLeft { get; init; } = true;
    public bool PanRight { get; init; } = true;
    /// <summary>Bolt holes for flat plates.</summary>
    public HoleSpec? Holes { get; init; }

    // ── Paddle blind (ASME B16.48 spade / "frying pan") ─────────────────────
    /// <summary>Spade disc outside diameter.</summary>
    public double PaddleOd { get; init; }
    /// <summary>Handle width.</summary>
    public double PaddleHandleWidth { get; init; }
    /// <summary>Disc centre to the end of the handle.</summary>
    public double PaddleCenterToEnd { get; init; }
    /// <summary>NPS label for display, e.g. "6" or "1-1/4".</summary>
    public string PaddleNps { get; init; } = "";
    /// <summary>Pressure class (150 or 300).</summary>
    public int PaddleClass { get; init; }

    // ── Structural column (base plate + tube/pipe + bearing plate) ──────────
    /// <summary>Full gap the column supports (overall height = base T + tube length + bearing T).</summary>
    public double ColumnFullHeight { get; init; }
    public double BaseLength { get; init; }
    public double BaseWidth { get; init; }
    public double BaseThickness { get; init; }
    public HoleSpec? BaseHoles { get; init; }
    public double BearingLength { get; init; }
    public double BearingWidth { get; init; }
    public double BearingThickness { get; init; }
    public HoleSpec? BearingHoles { get; init; }
    /// <summary>"round" (pipe), "square" or "rect" tube.</summary>
    public string ColumnShape { get; init; } = "square";
    /// <summary>Outer width (square/rect tube) or outside diameter (round pipe).</summary>
    public double ColumnOuterWidth { get; init; }
    /// <summary>Outer depth (rect tube). Equals the width for square / round.</summary>
    public double ColumnOuterDepth { get; init; }
    /// <summary>Tube/pipe wall thickness (informational; shown in the note).</summary>
    public double ColumnWall { get; init; }
    /// <summary>Product label for the title / cut note, e.g. "HSS 4x4x1/4".</summary>
    public string ColumnLabel { get; init; } = "";

    // ── Material + bend parameters ──────────────────────────────────────────
    /// <summary>Resolved material thickness T, in the spec's units.</summary>
    public double Thickness { get; init; }
    /// <summary>Inside bend radius Ri. Defaults to T when the user does not state one.</summary>
    public double InsideRadius { get; init; }
    /// <summary>K-factor for the bend-allowance estimate.</summary>
    public double KFactor { get; init; }
    /// <summary>Bend angle in degrees (90 = right angle). Fallback when <see cref="Bends"/> is null.</summary>
    public double AngleDeg { get; init; } = 90.0;
    /// <summary>When set, used directly as the per-bend deduction; the K estimate is skipped.</summary>
    public double? MeasuredBendDeduction { get; init; }

    /// <summary>
    /// Per-bend angle + fold direction for U / L / Z. Order follows the flanges left→right
    /// (U/Z: [flange-left bend, flange-right bend]; L: [the single bend]). Null ⇒ derive from
    /// <see cref="AngleDeg"/> + the shape's default directions (back-compat / shorthand input).
    /// </summary>
    public IReadOnlyList<BendSpec>? Bends { get; init; }
    /// <summary>True when the input carried explicit per-bend angles (gates the degree/arc callouts).</summary>
    public bool AnglesAnnotated { get; init; }

    /// <summary>Human-readable material label for display/labelling, e.g. "16ga CRS".</summary>
    public string Material { get; init; } = "";
    /// <summary>"in" or "mm".</summary>
    public string Units { get; init; } = "in";

    /// <summary>Which face carries the finish (and gets the "Finish" arrow on the drawing).</summary>
    public FinishSide Finish { get; init; } = FinishSide.None;
}
