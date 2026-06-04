namespace OutlookShredder.Proxy.Services.Drawing;

/// <summary>
/// Deterministic air-bend math. All values in the part's units (inches by default).
///
///   OSSB = (Ri + T) * tan(theta/2)              outside setback
///   BA   = (pi/180) * theta * (Ri + K*T)         bend allowance
///   BD   = 2*OSSB - BA                           bend deduction (per bend)
///   Flat = sum(outside flange lengths) - n*BD
/// </summary>
public static class BendMath
{
    public static double DegToRad(double deg) => deg * Math.PI / 180.0;

    /// <summary>Outside setback per bend.</summary>
    public static double Ossb(double ri, double t, double angleDeg)
        => (ri + t) * Math.Tan(DegToRad(angleDeg) / 2.0);

    /// <summary>Bend allowance per bend (neutral-axis arc length).</summary>
    public static double BendAllowance(double ri, double t, double k, double angleDeg)
        => DegToRad(angleDeg) * (ri + k * t);

    /// <summary>
    /// Per-bend deduction. Returns the user's measured value verbatim when supplied,
    /// otherwise the K-estimate 2*OSSB - BA.
    /// </summary>
    public static double BendDeduction(double ri, double t, double k, double angleDeg, double? measured = null)
        => measured ?? (2.0 * Ossb(ri, t, angleDeg) - BendAllowance(ri, t, k, angleDeg));

    /// <summary>
    /// Default air-bend K-factor by material family. K differs by material (springback /
    /// neutral-axis shift): soft aluminium sits low, mild steel mid, hard stainless high.
    /// These are starting defaults — the user can override K explicitly per part, and the
    /// values are tunable. (Ri/T also influences K in full bend tables; folded in later.)
    /// </summary>
    public static double DefaultK(MaterialFamily family) => family switch
    {
        MaterialFamily.Aluminium  => 0.33,   // soft
        MaterialFamily.Copper     => 0.38,   // soft
        MaterialFamily.Brass      => 0.40,
        MaterialFamily.Stainless  => 0.45,   // hard, more springback
        MaterialFamily.Galvanized => 0.42,
        MaterialFamily.ColdRolled => 0.42,
        MaterialFamily.HotRolled  => 0.42,
        _ => 0.42,                           // mild-steel default
    };
}
