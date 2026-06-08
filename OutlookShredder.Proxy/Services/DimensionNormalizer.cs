using System.Globalization;
using System.Text.RegularExpressions;
using OutlookShredder.Proxy.Models;

namespace OutlookShredder.Proxy.Services;

/// <summary>
/// Shared canonical-dimension utility: converts every dimension to DECIMAL INCHES (3 dp) on ONE basis so
/// supplier wording and catalog wording compare equally. Handles fractions (3/16 -> 0.188), feet (4' -> 48),
/// leading-dot decimals (.188 -> 0.188) and sheet-metal gauge per material (steel MSG / stainless / galvanized
/// / BWG for tube&amp;pipe / Brown &amp; Sharpe for non-ferrous — verified against catalog products carrying both
/// gauge and decimal). Used by the RFQ winner pool (via <c>GaugeToInches</c>) and the catalog tokenizer (via
/// <c>CanonicalizeDims</c> / <c>Apply</c>) so the same product yields the same MSPC / CUST_ id everywhere.
/// </summary>
public static class DimensionNormalizer
{
    // Manufacturers' Standard Gauge — carbon steel SHEET.
    private static readonly Dictionary<int, double> Steel = new() {
        {3,0.2391},{4,0.2242},{5,0.2092},{6,0.1943},{7,0.1793},{8,0.1644},{9,0.1495},{10,0.1345},{11,0.1196},
        {12,0.1046},{13,0.0897},{14,0.0747},{15,0.0673},{16,0.0598},{17,0.0538},{18,0.0478},{19,0.0418},
        {20,0.0359},{21,0.0329},{22,0.0299},{23,0.0269},{24,0.0239},{25,0.0209},{26,0.0179},{28,0.0149},{30,0.0120} };
    // Stainless SHEET gauge.
    private static readonly Dictionary<int, double> Ss = new() {
        {7,0.1875},{8,0.1719},{9,0.1563},{10,0.1406},{11,0.1250},{12,0.1094},{13,0.0938},{14,0.0781},{15,0.0703},
        {16,0.0625},{17,0.0563},{18,0.0500},{19,0.0438},{20,0.0375},{22,0.0313},{24,0.0250},{26,0.0188},{28,0.0156} };
    // Galvanized SHEET gauge (zinc coating adds vs bare steel).
    private static readonly Dictionary<int, double> Galv = new() {
        {8,0.1681},{9,0.1532},{10,0.1382},{11,0.1233},{12,0.1084},{13,0.0934},{14,0.0785},{15,0.0710},{16,0.0635},
        {17,0.0575},{18,0.0516},{19,0.0456},{20,0.0396},{21,0.0366},{22,0.0336},{24,0.0276},{26,0.0217},{28,0.0187} };
    // Birmingham Wire Gauge — TUBE/PIPE wall (any metal).
    private static readonly Dictionary<int, double> Bwg = new() {
        {4,0.238},{5,0.220},{6,0.203},{7,0.180},{8,0.165},{9,0.148},{10,0.134},{11,0.120},{12,0.109},{13,0.095},
        {14,0.083},{15,0.072},{16,0.065},{17,0.058},{18,0.049},{19,0.042},{20,0.035},{21,0.032},{22,0.028},{24,0.022} };
    // Brown & Sharpe (AWG) — non-ferrous SHEET (aluminum / brass / copper).
    private static readonly Dictionary<int, double> BnS = new() {
        {6,0.162},{7,0.144},{8,0.129},{9,0.114},{10,0.102},{11,0.091},{12,0.081},{13,0.072},{14,0.064},{15,0.057},
        {16,0.051},{17,0.045},{18,0.040},{19,0.036},{20,0.032},{21,0.029},{22,0.025},{23,0.023},{24,0.020},{25,0.018},{26,0.016} };

    /// <summary>Sheet-metal gauge -> decimal inches. tube=true selects BWG; otherwise the SHEET table for the
    /// metal class ("ss"/"stainless", "galv", "alum"/"brass"/"copper" -> B&amp;S, else carbon-steel MSG).</summary>
    public static double? GaugeToInches(string? metal, bool tube, int ga)
    {
        var m = (metal ?? "").ToLowerInvariant();
        var t = tube                                                        ? Bwg
              : m.Contains("stainless") || m == "ss"                        ? Ss
              : m.Contains("galv")                                          ? Galv
              : m.Contains("alum") || m.Contains("brass") || m.Contains("copper") ? BnS
              :                                                               Steel;
        return t.TryGetValue(ga, out var v) ? v : (double?)null;
    }

    private static readonly Regex RxGauge    = new(@"^#?\s*(\d{1,2})\s*(?:gauge|ga)$", RegexOptions.IgnoreCase | RegexOptions.Compiled);
    private static readonly Regex RxFraction = new(@"^(?:(\d+)\s+)?(\d+)\s*/\s*(\d+)$", RegexOptions.Compiled);

    /// <summary>One dimension token -> a 3-dp decimal-inch string. Handles gauge (per metal/shape), fractions
    /// (incl. "1 1/2"), leading-dot decimals and plain decimals. Non-numeric tokens are returned unchanged.</summary>
    public static string NormalizeToken(string token, string? metal, bool tube)
    {
        var t = (token ?? "").Trim();
        if (t.Length == 0 || t.Equals("null", StringComparison.OrdinalIgnoreCase)) return t;

        var g = RxGauge.Match(t);
        if (g.Success && int.TryParse(g.Groups[1].Value, out var ga) && GaugeToInches(metal, tube, ga) is double gth)
            return gth.ToString("0.###", CultureInfo.InvariantCulture);

        var f = RxFraction.Match(t);
        if (f.Success && double.TryParse(f.Groups[2].Value, out var num) && double.TryParse(f.Groups[3].Value, out var den) && den != 0)
        {
            double whole = f.Groups[1].Success && double.TryParse(f.Groups[1].Value, out var w) ? w : 0;
            return Math.Round(whole + num / den, 3).ToString("0.###", CultureInfo.InvariantCulture);
        }

        // Decimal / leading-dot / integer (NumberStyles.Any accepts ".188").
        if (double.TryParse(t, NumberStyles.Any, CultureInfo.InvariantCulture, out var d))
            return Math.Round(d, 3).ToString("0.###", CultureInfo.InvariantCulture);

        return t;   // non-numeric (e.g. a stray word) — leave as-is
    }

    /// <summary>Canonicalises an "AxBxC" dimension string to decimal inches so "3/16x48x96", "0.188x48x96" and
    /// "7gax48x96" all collapse to one value. Order is preserved; empty/null returns the input.</summary>
    public static string? CanonicalizeDims(string? dims, string? metal, string? shape)
    {
        if (string.IsNullOrWhiteSpace(dims)) return dims;
        var s = (shape ?? "").ToLowerInvariant();
        bool tube = s.Contains("tube") || s.Contains("tubing") || s.Contains("pipe");
        var parts = dims.Split('x', StringSplitOptions.TrimEntries)
                        .Select(p => NormalizeToken(p, metal, tube));
        return string.Join("x", parts);
    }

    /// <summary>Rewrites a token bag's <c>TkDims</c> to the canonical decimal-inch form, in place.</summary>
    public static void Apply(ProductTokens? t)
    {
        if (t is null || string.IsNullOrWhiteSpace(t.TkDims)) return;
        t.TkDims = CanonicalizeDims(t.TkDims, t.TkMetal, t.TkShape);
    }
}
