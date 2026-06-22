using OutlookShredder.Proxy.Models.CutOptimizer;

namespace OutlookShredder.Proxy.Services.CutOptimizer;

/// <summary>
/// 1D cutting-stock packer for long material (bar / tube / extrusion). Best-Fit-Decreasing with a
/// per-cut kerf: required pieces are cut from stock lengths, minimising waste (total stock consumed),
/// tie-broken by fewer bars. Honours on-hand stock quantities and reports what must be purchased to
/// finish (see <c>wip/feat-cut-optimizer.md</c>, Decision 4).
/// </summary>
public static class Long1DPacker
{
    /// <summary>Kerf consumed per cut when "Precision needed" is on — 1/8". A named, tunable constant.</summary>
    public const double LongKerf = 0.125;

    private const double Eps = 1e-9;

    /// <summary>
    /// Pack one (material, gauge) group. <paramref name="parts"/> and <paramref name="stock"/> are
    /// already filtered to this group. Appends un-cuttable / no-stock problems to <paramref name="issues"/>.
    /// </summary>
    public static List<Layout> Pack(
        string material, string gauge,
        IReadOnlyList<CutPart> parts, IReadOnlyList<StockSize> stock,
        bool precision, List<Issue> issues)
    {
        double kerf = precision ? LongKerf : 0.0;

        // Expand to individual pieces, longest first (decreasing).
        var pieces = new List<(double Len, string? Label)>();
        foreach (var p in parts)
            for (int i = 0; i < Math.Max(1, p.Qty); i++)
                pieces.Add((p.Length, p.Label));
        pieces.Sort((a, b) => b.Len.CompareTo(a.Len));

        if (stock.Count == 0)
        {
            if (pieces.Count > 0)
                issues.Add(new Issue { Message = $"No stock for {Group(material, gauge)} — {pieces.Count} part(s) can't be cut." });
            return new();
        }

        double maxStock = stock.Max(s => s.Length);
        var cuttable = new List<(double Len, string? Label)>();
        foreach (var pc in pieces)
        {
            if (pc.Len + kerf > maxStock + Eps)
                issues.Add(new Issue { Message = $"Part {Fmt(pc.Len)} exceeds the largest stock length {Fmt(maxStock)} for {Group(material, gauge)} — can't cut." });
            else
                cuttable.Add(pc);
        }
        if (cuttable.Count == 0) return new();

        // Two open-bar strategies (prefer largest / smallest fitting size); keep the less-wasteful plan.
        var large = PackOnce(material, gauge, cuttable, stock, kerf, preferLargest: true);
        var small = PackOnce(material, gauge, cuttable, stock, kerf, preferLargest: false);
        return Better(large, small);
    }

    private static List<Layout> PackOnce(
        string material, string gauge,
        List<(double Len, string? Label)> cuttable, IReadOnlyList<StockSize> stock,
        double kerf, bool preferLargest)
    {
        var sizes = stock.Select(s => new Bucket(s.Length, s.QtyAvailable ?? int.MaxValue)).ToList();
        // Deterministic candidate order for opening a fresh bar.
        var openOrder = (preferLargest
                ? sizes.OrderByDescending(b => b.Length)
                : sizes.OrderBy(b => b.Length))
            .ToList();

        var bars = new List<Bar>();
        foreach (var pc in cuttable)
        {
            double eff = pc.Len + kerf;

            // Best fit: the open bar whose remaining space is smallest but still fits this piece.
            Bar? best = null;
            foreach (var b in bars)
                if (b.Remaining + Eps >= eff && (best is null || b.Remaining < best.Remaining))
                    best = b;

            if (best is null)
            {
                best = OpenBar(openOrder, eff);
                if (best is null) continue;   // guarded upstream (cuttable guarantees a fit)
                bars.Add(best);
            }

            best.Remaining -= eff;
            best.Pieces.Add(new PlacedPiece { Length = pc.Len, Label = pc.Label });
        }

        return bars.Select(b => new Layout
        {
            Material = material,
            Gauge = gauge,
            StockLength = b.StockLength,
            Pieces = b.Pieces,
            Drop = Math.Max(0, b.Remaining),
            YieldPct = b.StockLength > 0 ? 100.0 * b.Pieces.Sum(p => p.Length) / b.StockLength : 0,
            Purchased = b.Purchased,
        }).ToList();
    }

    /// <summary>
    /// Open a fresh bar for a piece needing <paramref name="eff"/> length. Use on-hand stock first
    /// (any size with remaining quantity that fits); only when no on-hand size fits do we buy one,
    /// flagging the bar as a purchase. Unlimited sizes always count as on-hand (never purchased).
    /// </summary>
    private static Bar? OpenBar(List<Bucket> openOrder, double eff)
    {
        foreach (var b in openOrder)
            if (b.Length + Eps >= eff && b.UsedOnHand < b.OnHand)
            {
                b.UsedOnHand++;
                return new Bar(b.Length, purchased: false);
            }
        foreach (var b in openOrder)
            if (b.Length + Eps >= eff)
            {
                b.Purchased++;
                return new Bar(b.Length, purchased: true);
            }
        return null;
    }

    /// <summary>Least total stock consumed wins; tie-break on fewer bars.</summary>
    private static List<Layout> Better(List<Layout> a, List<Layout> b)
    {
        double Wa = a.Sum(l => l.StockLength), Wb = b.Sum(l => l.StockLength);
        if (Math.Abs(Wa - Wb) > Eps) return Wa < Wb ? a : b;
        return a.Count <= b.Count ? a : b;
    }

    private static string Group(string material, string gauge) =>
        string.IsNullOrWhiteSpace(gauge) ? material : $"{material} {gauge}";

    private static string Fmt(double v) =>
        Services.Drawing.DrawFormat.FracInch(v);

    // Per-stock-size accounting across one packing run.
    private sealed class Bucket
    {
        public Bucket(double length, int onHand) { Length = length; OnHand = onHand; }
        public double Length { get; }
        public int OnHand { get; }
        public int UsedOnHand;
        public int Purchased;
    }

    private sealed class Bar
    {
        public Bar(double stockLength, bool purchased) { StockLength = stockLength; Remaining = stockLength; Purchased = purchased; }
        public double StockLength { get; }
        public double Remaining;
        public bool Purchased { get; }
        public List<PlacedPiece> Pieces { get; } = new();
    }
}
