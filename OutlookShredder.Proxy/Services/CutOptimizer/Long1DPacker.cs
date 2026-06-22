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

    /// <summary>De-minimis leftover for long product: a drop this short (≤ 12") is negligible waste, NOT
    /// scrap — so it's fine to make the cut that leaves it rather than preserve a reusable remnant. Only a
    /// leftover BETWEEN this and the reuse minimum counts as scrap to avoid.</summary>
    public const double Deminimis = 12.0;

    private const double Eps = 1e-9;

    /// <summary>
    /// Pack one (material, gauge) group. <paramref name="parts"/> and <paramref name="stock"/> are
    /// already filtered to this group. Appends un-cuttable / no-stock problems to <paramref name="issues"/>.
    /// </summary>
    public static List<Layout> Pack(
        string material, string gauge,
        IReadOnlyList<CutPart> parts, IReadOnlyList<StockSize> stock,
        bool precision, List<Issue> issues, double minUsableDrop = 0)
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

        if (minUsableDrop <= 0)
            return Better(large, small);   // original objective: least stock, then fewest bars

        // Soft "usable drop" preference: also try a drop-aware plan, then pick the one with the least
        // SCRAP (a leftover in (0, min) — too short to reuse), tie-broken by least stock then fewest bars.
        var dropAware = PackDropAware(material, gauge, cuttable, stock, kerf, minUsableDrop);
        return BestByScrap(new[] { large, small, dropAware }, minUsableDrop);
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

            Place(best, eff, pc);
        }

        return ToLayouts(material, gauge, bars);
    }

    /// <summary>
    /// Drop-aware placement (only used when a usable-drop minimum is set). For each piece, prefer an
    /// existing bar whose remainder stays "good" (≈0 or ≥ min); otherwise open a fresh bar when that
    /// leaves a reusable (or zero) remainder rather than squeezing the piece into a bar and creating
    /// short scrap. Costs more stock than tight packing — that's the point (reusable remnants).
    /// </summary>
    private static List<Layout> PackDropAware(
        string material, string gauge,
        List<(double Len, string? Label)> cuttable, IReadOnlyList<StockSize> stock,
        double kerf, double min)
    {
        var sizes = stock.Select(s => new Bucket(s.Length, s.QtyAvailable ?? int.MaxValue)).ToList();
        var openOrder = sizes.OrderByDescending(b => b.Length).ToList();   // largest first -> bigger reusable drops

        var bars = new List<Bar>();
        foreach (var pc in cuttable)
        {
            double eff = pc.Len + kerf;

            Bar? bestGood = null; double bestGoodRem = double.MaxValue;
            Bar? bestAny = null; double bestAnyRem = double.MaxValue;
            foreach (var b in bars)
                if (b.Remaining + Eps >= eff)
                {
                    double post = b.Remaining - eff;
                    // "Good" = the leftover is either negligible (<= de-minimis) or a reusable remnant (>= min);
                    // a leftover strictly between the two is the scrap we try to avoid.
                    bool good = post <= Deminimis + Eps || post + Eps >= min;
                    if (post < bestAnyRem) { bestAnyRem = post; bestAny = b; }
                    if (good && post < bestGoodRem) { bestGoodRem = post; bestGood = b; }
                }

            if (bestGood is not null) { Place(bestGood, eff, pc); continue; }

            // No existing bar takes it cleanly — open a new one if that leaves a negligible/reusable remainder.
            double newPost = openOrder.Where(s => s.Length + Eps >= eff).Select(s => s.Length - eff).DefaultIfEmpty(-1).Max();
            bool newGood = newPost >= 0 && (newPost <= Deminimis + Eps || newPost + Eps >= min);
            if (newGood)
            {
                var nb = OpenBar(openOrder, eff);
                if (nb is not null) { bars.Add(nb); Place(nb, eff, pc); continue; }
            }
            if (bestAny is not null) { Place(bestAny, eff, pc); continue; }   // unavoidable scrap
            var fb = OpenBar(openOrder, eff);
            if (fb is not null) { bars.Add(fb); Place(fb, eff, pc); }
        }
        return ToLayouts(material, gauge, bars);
    }

    private static void Place(Bar b, double eff, (double Len, string? Label) pc)
    {
        b.Remaining -= eff;
        b.Pieces.Add(new PlacedPiece { Length = pc.Len, Label = pc.Label });
    }

    private static List<Layout> ToLayouts(string material, string gauge, List<Bar> bars) =>
        bars.Select(b => new Layout
        {
            Material = material,
            Gauge = gauge,
            StockLength = b.StockLength,
            Pieces = b.Pieces,
            Drop = Math.Max(0, b.Remaining),
            YieldPct = b.StockLength > 0 ? 100.0 * b.Pieces.Sum(p => p.Length) / b.StockLength : 0,
            Purchased = b.Purchased,
        }).ToList();

    /// <summary>Pick the plan with the least scrap (drops in (0, min)); tie-break least stock, then fewest bars.</summary>
    private static List<Layout> BestByScrap(IEnumerable<List<Layout>> plans, double min)
    {
        List<Layout>? best = null;
        double bestScrap = 0, bestStock = 0; int bestBars = 0;
        foreach (var p in plans)
        {
            double scrap = Scrap(p, min);
            double stock = p.Sum(l => l.StockLength);
            int barsN = p.Count;
            bool take = best is null
                || scrap + Eps < bestScrap
                || (Math.Abs(scrap - bestScrap) <= Eps && stock + Eps < bestStock)
                || (Math.Abs(scrap - bestScrap) <= Eps && Math.Abs(stock - bestStock) <= Eps && barsN < bestBars);
            if (take) { best = p; bestScrap = scrap; bestStock = stock; bestBars = barsN; }
        }
        return best ?? new();
    }

    // Scrap = a leftover too big to be negligible yet too small to reuse: de-minimis < drop < min.
    private static double Scrap(List<Layout> plan, double min) =>
        plan.Sum(l => l.Drop > Deminimis + Eps && l.Drop + Eps < min ? l.Drop : 0);

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
