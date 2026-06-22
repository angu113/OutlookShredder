using OutlookShredder.Proxy.Models.CutOptimizer;
using RectangleBinPacking;

namespace OutlookShredder.Proxy.Services.CutOptimizer;

/// <summary>
/// 2D flat-sheet nesting for the cut optimizer. The placement strategy is chosen by the machine
/// profile (cut method), NEVER by best yield: a shear/guillotine job MUST use guillotine placement
/// (MaxRects packs denser but its layouts straddle split lines and are un-cuttable on a shear), and
/// only a free (laser/plasma) job may use MaxRects. Rectangular parts only for now; irregular outlines
/// route to libnest2d later, behind <see cref="INestStrategy"/>. See <c>wip/feat-cut-optimizer.md</c>.
/// </summary>
public static class Flat2DPacker
{
    /// <summary>Laser inter-part web (combined kerf + lead-in allowance), always applied on a free nest.</summary>
    public const double LaserKerf = 0.100;

    private const double MilPerInch = 1000.0;   // pack in integer thousandths to keep fractional precision

    public static List<Layout> Pack(
        string material, string gauge,
        IReadOnlyList<CutPart> parts, IReadOnlyList<StockSize> stock,
        CutMethod method, List<Issue> issues)
    {
        var sheets = stock.Where(s => s.Length > 0 && (s.Width ?? 0) > 0).ToList();
        // Expand parts (qty), inflate by the method's kerf/web, descending by area for a good packing order.
        double kerf = method == CutMethod.Laser ? LaserKerf : 0.0;
        var pieces = new List<Piece>();
        int idx = 0;
        foreach (var p in parts)
        {
            double w = (p.Width ?? 0), l = p.Length;
            if (w <= 0 || l <= 0) continue;
            for (int i = 0; i < Math.Max(1, p.Qty); i++)
                pieces.Add(new Piece(idx++, Mil(w + kerf), Mil(l + kerf), Mil(w), Mil(l), !p.FinishDirectionLocked, p.Label));
        }
        pieces.Sort((a, b) => (b.IW * (long)b.IH).CompareTo(a.IW * (long)a.IH));

        if (pieces.Count == 0) return new();
        if (sheets.Count == 0)
        {
            issues.Add(new Issue { Message = $"No stock for {GroupLabel(material, gauge)} — {pieces.Count} part(s) can't be cut." });
            return new();
        }

        // Drop parts that fit no sheet (even rotated when allowed).
        var sizes = sheets.Select(s => new SheetBucket(Mil(s.Width!.Value), Mil(s.Length), s.QtyAvailable ?? int.MaxValue, s.Width!.Value, s.Length)).ToList();
        var remaining = new List<Piece>();
        foreach (var pc in pieces)
        {
            if (sizes.Any(sz => Fits(pc, sz.W, sz.H))) remaining.Add(pc);
            else issues.Add(new Issue { Message = $"Part {Fmt(pc.TrueW)} x {Fmt(pc.TrueH)} fits no stock sheet for {GroupLabel(material, gauge)} — can't cut." });
        }

        INestStrategy strategy = method == CutMethod.Laser ? new MaxRectsStrategy() : new GuillotineStrategy();

        var layouts = new List<Layout>();
        int guard = remaining.Count + 2;
        while (remaining.Count > 0 && guard-- > 0)
        {
            // Candidate sizes that fit at least the largest remaining piece; prefer on-hand, then purchase.
            var candidates = sizes.Where(sz => remaining.Any(pc => Fits(pc, sz.W, sz.H))).ToList();
            if (candidates.Count == 0) break;
            var onHand = candidates.Where(sz => sz.UsedOnHand < sz.OnHand).ToList();
            bool purchasing = onHand.Count == 0;
            var pool = purchasing ? candidates : onHand;

            // Greedy: pick the size whose sheet packs the most part area this step.
            SheetBucket best = pool[0];
            List<Placement> bestPlaced = strategy.Nest(best.W, best.H, remaining);
            long bestArea = AreaOf(bestPlaced);
            foreach (var sz in pool.Skip(1))
            {
                var placed = strategy.Nest(sz.W, sz.H, remaining);
                long area = AreaOf(placed);
                if (area > bestArea || (area == bestArea && sz.W * (long)sz.H < best.W * (long)best.H))
                {
                    best = sz; bestPlaced = placed; bestArea = area;
                }
            }
            if (bestPlaced.Count == 0) break;   // nothing fit (guarded above; defensive)

            if (purchasing) best.Purchased++; else best.UsedOnHand++;

            double usefulSqIn = bestPlaced.Sum(pl => pl.TrueWIn * pl.TrueHIn);
            double sheetSqIn = best.WIn * best.HIn;
            layouts.Add(new Layout
            {
                Material = material,
                Gauge = gauge,
                StockLength = best.HIn,
                StockWidth = best.WIn,
                Pieces = bestPlaced.Select(pl => new PlacedPiece
                {
                    Label = pl.Label, Length = pl.TrueHIn,
                    X = pl.XIn, Y = pl.YIn, W = pl.TrueWIn, L = pl.TrueHIn, Rotated = pl.Rotated,
                }).ToList(),
                Drop = Math.Max(0, sheetSqIn - usefulSqIn),
                YieldPct = sheetSqIn > 0 ? 100.0 * usefulSqIn / sheetSqIn : 0,
                Purchased = purchasing,
            });

            var placedIds = bestPlaced.Select(p => p.Index).ToHashSet();
            remaining.RemoveAll(p => placedIds.Contains(p.Index));
        }

        if (remaining.Count > 0)
            issues.Add(new Issue { Message = $"{remaining.Count} part(s) for {GroupLabel(material, gauge)} could not be placed." });

        return layouts;
    }

    private static bool Fits(Piece p, int sheetW, int sheetH) =>
        (p.IW <= sheetW && p.IH <= sheetH) || (p.Rotatable && p.IH <= sheetW && p.IW <= sheetH);

    private static long AreaOf(List<Placement> placed) => placed.Sum(p => (long)p.IW * p.IH);

    private static int Mil(double inches) => (int)Math.Round(inches * MilPerInch);
    private static string Fmt(double v) => Services.Drawing.DrawFormat.FracInch(v);
    private static string GroupLabel(string material, string gauge) =>
        string.IsNullOrWhiteSpace(gauge) ? material : $"{material} {gauge}";

    // A required piece in mil units: inflated (IW/IH, for packing) + true (for reporting).
    private readonly record struct Piece(int Index, int IW, int IH, int TrueWMil, int TrueHMil, bool Rotatable, string? Label)
    {
        public double TrueW => TrueWMil / MilPerInch;
        public double TrueH => TrueHMil / MilPerInch;
    }

    // A placed piece in mil units (inflated rect from the packer) + the true reported size in inches.
    private sealed class Placement
    {
        public int Index;
        public int XMil, YMil, IW, IH;
        public int TrueWMil, TrueHMil;
        public bool Rotated;
        public string? Label;
        public double XIn => XMil / MilPerInch;
        public double YIn => YMil / MilPerInch;
        public double TrueWIn => TrueWMil / MilPerInch;
        public double TrueHIn => TrueHMil / MilPerInch;
    }

    private sealed class SheetBucket
    {
        public SheetBucket(int w, int h, int onHand, double wIn, double hIn) { W = w; H = h; OnHand = onHand; WIn = wIn; HIn = hIn; }
        public int W, H, OnHand, UsedOnHand, Purchased;
        public double WIn, HIn;
    }

    /// <summary>Packs as many of <paramref name="parts"/> as fit on ONE sheet (mil units); returns placements.</summary>
    private interface INestStrategy
    {
        List<Placement> Nest(int sheetW, int sheetH, IReadOnlyList<Piece> parts);
    }

    // ── Shear: guillotine (edge-to-edge cuts). Cuttability over density. ─────────
    private sealed class GuillotineStrategy : INestStrategy
    {
        public List<Placement> Nest(int sheetW, int sheetH, IReadOnlyList<Piece> parts)
        {
            var bin = new GuillotineBinPack(sheetW, sheetH);
            var placed = new List<Placement>();
            foreach (var p in parts)
            {
                var r = bin.Insert(p.IW, p.IH, true,
                    GuillotineBinPack.FreeRectChoiceHeuristic.RectBestAreaFit,
                    GuillotineBinPack.GuillotineSplitHeuristic.SplitMinimizeArea);
                bool rotated = false;
                if ((r.Width <= 0 || r.Height <= 0) && p.Rotatable && p.IW != p.IH)
                {
                    r = bin.Insert(p.IH, p.IW, true,
                        GuillotineBinPack.FreeRectChoiceHeuristic.RectBestAreaFit,
                        GuillotineBinPack.GuillotineSplitHeuristic.SplitMinimizeArea);
                    rotated = true;
                }
                if (r.Width > 0 && r.Height > 0) placed.Add(Place(p, r, rotated));
            }
            return placed;
        }
    }

    // ── Free (laser): MaxRects (denser non-guillotine). Rotation controlled per-part. ──
    private sealed class MaxRectsStrategy : INestStrategy
    {
        public List<Placement> Nest(int sheetW, int sheetH, IReadOnlyList<Piece> parts)
        {
            var bin = new MaxRectsBinPack(sheetW, sheetH, false);   // rotation off — we control it per part
            var placed = new List<Placement>();
            foreach (var p in parts)
            {
                var r = bin.Insert(p.IW, p.IH, FreeRectChoiceHeuristic.RectBestShortSideFit);
                bool rotated = false;
                if ((r.Width <= 0 || r.Height <= 0) && p.Rotatable && p.IW != p.IH)
                {
                    r = bin.Insert(p.IH, p.IW, FreeRectChoiceHeuristic.RectBestShortSideFit);
                    rotated = true;
                }
                if (r.Width > 0 && r.Height > 0) placed.Add(Place(p, r, rotated));
            }
            return placed;
        }
    }

    private static Placement Place(Piece p, Rect r, bool rotated) => new()
    {
        Index = p.Index,
        XMil = r.X, YMil = r.Y, IW = r.Width, IH = r.Height,
        TrueWMil = rotated ? p.TrueHMil : p.TrueWMil,
        TrueHMil = rotated ? p.TrueWMil : p.TrueHMil,
        Rotated = rotated,
        Label = p.Label,
    };
}
