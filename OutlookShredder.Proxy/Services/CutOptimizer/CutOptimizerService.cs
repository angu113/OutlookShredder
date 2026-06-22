using System.Text;
using OutlookShredder.Proxy.Models.CutOptimizer;
using OutlookShredder.Proxy.Services.Drawing;

namespace OutlookShredder.Proxy.Services.CutOptimizer;

/// <summary>
/// Orchestrates a cut-optimization run: validate, group parts by (material, gauge) so a part is only
/// ever cut from same-material stock, dispatch each group to the 1D (long) or 2D (flat) packer, then
/// assemble the structured plan + a plain-text summary. The PDF report is added in Phase 3.
/// </summary>
public static class CutOptimizerService
{
    public static OptimizeResult Optimize(OptimizeRequest req)
    {
        var result = new OptimizeResult();
        if (req is null) { result.Issues.Add(new Issue { Message = "Empty request." }); result.TextSummary = "Nothing to optimize."; return result; }

        var parts = (req.Parts ?? new()).Where(p => p.Length > 0).ToList();
        var stock = (req.Stock ?? new()).Where(s => s.Length > 0).ToList();
        if (parts.Count == 0)
        {
            result.Issues.Add(new Issue { Message = "Add at least one part to cut." });
            result.TextSummary = "Nothing to optimize — add parts to cut.";
            return result;
        }

        bool isLong = req.ResolvedForm == MaterialForm.Long;

        // Group by (material, gauge); a part is only cut from same-material stock.
        var sb = new StringBuilder();
        foreach (var grp in parts.GroupBy(p => Key(p.Material, p.Gauge)).OrderBy(g => g.Key))
        {
            var groupParts = grp.ToList();
            var groupStock = stock.Where(s => Key(s.Material, s.Gauge) == grp.Key).ToList();
            string material = groupParts[0].Material;
            string gauge = groupParts[0].Gauge;

            List<Layout> layouts = isLong
                ? Long1DPacker.Pack(material, gauge, groupParts, groupStock, req.PrecisionNeeded, result.Issues)
                : Flat2DPacker.Pack(material, gauge, groupParts, groupStock, req.ResolvedMethod, result.Issues);

            result.Layouts.AddRange(layouts);
            AppendGroupSummary(sb, material, gauge, layouts, isLong);
        }

        BuildUsageAndPurchases(result);
        AppendGlobalSummary(sb, result);
        result.TextSummary = sb.ToString().TrimEnd();

        // Phase 3: a printable PDF report (only when there's an actual plan to draw).
        if (result.Layouts.Count > 0)
            result.PdfBase64 = Convert.ToBase64String(
                CutLayoutPdfRenderer.Render(result, req.ResolvedForm, req.ResolvedMethod, req.PrecisionNeeded));

        return result;
    }

    private static void AppendGroupSummary(StringBuilder sb, string material, string gauge, List<Layout> layouts, bool isLong)
    {
        if (layouts.Count == 0) return;
        // Long jobs carry no metal/gauge (the client hides those for long) — then just name the form.
        bool noMaterial = string.IsNullOrWhiteSpace(material) && string.IsNullOrWhiteSpace(gauge);
        sb.AppendLine(noMaterial
            ? (isLong ? "Long stock" : "Flat sheet")
            : $"{GroupLabel(material, gauge)} — {(isLong ? "long stock" : "flat sheet")}");

        int parts = layouts.Sum(l => l.Pieces.Count);
        // Yield basis: total stock length (long, 1D) vs total sheet area (flat, 2D).
        double useful = isLong ? layouts.Sum(l => l.Pieces.Sum(p => p.Length))
                               : layouts.Sum(l => l.Pieces.Sum(p => p.W * p.L));
        double total = isLong ? layouts.Sum(l => l.StockLength)
                              : layouts.Sum(l => l.StockLength * (l.StockWidth ?? 0));
        double yield = total > 0 ? 100.0 * useful / total : 0;

        // One "need N @ L" / "need N sheets W x L" line per distinct stock size used.
        foreach (var bySize in layouts.GroupBy(l => (l.StockWidth, l.StockLength)).OrderByDescending(g => g.Key.StockLength))
        {
            int n = bySize.Count();
            string noun = isLong ? "length" : "sheet";
            sb.AppendLine($"  Need {n} {noun}{(n == 1 ? "" : "s")} {SizeLabel(bySize.Key.StockWidth, bySize.Key.StockLength)}");
        }
        sb.AppendLine($"  {parts} part{(parts == 1 ? "" : "s")} cut, {yield:0.#}% yield");

        // Long: list the usable end drops. (Flat waste is an area, conveyed by the yield %.)
        if (isLong)
        {
            var drops = layouts.Where(l => l.Drop > 1e-6).Select(l => DrawFormat.FracInch(l.Drop)).ToList();
            if (drops.Count > 0)
                sb.AppendLine($"  Drops: {string.Join(", ", drops)}");
        }
        sb.AppendLine();
    }

    /// <summary>"@ 240"" for long (no width) or "48" x 96"" for a flat sheet.</summary>
    private static string SizeLabel(double? width, double length) =>
        width is double w && w > 0
            ? $"{DrawFormat.FracInch(w)} x {DrawFormat.FracInch(length)}"
            : $"@ {DrawFormat.FracInch(length)}";

    private static void BuildUsageAndPurchases(OptimizeResult result)
    {
        foreach (var g in result.Layouts
                     .GroupBy(l => (l.Material, l.Gauge, l.StockLength, l.StockWidth))
                     .OrderBy(g => g.Key.Material).ThenByDescending(g => g.Key.StockLength))
        {
            int count = g.Count();
            double useful = g.Sum(l => l.Pieces.Sum(p => p.Length));
            double total = g.Sum(l => l.StockLength);
            result.Usage.Add(new StockUsage
            {
                Material = g.Key.Material,
                Gauge = g.Key.Gauge,
                Length = g.Key.StockLength,
                Width = g.Key.StockWidth,
                Count = count,
                YieldPct = total > 0 ? 100.0 * useful / total : 0,
            });

            int purchased = g.Count(l => l.Purchased);
            if (purchased > 0)
                result.ToPurchase.Add(new Purchase
                {
                    Material = g.Key.Material,
                    Gauge = g.Key.Gauge,
                    Length = g.Key.StockLength,
                    Width = g.Key.StockWidth,
                    Count = purchased,
                });
        }
    }

    private static void AppendGlobalSummary(StringBuilder sb, OptimizeResult result)
    {
        if (result.ToPurchase.Count > 0)
        {
            sb.AppendLine("To purchase (on-hand can't finish the job):");
            foreach (var p in result.ToPurchase)
            {
                string lbl = string.IsNullOrWhiteSpace(p.Material) && string.IsNullOrWhiteSpace(p.Gauge)
                    ? "stock" : GroupLabel(p.Material, p.Gauge);
                sb.AppendLine($"  {p.Count} x {lbl} {SizeLabel(p.Width, p.Length)}");
            }
            sb.AppendLine();
        }
        if (result.Issues.Count > 0)
        {
            sb.AppendLine("Issues:");
            foreach (var i in result.Issues)
                sb.AppendLine($"  - {i.Message}");
        }
    }

    private static string Key(string? material, string? gauge) =>
        $"{(material ?? "").Trim().ToLowerInvariant()}|{(gauge ?? "").Trim().ToLowerInvariant()}";

    private static string GroupLabel(string material, string gauge) =>
        string.IsNullOrWhiteSpace(gauge) ? (string.IsNullOrWhiteSpace(material) ? "(unspecified)" : material)
                                         : $"{material} {gauge}";
}
