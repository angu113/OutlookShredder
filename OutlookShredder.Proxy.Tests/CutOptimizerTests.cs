using System.Collections.Generic;
using System.Linq;
using OutlookShredder.Proxy.Models.CutOptimizer;
using OutlookShredder.Proxy.Services.CutOptimizer;
using Xunit;

namespace OutlookShredder.Proxy.Tests;

// Cut optimizer — LONG (1D) path. Behaviour is pinned, not internals: the kerf model (1/8" per cut
// when precision is on), the on-hand-vs-purchase plan, per-(material,gauge) grouping, and issues.
// Changing any of this behaviour must update this test in the same commit.
public class CutOptimizerTests
{
    private static CutPart Part(double len, int qty, string mat = "CRS", string ga = "11 ga") =>
        new() { Material = mat, Gauge = ga, Length = len, Qty = qty };

    private static StockSize Stock(double len, int? qty = null, string mat = "CRS", string ga = "11 ga") =>
        new() { Material = mat, Gauge = ga, Length = len, QtyAvailable = qty };

    private static OptimizeRequest Long(bool precision, IEnumerable<CutPart> parts, IEnumerable<StockSize> stock) =>
        new() { Form = "long", PrecisionNeeded = precision, Parts = parts.ToList(), Stock = stock.ToList() };

    // ── The pinned kerf example (wip/feat-cut-optimizer.md §4 + §11) ──────────────
    // 240" stock, 60" part x4, precision ON -> a full bar holds 3 @ 60" with 1/8"/cut, drop 59-5/8".
    [Fact]
    public void Long_precision_on_consumes_kerf_per_cut()
    {
        var r = CutOptimizerService.Optimize(Long(true, new[] { Part(60, 4) }, new[] { Stock(240) }));

        // The first (full) bar fits exactly 3 pieces; consumed 3x60 + 3x0.125 = 180.375 -> drop 59.625.
        var full = r.Layouts.Single(l => l.Pieces.Count == 3);
        Assert.Equal(59.625, full.Drop, 3);
        Assert.Equal(2, r.Layouts.Count);                 // 4 pieces need a 2nd bar (holds the 4th)
        Assert.Equal(4, r.Layouts.Sum(l => l.Pieces.Count));
        Assert.Empty(r.Issues);
    }

    [Fact]
    public void Long_precision_off_packs_with_no_kerf()
    {
        var r = CutOptimizerService.Optimize(Long(false, new[] { Part(60, 4) }, new[] { Stock(240) }));

        // 4 x 60 = 240 fills exactly one bar, zero drop.
        var bar = Assert.Single(r.Layouts);
        Assert.Equal(4, bar.Pieces.Count);
        Assert.Equal(0, bar.Drop, 3);
        Assert.Equal(100.0, bar.YieldPct, 1);
    }

    // ── Min usable drop (soft preference, long only) ─────────────────────────────
    [Fact]
    public void Long_min_usable_drop_keeps_drops_reusable_not_scrap()
    {
        // 3 x 100" from 240" stock. Tight packing leaves a 40" scrap on one bar; with a 96" usable-drop
        // minimum the optimizer instead leaves reusable >=96" remnants (more bars, but no short scrap).
        var withMin = CutOptimizerService.Optimize(new OptimizeRequest
        {
            Form = "long", MinUsableDrop = 96,
            Parts = new() { Part(100, 3) }, Stock = new() { Stock(240) },
        });
        Assert.DoesNotContain(withMin.Layouts, l => l.Drop > 1e-6 && l.Drop + 1e-6 < 96);   // no short scrap
        Assert.Contains(withMin.Layouts, l => l.Drop + 1e-6 >= 96);                          // a reusable remnant
        Assert.Contains("Reusable drops", withMin.TextSummary);

        var noMin = CutOptimizerService.Optimize(new OptimizeRequest
        {
            Form = "long",
            Parts = new() { Part(100, 3) }, Stock = new() { Stock(240) },
        });
        Assert.Contains(noMin.Layouts, l => l.Drop > 1e-6 && l.Drop + 1e-6 < 96);            // tight pack leaves scrap
    }

    // ── Rendering: identical cuts collapse to one pro-forma + count ───────────────
    [Fact]
    public void Render_collapses_identical_cuts()
    {
        var a = new Layout { StockLength = 240, Pieces = new() { new PlacedPiece { Length = 60 }, new PlacedPiece { Length = 60 } }, Drop = 120 };
        var b = new Layout { StockLength = 240, Pieces = new() { new PlacedPiece { Length = 60 }, new PlacedPiece { Length = 60 } }, Drop = 120 };
        var c = new Layout { StockLength = 120, Pieces = new() { new PlacedPiece { Length = 120 } }, Drop = 0 };

        var groups = CutLayoutPdfRenderer.GroupIdenticalCuts(new List<Layout> { a, b, c }, flat: false);
        Assert.Equal(2, groups.Count);
        Assert.Equal(2, groups[0].Count);   // a + b are the same cut
        Assert.Equal(1, groups[1].Count);   // c is distinct
    }

    // ── On-hand vs purchase (Decision 4) ─────────────────────────────────────────
    [Fact]
    public void Long_unlimited_stock_reports_lengths_needed_no_purchase()
    {
        // 10 lengths needed @ 240", stock unlimited -> "need 10", nothing to purchase.
        var r = CutOptimizerService.Optimize(Long(false, new[] { Part(240, 10) }, new[] { Stock(240) }));
        Assert.Equal(10, r.Layouts.Count);
        Assert.Empty(r.ToPurchase);
        Assert.Contains("Need 10 lengths @ 240\"", r.TextSummary);
    }

    [Fact]
    public void Long_capped_stock_surfaces_a_purchase_shortfall()
    {
        // Need 5 bars' worth but only 3 on hand -> cut 3 on-hand + purchase 2 to finish.
        var r = CutOptimizerService.Optimize(Long(false, new[] { Part(240, 5) }, new[] { Stock(240, qty: 3) }));
        Assert.Equal(5, r.Layouts.Count);
        Assert.Equal(3, r.Layouts.Count(l => !l.Purchased));
        var buy = Assert.Single(r.ToPurchase);
        Assert.Equal(2, buy.Count);
        Assert.Equal(240, buy.Length, 3);
    }

    [Fact]
    public void Long_with_no_material_reads_cleanly()
    {
        // The client hides metal/gauge for long, so material/gauge arrive empty — summary must not say "(unspecified)".
        var r = CutOptimizerService.Optimize(Long(false,
            new[] { new CutPart { Material = "", Gauge = "", Length = 60, Qty = 2 } },
            new[] { new StockSize { Material = "", Gauge = "", Length = 240 } }));
        Assert.Contains("Long stock", r.TextSummary);
        Assert.DoesNotContain("unspecified", r.TextSummary);
    }

    // ── Grouping + issues ────────────────────────────────────────────────────────
    [Fact]
    public void Parts_are_grouped_by_material_and_gauge()
    {
        var r = CutOptimizerService.Optimize(Long(false,
            new[] { Part(60, 2, "CRS", "11 ga"), Part(60, 2, "SS", "16 ga") },
            new[] { Stock(240, mat: "CRS", ga: "11 ga"), Stock(240, mat: "SS", ga: "16 ga") }));

        Assert.Empty(r.Issues);
        Assert.Contains(r.Layouts, l => l.Material == "CRS");
        Assert.Contains(r.Layouts, l => l.Material == "SS");
        // No cross-material cutting: each group's stock matches its parts' material.
        Assert.All(r.Layouts, l => Assert.Contains(r.Usage, u => u.Material == l.Material && u.Gauge == l.Gauge));
    }

    [Fact]
    public void No_matching_stock_is_an_issue_not_a_silent_drop()
    {
        var r = CutOptimizerService.Optimize(Long(false,
            new[] { Part(60, 2, "CRS", "11 ga") },
            new[] { Stock(240, mat: "SS", ga: "16 ga") }));   // wrong material — no CRS stock
        Assert.Empty(r.Layouts);
        Assert.Contains(r.Issues, i => i.Message.Contains("No stock for CRS 11 ga"));
    }

    [Fact]
    public void A_part_longer_than_any_stock_is_flagged_uncuttable()
    {
        var r = CutOptimizerService.Optimize(Long(false, new[] { Part(300, 1) }, new[] { Stock(240) }));
        Assert.Empty(r.Layouts);
        Assert.Contains(r.Issues, i => i.Message.Contains("exceeds the largest stock"));
    }

    // ── FLAT (2D) path ───────────────────────────────────────────────────────────
    private static CutPart FlatPart(double w, double l, int qty, bool lockDir = false, string mat = "CRS", string ga = "11 ga") =>
        new() { Material = mat, Gauge = ga, Width = w, Length = l, Qty = qty, FinishDirectionLocked = lockDir };

    private static StockSize FlatStock(double w, double l, int? qty = null, string mat = "CRS", string ga = "11 ga") =>
        new() { Material = mat, Gauge = ga, Width = w, Length = l, QtyAvailable = qty };

    private static OptimizeRequest Flat(string method, IEnumerable<CutPart> parts, IEnumerable<StockSize> stock) =>
        new() { Form = "flat", Method = method, Parts = parts.ToList(), Stock = stock.ToList() };

    [Fact]
    public void Flat_shear_tiles_a_sheet_exactly()
    {
        // 4 x (24 x 48) tile a 48 x 96 sheet exactly -> one sheet, ~100% yield, no kerf (shear).
        var r = CutOptimizerService.Optimize(Flat("shear", new[] { FlatPart(24, 48, 4) }, new[] { FlatStock(48, 96) }));
        Assert.Empty(r.Issues);
        var sheet = Assert.Single(r.Layouts);
        Assert.Equal(4, sheet.Pieces.Count);
        Assert.True(sheet.YieldPct > 99.0, $"expected ~100% yield, got {sheet.YieldPct:0.#}%");
    }

    [Fact]
    public void Flat_parts_too_big_to_share_a_sheet_use_one_sheet_each()
    {
        // 40 x 50 can't pair on a 48 x 96 sheet (neither stacks nor rotates in) -> one per sheet.
        var r = CutOptimizerService.Optimize(Flat("shear", new[] { FlatPart(40, 50, 2) }, new[] { FlatStock(48, 96) }));
        Assert.Empty(r.Issues);
        Assert.Equal(2, r.Layouts.Count);
        Assert.All(r.Layouts, l => Assert.Single(l.Pieces));
    }

    [Fact]
    public void Flat_finish_lock_prevents_the_rotation_that_would_fit()
    {
        // 8 x 4 only fits a 6-wide sheet rotated. Unlocked -> rotates in; locked -> can't be cut.
        var unlocked = CutOptimizerService.Optimize(Flat("shear", new[] { FlatPart(8, 4, 1, lockDir: false) }, new[] { FlatStock(6, 100) }));
        var placed = Assert.Single(unlocked.Layouts).Pieces.Single();
        Assert.True(placed.Rotated);

        var locked = CutOptimizerService.Optimize(Flat("shear", new[] { FlatPart(8, 4, 1, lockDir: true) }, new[] { FlatStock(6, 100) }));
        Assert.Empty(locked.Layouts);
        Assert.Contains(locked.Issues, i => i.Message.Contains("fits no stock sheet"));
    }

    [Fact]
    public void Flat_laser_web_costs_space_vs_shear_on_a_tight_tile()
    {
        // The 0.100" laser web breaks an exact 2x2 tile: shear fits 4 on one sheet, laser needs another.
        var shear = CutOptimizerService.Optimize(Flat("shear", new[] { FlatPart(24, 48, 4) }, new[] { FlatStock(48, 96) }));
        var laser = CutOptimizerService.Optimize(Flat("laser", new[] { FlatPart(24, 48, 4) }, new[] { FlatStock(48, 96) }));
        Assert.Single(shear.Layouts);
        Assert.True(laser.Layouts.Count >= 2, $"laser web should force a 2nd sheet, got {laser.Layouts.Count}");
        Assert.Equal(4, laser.Layouts.Sum(l => l.Pieces.Count));   // still cuts all 4 parts
    }

    // ── Flat reusable offcut (soft preference, flat only) ────────────────────────
    [Fact]
    public void Flat_min_offcut_leaves_a_reusable_remnant_rectangle()
    {
        // One 48x48 part on a 96x48 sheet leaves a clean 48x48 offcut; with a 24" offcut minimum it's
        // reported/kept as reusable inventory rather than waste.
        var r = CutOptimizerService.Optimize(new OptimizeRequest
        {
            Form = "flat", Method = "shear", MinUsableOffcut = 24,
            Parts = new() { FlatPart(48, 48, 1) }, Stock = new() { FlatStock(96, 48) },
        });
        var sheet = Assert.Single(r.Layouts);
        Assert.True(sheet.OffcutW >= 24 && sheet.OffcutL >= 24, $"expected a reusable offcut, got {sheet.OffcutW} x {sheet.OffcutL}");
        Assert.Contains("Reusable offcuts", r.TextSummary);

        // Without the offcut minimum, no offcut is tracked.
        var off = CutOptimizerService.Optimize(Flat("shear", new[] { FlatPart(48, 48, 1) }, new[] { FlatStock(96, 48) }));
        Assert.Equal(0, off.Layouts[0].OffcutW, 3);
    }

    // ── PDF report (Phase 3) ───────────────────────────────────────────────────────
    [Fact]
    public void Long_plan_renders_a_pdf_report()
    {
        var r = CutOptimizerService.Optimize(Long(true, new[] { Part(60, 5), Part(40, 3) }, new[] { Stock(240, qty: 2) }));
        AssertRealPdf(r.PdfBase64);
    }

    [Fact]
    public void Flat_plan_renders_a_pdf_report()
    {
        var r = CutOptimizerService.Optimize(Flat("shear", new[] { FlatPart(24, 48, 6) }, new[] { FlatStock(48, 96) }));
        AssertRealPdf(r.PdfBase64);
    }

    [Fact]
    public void No_plan_means_no_pdf()
    {
        var r = CutOptimizerService.Optimize(Long(false, new[] { Part(300, 1) }, new[] { Stock(240) }));   // uncuttable
        Assert.Null(r.PdfBase64);
    }

    private static void AssertRealPdf(string? base64)
    {
        Assert.False(string.IsNullOrEmpty(base64));
        var bytes = System.Convert.FromBase64String(base64!);
        Assert.True(bytes.Length > 1000, $"expected a real PDF, got {bytes.Length} bytes");
        Assert.Equal((byte)'%', bytes[0]);   // %PDF header
        Assert.Equal((byte)'P', bytes[1]);
    }
}
