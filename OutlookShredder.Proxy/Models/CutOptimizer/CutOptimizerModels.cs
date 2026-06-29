using System.Globalization;

namespace OutlookShredder.Proxy.Models.CutOptimizer;

/// <summary>Which packing algorithm a job uses: 2D sheet nesting vs 1D length cutting.</summary>
public enum MaterialForm { Flat, Long }

/// <summary>Flat-sheet cut method — drives the 2D placement (guillotine vs free nest). Long is always a saw/shear.</summary>
public enum CutMethod { Shear, Laser }

/// <summary>One part that must be cut: a material/gauge, size(s), and a quantity.</summary>
public sealed class CutPart
{
    public string? Label { get; set; }
    public string Material { get; set; } = "";
    public string Gauge { get; set; } = "";
    public double? Width { get; set; }   // inches — FLAT only
    public double Length { get; set; }   // inches
    public int Qty { get; set; } = 1;
    /// <summary>FLAT only: when true the part keeps its orientation (grain/finish-fixed) and is never rotated.</summary>
    public bool FinishDirectionLocked { get; set; }
}

/// <summary>One available stock size. <see cref="QtyAvailable"/> null ⇒ unlimited (buy as needed).</summary>
public sealed class StockSize
{
    public string Material { get; set; } = "";
    public string Gauge { get; set; } = "";
    public double? Width { get; set; }    // inches — FLAT only
    public double Length { get; set; }    // inches
    public int? QtyAvailable { get; set; } // null = unlimited; a number = on-hand cap
}

/// <summary>
/// Wire + service request. <see cref="Form"/>/<see cref="Method"/> are strings on the wire so the
/// client never depends on enum integer values; <see cref="ResolvedForm"/>/<see cref="ResolvedMethod"/>
/// parse them.
/// </summary>
public sealed class OptimizeRequest
{
    public string? Form { get; set; }      // "flat" | "long"
    public string? Method { get; set; }    // "shear" | "laser" (flat only)
    public bool PrecisionNeeded { get; set; }
    /// <summary>Long only: a soft preference to keep each leftover either ~0 or ≥ this many inches
    /// (a reusable remnant) rather than short scrap. 0 = off. Ignored for flat sheets.</summary>
    public double MinUsableDrop { get; set; }
    /// <summary>Flat only: a soft preference to leave a reusable offcut rectangle ≥ this many inches on
    /// BOTH sides (parts packed to a corner so the offcut stays contiguous). 0 = off. Ignored for long.</summary>
    public double MinUsableOffcut { get; set; }
    public List<CutPart> Parts { get; set; } = new();
    public List<StockSize> Stock { get; set; } = new();
    /// <summary>The requesting operator's Shredder username, stamped as "Created by" on the layout PDF.
    /// Null/empty ⇒ the proxy falls back to its own Windows account.</summary>
    public string? CreatedBy { get; set; }

    public MaterialForm ResolvedForm =>
        string.Equals(Form, "long", StringComparison.OrdinalIgnoreCase) ? MaterialForm.Long : MaterialForm.Flat;

    public CutMethod ResolvedMethod =>
        string.Equals(Method, "laser", StringComparison.OrdinalIgnoreCase) ? CutMethod.Laser : CutMethod.Shear;
}

/// <summary>A single piece placed on a stock bar/sheet. Long uses <see cref="Length"/>; flat uses the rect.</summary>
public sealed class PlacedPiece
{
    public string? Label { get; set; }
    public double Length { get; set; }   // long: cut length
    // Flat placement (Phase 2):
    public double X { get; set; }
    public double Y { get; set; }
    public double W { get; set; }
    public double L { get; set; }
    public bool Rotated { get; set; }
}

/// <summary>One stock bar (long) or sheet (flat) with its placed pieces and leftover.</summary>
public sealed class Layout
{
    public string Material { get; set; } = "";
    public string Gauge { get; set; } = "";
    public double StockLength { get; set; }
    public double? StockWidth { get; set; }     // flat
    public List<PlacedPiece> Pieces { get; set; } = new();
    public double Drop { get; set; }            // long: end drop (in); flat: waste area (sq in)
    public double YieldPct { get; set; }
    public bool Purchased { get; set; }         // true = this bar/sheet was bought (on-hand exhausted)
    // Flat reusable offcut rectangle (in), when a usable-offcut minimum is set; W/L = 0 ⇒ none.
    public double OffcutX { get; set; }
    public double OffcutY { get; set; }
    public double OffcutW { get; set; }
    public double OffcutL { get; set; }
}

/// <summary>Per distinct stock size: how many used and the average yield.</summary>
public sealed class StockUsage
{
    public string Material { get; set; } = "";
    public string Gauge { get; set; } = "";
    public double Length { get; set; }
    public double? Width { get; set; }
    public int Count { get; set; }
    public double YieldPct { get; set; }
}

/// <summary>Stock to buy when on-hand can't finish the job.</summary>
public sealed class Purchase
{
    public string Material { get; set; } = "";
    public string Gauge { get; set; } = "";
    public double Length { get; set; }
    public double? Width { get; set; }
    public int Count { get; set; }
}

/// <summary>A problem that did not stop the run (un-cuttable part, no matching stock, …).</summary>
public sealed class Issue
{
    public string Message { get; set; } = "";
}

/// <summary>The full optimizer result: text summary + structured plan + (Phase 3) a PDF report.</summary>
public sealed class OptimizeResult
{
    public string TextSummary { get; set; } = "";
    public string? PdfBase64 { get; set; }
    public List<StockUsage> Usage { get; set; } = new();
    public List<Purchase> ToPurchase { get; set; } = new();
    public List<Layout> Layouts { get; set; } = new();
    public List<Issue> Issues { get; set; } = new();
}
