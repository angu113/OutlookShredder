using OutlookShredder.Proxy.Models;

namespace OutlookShredder.Proxy.Services;

/// <summary>
/// Pre-insert deduplication for extracted product lines.
///
/// The AI sometimes extracts the same product multiple times when a quote PDF
/// describes pricing in more than one place (e.g. a line-item table and a summary
/// section). This class collapses those duplicates before the SP write loop runs.
///
/// Identity (two passes):
///   Pass 1 — MSPC: rows sharing the same non-null ProductSearchKey are the same
///             catalog product. Unmatched rows (null MSPC) proceed to pass 2.
///   Pass 2 — Price+Units: rows with the same TotalPrice (±1 cent) AND UnitsQuoted
///             are treated as the same product line. Rows where either field is null
///             are passed through unchanged.
///
/// Winner selection: the row with the highest richness score (weighted non-null
/// field count + text-length bonus) survives. Ties go to the first occurrence.
///
/// Dry-run mode (Dedup:DryRun = true): all rows are returned unchanged; decisions
/// are logged with a [Dedup DRY] prefix so you can audit without data loss.
/// </summary>
internal static class ProductDeduplicator
{
    private const double PriceTolerance = 0.01;

    /// <summary>
    /// Returns a deduplicated list. The input list is never mutated.
    /// Logs one line per group (whether collapsed or not) so every decision
    /// is visible in the log for post-hoc analysis.
    /// </summary>
    public static List<ProductLine> Deduplicate(
        IReadOnlyList<ProductLine> products,
        string source,
        bool dryRun,
        ILogger log)
    {
        if (products.Count <= 1)
        {
            if (products.Count == 1)
                log.LogInformation("[Dedup] 1 product from {Source} — no dedup needed", source);
            return products.ToList();
        }

        log.LogInformation("[Dedup] Evaluating {Count} product(s) from {Source} (dryRun={DryRun})",
            products.Count, source, dryRun);

        var prefix = dryRun ? "[Dedup DRY]" : "[Dedup]";

        // ── Pass 1: MSPC grouping ─────────────────────────────────────────────
        // Rows sharing the same non-null ProductSearchKey are the same catalog
        // product regardless of whether pricing fields are present. This handles
        // PDFs where different structural sections (description block, pricing
        // grid, comparison note) each produce an extraction for the same item.
        var mspcGroups = new Dictionary<string, List<(int Index, ProductLine P)>>(
            StringComparer.OrdinalIgnoreCase);
        var noMspc = new List<(int Index, ProductLine P)>();

        for (int i = 0; i < products.Count; i++)
        {
            var p = products[i];
            if (!string.IsNullOrWhiteSpace(p.ProductSearchKey))
            {
                if (!mspcGroups.TryGetValue(p.ProductSearchKey!, out var bucket))
                    mspcGroups[p.ProductSearchKey!] = bucket = [];
                bucket.Add((i, p));
            }
            else
                noMspc.Add((i, p));
        }

        var pass1Winners = new List<(int OriginalIndex, ProductLine P)>();

        foreach (var (mspc, members) in mspcGroups)
        {
            if (members.Count == 1)
            {
                var solo = members[0];
                log.LogInformation(
                    "[Dedup] Row {Idx} '{Name}' MSPC={Mspc} — unique, kept",
                    solo.Index, solo.P.ProductName ?? "(unnamed)", mspc);
                pass1Winners.Add((solo.Index, solo.P));
            }
            else
            {
                var winner = PickWinner(members, prefix, $"MSPC={mspc}", log, dryRun);
                if (dryRun)
                    pass1Winners.AddRange(members.Select(m => (m.Index, m.P)));
                else
                    pass1Winners.Add(winner);
            }
        }

        // ── Pass 2: Price+Units grouping for rows without an MSPC ────────────
        var keyed       = new List<(int Index, ProductLine P, string Key)>();
        var passThrough = new List<(int Index, ProductLine P)>();

        foreach (var (i, p) in noMspc)
        {
            if (p.TotalPrice is null || p.UnitsQuoted is null)
            {
                passThrough.Add((i, p));
                log.LogInformation(
                    "[Dedup] Row {Idx} '{Name}' — no MSPC, no price key (TotalPrice={TP} UnitsQuoted={UQ}), passing through",
                    i, p.ProductName ?? "(unnamed)", p.TotalPrice, p.UnitsQuoted);
            }
            else
            {
                var key = $"{Math.Round(p.TotalPrice.Value, 2):F2}|{p.UnitsQuoted.Value}";
                keyed.Add((i, p, key));
            }
        }

        var pass2Winners = new List<(int OriginalIndex, ProductLine P)>();

        foreach (var group in keyed.GroupBy(x => x.Key))
        {
            var members = group.Select(x => (x.Index, x.P)).ToList();
            if (members.Count == 1)
            {
                var solo = members[0];
                log.LogInformation(
                    "[Dedup] Row {Idx} '{Name}' — unique (TotalPrice={TP}, UnitsQuoted={UQ}), kept",
                    solo.Index, solo.P.ProductName ?? "(unnamed)",
                    solo.P.TotalPrice, solo.P.UnitsQuoted);
                pass2Winners.Add((solo.Index, solo.P));
            }
            else
            {
                var key = group.Key;
                var winner = PickWinner(members, prefix, $"Price+Units={key}", log, dryRun);
                if (dryRun)
                    pass2Winners.AddRange(members.Select(m => (m.Index, m.P)));
                else
                    pass2Winners.Add(winner);
            }
        }

        // ── Merge all survivors, preserving original order ────────────────────
        var combined = pass1Winners
            .Concat(pass2Winners)
            .Concat(passThrough.Select(x => (OriginalIndex: x.Index, x.P)))
            .OrderBy(x => x.OriginalIndex)
            .Select(x => x.P)
            .ToList();

        int removed = products.Count - combined.Count;
        log.LogInformation("[Dedup] Result: {In} in → {Out} out ({Removed} duplicate(s) {Action})",
            products.Count, combined.Count, removed,
            dryRun ? "flagged (dry-run, not removed)" : "removed");

        return combined;
    }

    // ── Helpers ───────────────────────────────────────────────────────────────

    private static (int OriginalIndex, ProductLine P) PickWinner(
        List<(int Index, ProductLine P)> members,
        string prefix,
        string groupDesc,
        ILogger log,
        bool dryRun)
    {
        var scored = members
            .Select(m => (m.Index, m.P, Score: Richness(m.P)))
            .OrderByDescending(x => x.Score)
            .ThenBy(x => x.Index)
            .ToList();

        var winner  = scored[0];
        var dropped = scored.Skip(1).ToList();

        log.LogInformation(
            "{Prefix} DUPLICATE GROUP — {Desc}, {Total} rows:",
            prefix, groupDesc, members.Count);
        log.LogInformation(
            "{Prefix}   KEEP  Row {Idx} '{Name}' score={Score} | Len={LenP}/{LenC} | Fields: {Fields}",
            prefix, winner.Index, winner.P.ProductName ?? "(unnamed)", winner.Score,
            winner.P.ProductName?.Length ?? 0, winner.P.SupplierProductComments?.Length ?? 0,
            FieldSummary(winner.P));
        foreach (var d in dropped)
            log.LogInformation(
                "{Prefix}   DROP  Row {Idx} '{Name}' score={Score} | Len={LenP}/{LenC} | Fields: {Fields}",
                prefix, d.Index, d.P.ProductName ?? "(unnamed)", d.Score,
                d.P.ProductName?.Length ?? 0, d.P.SupplierProductComments?.Length ?? 0,
                FieldSummary(d.P));

        return (winner.Index, winner.P);
    }

    // ── Richness scoring ──────────────────────────────────────────────────────

    private static int Richness(ProductLine p)
    {
        int score = 0;

        if (p.UnitsRequested  is not null) score += 1;
        if (p.UnitsQuoted     is not null) score += 1;
        if (p.LengthPerUnit   is not null) score += 1;
        if (p.WeightPerUnit   is not null) score += 1;
        if (p.PricePerPound   is not null) score += 1;
        if (p.PricePerFoot    is not null) score += 1;
        if (p.PricePerPiece   is not null) score += 1;
        if (p.TotalPrice      is not null) score += 1;

        if (!string.IsNullOrWhiteSpace(p.LengthUnit)) score += 1;
        if (!string.IsNullOrWhiteSpace(p.WeightUnit)) score += 1;

        score += TextScore(p.ProductName,             basePoints: 2, bonusPer: 20);
        score += TextScore(p.LeadTimeText,            basePoints: 1, bonusPer: 30);
        score += TextScore(p.Certifications,          basePoints: 1, bonusPer: 40);
        score += TextScore(p.SupplierProductComments, basePoints: 1, bonusPer: 50);

        return score;
    }

    private static int TextScore(string? value, int basePoints, int bonusPer)
    {
        if (string.IsNullOrWhiteSpace(value)) return 0;
        return basePoints + value.Length / bonusPer;
    }

    // ── Logging helpers ───────────────────────────────────────────────────────

    private static string FieldSummary(ProductLine p)
    {
        var present = new List<string>();
        if (p.UnitsRequested  is not null)                        present.Add($"UnitsReq={p.UnitsRequested}");
        if (p.UnitsQuoted     is not null)                        present.Add($"UnitsQ={p.UnitsQuoted}");
        if (p.LengthPerUnit   is not null)                        present.Add($"Len={p.LengthPerUnit}{p.LengthUnit}");
        if (p.WeightPerUnit   is not null)                        present.Add($"Wt={p.WeightPerUnit}{p.WeightUnit}");
        if (p.PricePerPound   is not null)                        present.Add($"$/lb={p.PricePerPound}");
        if (p.PricePerFoot    is not null)                        present.Add($"$/ft={p.PricePerFoot}");
        if (p.PricePerPiece   is not null)                        present.Add($"$/pc={p.PricePerPiece}");
        if (p.TotalPrice      is not null)                        present.Add($"Total={p.TotalPrice}");
        if (!string.IsNullOrWhiteSpace(p.LeadTimeText))           present.Add("LeadTime");
        if (!string.IsNullOrWhiteSpace(p.Certifications))         present.Add("Certs");
        if (!string.IsNullOrWhiteSpace(p.SupplierProductComments))present.Add("Comments");
        return present.Count > 0 ? string.Join(", ", present) : "(none)";
    }
}
