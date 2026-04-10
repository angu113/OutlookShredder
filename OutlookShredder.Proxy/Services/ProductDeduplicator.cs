using OutlookShredder.Proxy.Models;

namespace OutlookShredder.Proxy.Services;

/// <summary>
/// Pre-insert deduplication for extracted product lines.
///
/// Claude sometimes extracts the same product multiple times when a quote PDF
/// describes pricing in more than one place (e.g. a line-item table and a summary
/// section). This class collapses those duplicates before the SP write loop runs.
///
/// Identity: two rows are considered the same product when they share the same
/// TotalPrice (within 1 cent) AND the same UnitsQuoted. Rows where either field
/// is null are passed through unchanged — not enough signal to safely deduplicate.
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

        // Separate deduplicatable rows (both key fields present) from pass-throughs.
        var keyed     = new List<(int Index, ProductLine P, string Key)>();
        var passThrough = new List<(int Index, ProductLine P)>();

        for (int i = 0; i < products.Count; i++)
        {
            var p = products[i];
            if (p.TotalPrice is null || p.UnitsQuoted is null)
            {
                passThrough.Add((i, p));
                log.LogInformation(
                    "[Dedup] Row {Idx} '{Name}' — no identity key (TotalPrice={TP} UnitsQuoted={UQ}), passing through",
                    i, p.ProductName ?? "(unnamed)", p.TotalPrice, p.UnitsQuoted);
            }
            else
            {
                // Bucket key: price rounded to nearest cent + units
                var key = $"{Math.Round(p.TotalPrice.Value, 2):F2}|{p.UnitsQuoted.Value}";
                keyed.Add((i, p, key));
            }
        }

        // Group keyed rows by identity key.
        var groups = keyed.GroupBy(x => x.Key).ToList();

        var winners = new List<(int OriginalIndex, ProductLine P)>();

        foreach (var group in groups)
        {
            var members = group.ToList();

            if (members.Count == 1)
            {
                var solo = members[0];
                log.LogInformation(
                    "[Dedup] Row {Idx} '{Name}' — unique (TotalPrice={TP}, UnitsQuoted={UQ}), kept",
                    solo.Index, solo.P.ProductName ?? "(unnamed)",
                    solo.P.TotalPrice, solo.P.UnitsQuoted);
                winners.Add((solo.Index, solo.P));
                continue;
            }

            // Score each member and pick the richest.
            var scored = members.Select(m => (m.Index, m.P, Score: Richness(m.P))).ToList();
            scored.Sort((a, b) =>
            {
                int c = b.Score.CompareTo(a.Score);    // descending score
                return c != 0 ? c : a.Index.CompareTo(b.Index); // tie → earliest
            });

            var winner = scored[0];
            var dropped = scored.Skip(1).ToList();

            var prefix = dryRun ? "[Dedup DRY]" : "[Dedup]";

            log.LogInformation(
                "{Prefix} DUPLICATE GROUP — key=(TotalPrice={TP}, UnitsQuoted={UQ}), {Total} rows:",
                prefix, winner.P.TotalPrice, winner.P.UnitsQuoted, members.Count);

            log.LogInformation(
                "{Prefix}   KEEP  Row {Idx} '{Name}' score={Score} | Len={LenP}/{LenC} | Fields: {Fields}",
                prefix, winner.Index, winner.P.ProductName ?? "(unnamed)", winner.Score,
                winner.P.ProductName?.Length ?? 0, winner.P.SupplierProductComments?.Length ?? 0,
                FieldSummary(winner.P));

            foreach (var d in dropped)
            {
                log.LogInformation(
                    "{Prefix}   DROP  Row {Idx} '{Name}' score={Score} | Len={LenP}/{LenC} | Fields: {Fields}",
                    prefix, d.Index, d.P.ProductName ?? "(unnamed)", d.Score,
                    d.P.ProductName?.Length ?? 0, d.P.SupplierProductComments?.Length ?? 0,
                    FieldSummary(d.P));
            }

            // In dry-run mode keep all members but still log the decision.
            if (dryRun)
                winners.AddRange(members.Select(m => (m.Index, m.P)));
            else
                winners.Add((winner.Index, winner.P));
        }

        // Merge winners with pass-throughs, preserving original order.
        var combined = winners.Select(x => (x.OriginalIndex, x.P))
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

    // ── Richness scoring ──────────────────────────────────────────────────────

    /// <summary>
    /// Higher = more informative row. Used to pick the winner in a duplicate group.
    /// </summary>
    private static int Richness(ProductLine p)
    {
        int score = 0;

        // Scalar numeric fields — each contributes 1 point when present.
        if (p.UnitsRequested  is not null) score += 1;
        if (p.UnitsQuoted     is not null) score += 1;
        if (p.LengthPerUnit   is not null) score += 1;
        if (p.WeightPerUnit   is not null) score += 1;
        if (p.PricePerPound   is not null) score += 1;
        if (p.PricePerFoot    is not null) score += 1;
        if (p.PricePerPiece   is not null) score += 1;
        if (p.TotalPrice      is not null) score += 1;

        // Unit strings — 1 each.
        if (!string.IsNullOrWhiteSpace(p.LengthUnit)) score += 1;
        if (!string.IsNullOrWhiteSpace(p.WeightUnit)) score += 1;

        // Text fields with length bonus — more descriptive text = higher score.
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
