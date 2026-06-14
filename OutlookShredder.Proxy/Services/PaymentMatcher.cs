using System.Globalization;
using System.Security.Cryptography;
using System.Text;
using OutlookShredder.Proxy.Models;

namespace OutlookShredder.Proxy.Services;

public sealed class ReconMatchOptions
{
    public decimal AmountTolerance   { get; set; } = 0.01m;
    public int     DateToleranceDays { get; set; } = 1;

    // A "possible match" is a near-but-not-exact amount on a matching date — a small difference that
    // could be a keying typo. Flagged (not auto-accepted) for the user to confirm. Near = within the
    // absolute band OR the percentage band, but outside AmountTolerance.
    public decimal PossibleAmountTolerance { get; set; } = 1.00m;
    public decimal PossiblePercent         { get; set; } = 0.05m;
}

/// <summary>
/// Pure, deterministic payment reconciliation matcher (no I/O — unit-testable). Mirrors the
/// BillToPoMatcherService philosophy: strongest signal first, a tier only matches when it resolves
/// to EXACTLY ONE candidate on each side, otherwise rows fall through; nothing is auto-resolved and
/// a false match is never preferred over an Ambiguous flag. Amounts are signed (refunds match refunds).
/// </summary>
public static class PaymentMatcher
{
    private sealed class Item(PaymentTxn txn)
    {
        public PaymentTxn Txn = txn;
        public bool Consumed;
    }

    public static ReconRunResult Match(
        IReadOnlyList<PaymentTxn> ob,
        IReadOnlyList<PaymentTxn> processor,
        ReconMatchOptions opts)
    {
        var obItems   = ob.Select(t => new Item(t)).ToList();
        var procItems = processor.Select(t => new Item(t)).ToList();
        var rows = new List<ReconRow>();

        // Tier 1a/1b — exact id then exact auth-code (1:1 only; amount may differ -> AmountMismatch).
        MatchByKey(obItems, procItems, t => Norm(t.TxnId),   "txnId", opts, rows);
        MatchByKey(obItems, procItems, t => Norm(t.AuthCode), "auth", opts, rows);

        // Tier 2/3 — fuzzy amount+date(+last4); predicate already requires amount close -> Matched.
        MatchFuzzy(obItems, procItems, requireLast4: true,  "amt+date+last4", opts, rows);
        MatchFuzzy(obItems, procItems, requireLast4: false, "amt+date",       opts, rows);

        // Tier 4 — possible match: near-but-not-exact amount on a matching date (typo-explainable).
        MatchPossible(obItems, procItems, opts, rows);

        // Leftovers: bucket by exact (amount, date). A 1:1 exact bucket would already be matched by
        // tier 3, so a bucket with both sides present here is genuinely ambiguous; one-sided -> a gap.
        foreach (var grp in obItems.Where(x => !x.Consumed).Select(x => x.Txn)
                     .Concat(procItems.Where(x => !x.Consumed).Select(x => x.Txn))
                     .GroupBy(t => (Math.Round(t.Amount, 2), t.Date)))
        {
            var obSide   = grp.Where(t => t.Source == "ob").ToList();
            var procSide = grp.Where(t => t.Source == "processor").ToList();

            if (obSide.Count > 0 && procSide.Count > 0)
            {
                foreach (var t in obSide)   rows.Add(Row(ReconRowStatus.Ambiguous, t, null, note: $"{obSide.Count} OB / {procSide.Count} processor at same amount+date"));
                foreach (var t in procSide) rows.Add(Row(ReconRowStatus.Ambiguous, null, t, note: $"{obSide.Count} OB / {procSide.Count} processor at same amount+date"));
            }
            else if (obSide.Count > 0)
            {
                foreach (var t in obSide)   rows.Add(Row(ReconRowStatus.MissingInProcessor, t, null));
            }
            else
            {
                foreach (var t in procSide) rows.Add(Row(ReconRowStatus.MissingInOb, null, t));
            }
        }

        var result = new ReconRunResult
        {
            ObCount        = ob.Count,
            ProcessorCount = processor.Count,
            Rows           = rows,
            BatchTotals    = BatchTotals(ob, processor, opts),
        };
        result.Counts = Tally(rows);
        return result;
    }

    // ── tier passes ───────────────────────────────────────────────────────────

    private static void MatchByKey(
        List<Item> ob, List<Item> proc, Func<PaymentTxn, string?> key, string via,
        ReconMatchOptions opts, List<ReconRow> rows)
    {
        var obByKey = ob.Where(x => !x.Consumed && key(x.Txn) is not null)
                        .GroupBy(x => key(x.Txn)!).Where(g => g.Count() == 1)
                        .ToDictionary(g => g.Key, g => g.First());
        var procByKey = proc.Where(x => !x.Consumed && key(x.Txn) is not null)
                            .GroupBy(x => key(x.Txn)!).Where(g => g.Count() == 1)
                            .ToDictionary(g => g.Key, g => g.First());

        foreach (var (k, obItem) in obByKey)
        {
            if (!procByKey.TryGetValue(k, out var procItem)) continue;
            if (obItem.Consumed || procItem.Consumed) continue;
            obItem.Consumed = procItem.Consumed = true;
            var matched = AmountClose(obItem.Txn, procItem.Txn, opts);
            rows.Add(Row(matched ? ReconRowStatus.Matched : ReconRowStatus.AmountMismatch,
                         obItem.Txn, procItem.Txn, via));
        }
    }

    private static void MatchFuzzy(
        List<Item> ob, List<Item> proc, bool requireLast4, string via,
        ReconMatchOptions opts, List<ReconRow> rows)
    {
        bool Pred(PaymentTxn a, PaymentTxn b) =>
            AmountClose(a, b, opts) && DateClose(a, b, opts) &&
            (!requireLast4 || (a.Last4 is not null && b.Last4 is not null && a.Last4 == b.Last4));

        bool changed = true;
        while (changed)
        {
            changed = false;
            foreach (var o in ob.Where(x => !x.Consumed))
            {
                var cands = proc.Where(x => !x.Consumed && Pred(o.Txn, x.Txn)).ToList();
                if (cands.Count != 1) continue;
                var p = cands[0];
                // mutual uniqueness — p must be the unique partner for exactly one OB too
                if (ob.Count(x => !x.Consumed && Pred(x.Txn, p.Txn)) != 1) continue;
                o.Consumed = p.Consumed = true;
                rows.Add(Row(ReconRowStatus.Matched, o.Txn, p.Txn, via));
                changed = true;
            }
        }
    }

    private static void MatchPossible(
        List<Item> ob, List<Item> proc, ReconMatchOptions opts, List<ReconRow> rows)
    {
        bool Pred(PaymentTxn a, PaymentTxn b) =>
            DateClose(a, b, opts) && AmountNear(a, b, opts) && !AmountClose(a, b, opts);

        bool changed = true;
        while (changed)
        {
            changed = false;
            foreach (var o in ob.Where(x => !x.Consumed))
            {
                var cands = proc.Where(x => !x.Consumed && Pred(o.Txn, x.Txn)).ToList();
                if (cands.Count != 1) continue;
                var p = cands[0];
                if (ob.Count(x => !x.Consumed && Pred(x.Txn, p.Txn)) != 1) continue;
                o.Consumed = p.Consumed = true;
                var delta = o.Txn.Amount - p.Txn.Amount;
                rows.Add(Row(ReconRowStatus.PossibleMatch, o.Txn, p.Txn, "amt~date(possible)",
                    $"amounts differ by {delta:+0.00;-0.00} — confirm"));
                changed = true;
            }
        }
    }

    // ── batch totals ────────────────────────────────────────────────────────────

    private static List<BatchTotalRow> BatchTotals(
        IReadOnlyList<PaymentTxn> ob, IReadOnlyList<PaymentTxn> processor, ReconMatchOptions opts)
    {
        var obByDate   = ob.GroupBy(t => t.Date).ToDictionary(g => g.Key, g => g.Sum(t => t.Amount));
        var procByDate = processor.GroupBy(t => t.Date).ToDictionary(g => g.Key, g => g.Sum(t => t.Amount));

        return obByDate.Keys.Union(procByDate.Keys).OrderBy(d => d).Select(d =>
        {
            var o = obByDate.GetValueOrDefault(d);
            var p = procByDate.GetValueOrDefault(d);
            return new BatchTotalRow
            {
                Date = d, ObTotal = o, ProcessorTotal = p,
                Delta = o - p, Balanced = Math.Abs(o - p) <= opts.AmountTolerance,
            };
        }).ToList();
    }

    // ── helpers ──────────────────────────────────────────────────────────────────

    private static bool AmountClose(PaymentTxn a, PaymentTxn b, ReconMatchOptions o) =>
        Math.Abs(a.Amount - b.Amount) <= o.AmountTolerance;

    // Near = same sign, within the absolute OR percentage band (typo-explainable).
    private static bool AmountNear(PaymentTxn a, PaymentTxn b, ReconMatchOptions o)
    {
        if (Math.Sign(a.Amount) != Math.Sign(b.Amount)) return false;
        var delta = Math.Abs(a.Amount - b.Amount);
        var band  = Math.Max(o.PossibleAmountTolerance, Math.Abs(a.Amount) * o.PossiblePercent);
        return delta <= band;
    }

    private static bool DateClose(PaymentTxn a, PaymentTxn b, ReconMatchOptions o) =>
        Math.Abs((a.Date - b.Date).TotalDays) <= o.DateToleranceDays;

    private static string? Norm(string? s) => string.IsNullOrWhiteSpace(s) ? null : s.Trim();

    private static ReconRow Row(ReconRowStatus status, PaymentTxn? ob, PaymentTxn? proc,
        string? via = null, string? note = null)
    {
        var row = new ReconRow
        {
            Status      = status,
            Ob          = ob,
            Processor   = proc,
            AmountDelta = ob is not null && proc is not null ? ob.Amount - proc.Amount : null,
            MatchedVia  = via,
            Note        = note,
        };
        row.RowId = status == ReconRowStatus.Matched ? "" : ReconKey.For(row);
        return row;
    }

    private static ReconCounts Tally(List<ReconRow> rows) => new()
    {
        Matched            = rows.Count(r => r.Status == ReconRowStatus.Matched),
        PossibleMatch      = rows.Count(r => r.Status == ReconRowStatus.PossibleMatch),
        MissingInOb        = rows.Count(r => r.Status == ReconRowStatus.MissingInOb),
        MissingInProcessor = rows.Count(r => r.Status == ReconRowStatus.MissingInProcessor),
        Misclassified      = rows.Count(r => r.Status == ReconRowStatus.Misclassified),
        NeverCharged       = rows.Count(r => r.Status == ReconRowStatus.NeverCharged),
        Informational      = rows.Count(r => r.Status == ReconRowStatus.Informational),
        AmountMismatch     = rows.Count(r => r.Status == ReconRowStatus.AmountMismatch),
        Ambiguous          = rows.Count(r => r.Status == ReconRowStatus.Ambiguous),
    };

    /// <summary>
    /// Cross-references the card↔GP leftovers against OB's non-card payments (cash/check/ACH) to explain
    /// gaps: a card-side gap that matches an other-type entry is a MISCLASSIFICATION (wrong payment type);
    /// an OB card-only gap with no other-type match is NEVER CHARGED. Unused non-card OB payments are
    /// appended as Informational so the grid shows every payment from both sources. Recomputes Counts.
    /// </summary>
    public static void ApplyCrossTypeInsight(ReconRunResult result, IReadOnlyList<PaymentTxn> obOther, ReconMatchOptions opts)
    {
        var others = obOther.Select(t => new Item(t)).ToList();

        PaymentTxn? TakeMatch(PaymentTxn x)
        {
            var hit = others.FirstOrDefault(o => !o.Consumed
                && Math.Abs(o.Txn.Amount - x.Amount) <= opts.AmountTolerance
                && Math.Abs((o.Txn.Date - x.Date).TotalDays) <= opts.DateToleranceDays);
            if (hit is null) return null;
            hit.Consumed = true;
            return hit.Txn;
        }

        // GP-only (a card charge with no OB card payment).
        foreach (var row in result.Rows.Where(r => r.Status == ReconRowStatus.MissingInOb && r.Processor is not null).ToList())
        {
            var other = TakeMatch(row.Processor!);
            if (other is not null)
            {
                row.Status      = ReconRowStatus.Misclassified;
                row.Ob          = other;
                row.AmountDelta = other.Amount - row.Processor!.Amount;
                row.Note        = $"Paid by card per Heartland but recorded in OB as {other.PayType} — reclassify to card";
            }
            else row.Note = "No OB payment — create it (terminal charge not recorded)";
        }

        // OB card-only (an OB card entry with no GP charge).
        foreach (var row in result.Rows.Where(r => r.Status == ReconRowStatus.MissingInProcessor && r.Ob is not null).ToList())
        {
            var other = TakeMatch(row.Ob!);
            if (other is not null)
            {
                row.Status = ReconRowStatus.Misclassified;
                row.Note   = $"No Heartland charge but matches an OB {other.PayType} of {Fmt(other.Amount)} — likely mis-typed as card";
            }
            else
            {
                row.Status = ReconRowStatus.NeverCharged;
                row.Note   = "Marked paid in OB but not in Heartland — verify it was actually charged";
            }
        }

        // Remaining non-card OB payments — informational (expected; not card-reconciled).
        foreach (var o in others.Where(x => !x.Consumed))
            result.Rows.Add(Row(ReconRowStatus.Informational, o.Txn, null, note: o.Txn.PayType));

        foreach (var row in result.Rows)
            row.RowId = row.Status == ReconRowStatus.Matched ? "" : ReconKey.For(row);
        result.Counts = Tally(result.Rows);
    }

    private static string Fmt(decimal d) => d.ToString("C", CultureInfo.GetCultureInfo("en-US"));
}

/// <summary>
/// Run-independent key for a discrepancy — derived from identity fields ONLY (no run timestamps),
/// so a resolved item keeps the same key across runs and never resurfaces. Used as the
/// ReconDiscrepancies list Title and the ReconRow.RowId the client acts on.
/// </summary>
public static class ReconKey
{
    public static string For(ReconRow row)
    {
        var t       = row.Ob ?? row.Processor;
        var date    = t?.Date ?? default;
        var amount  = row.Ob?.Amount ?? row.Processor?.Amount ?? 0m;
        var last4   = row.Ob?.Last4    ?? row.Processor?.Last4;
        var auth    = row.Ob?.AuthCode ?? row.Processor?.AuthCode;
        var refr    = row.Ob?.SourceDoc ?? row.Processor?.Reference;
        var extra   = row.Status is ReconRowStatus.AmountMismatch or ReconRowStatus.PossibleMatch
            ? $"{row.Ob?.Amount.ToString(CultureInfo.InvariantCulture)}->{row.Processor?.Amount.ToString(CultureInfo.InvariantCulture)}"
            : "";
        var raw = string.Join('|', row.Status, date.ToString("yyyyMMdd"),
            amount.ToString("0.00", CultureInfo.InvariantCulture), last4, auth, refr, extra);
        var hash = SHA256.HashData(Encoding.UTF8.GetBytes(raw));
        return Convert.ToHexString(hash, 0, 8).ToLowerInvariant();
    }
}
