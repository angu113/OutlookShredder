using System.Text.RegularExpressions;
using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.Hosting;
using Microsoft.Extensions.Logging;
using OutlookShredder.Proxy.Models;

namespace OutlookShredder.Proxy.Services;

/// <summary>
/// Fulfillment loop (wip/fulfillment-loop.md): matches a classified payment-processor BILL
/// (Supplier/Invoices and Bills) to the open PurchaseOrder it pays for, then flips that PO to
/// PaymentStatus="Required" and stores the bill's MailItemId so the surface can open the bill + its
/// pay link.
///
/// Bills do NOT reliably carry our PO# (the supplier's processor addresses us by their own refs), so
/// matching is multi-signal, in priority:
///   1. our PoNumber on the bill (subject/PDF) -> exact PO match (opportunistic, strongest).
///   2. supplierReference == a captured SLI QuoteReference -> that RFQ's PO (+ supplier).
///   3. exact amount == the PO's total (sum of the supplier's SLI TotalPrice for the RFQ), scoped by
///      fuzzy supplier — the user's "worst case, the total should match exactly".
/// A signal counts only when it resolves to EXACTLY ONE candidate PO; otherwise the bill is left
/// unmatched/ambiguous (surfaced, never silently mis-flagged).
///
/// SUGGEST-first: by default (BillMatcher:Auto=false) it only proposes matches (logged + returned by
/// the on-demand endpoint). Set Auto=true (or call the endpoint with apply=true) to write
/// PaymentStatus=Required. POs already Paid/Required and POs with no derivable total/ref (rfqId
/// UNKNOWN/000000, raised outside the RFQ flow) can only be matched by an explicit Po# on the bill.
/// </summary>
public sealed class BillToPoMatcherService : IHostedService, IDisposable
{
    private const string BillCategory = "Supplier/Invoices and Bills";
    private static readonly Regex HskPoRx = new(@"\bHSK-PO\d+\b", RegexOptions.IgnoreCase | RegexOptions.Compiled);
    private static readonly Regex NonAlnum = new(@"[^A-Za-z0-9]", RegexOptions.Compiled);

    // Guard: a payment-processor AUTH/RECEIPT (the card was already charged) sometimes lands under
    // "Supplier/Invoices and Bills" instead of "Supplier/Receipts". It must NOT flag PaymentStatus=
    // Required — the payment already happened. Skip it by subject markers ("Receipt", "approved",
    // "thank you for your payment", "authoriz…", "charged") or a known auth/processor sender.
    private static readonly Regex AuthReceiptSubjectRx = new(
        @"\b(receipt|payment\s+received|payment\s+confirmation|thank\s+you\s+for\s+your\s+payment|authoriz|auth\s*code|approved|charged|paid\s+in\s+full|transaction\s+approved)\b",
        RegexOptions.IgnoreCase | RegexOptions.Compiled);
    private static readonly Regex AuthReceiptSenderRx = new(
        @"(creditcardauth|slimcd\.com)",
        RegexOptions.IgnoreCase | RegexOptions.Compiled);

    private static bool LooksLikeAuthOrReceipt(string? subject, string? from)
        => (!string.IsNullOrWhiteSpace(subject) && AuthReceiptSubjectRx.IsMatch(subject))
        || (!string.IsNullOrWhiteSpace(from)    && AuthReceiptSenderRx.IsMatch(from));

    private readonly SharePointService   _sp;
    private readonly MailCacheService    _mail;
    private readonly SliCacheService     _sli;
    private readonly SupplierCacheService _suppliers;
    private readonly RfqNotificationService _notifications;
    private readonly IConfiguration      _config;
    private readonly ILogger<BillToPoMatcherService> _log;
    private Timer? _timer;
    private Timer? _kickTimer;
    private readonly object _kickLock = new();

    public BillToPoMatcherService(SharePointService sp, MailCacheService mail, SliCacheService sli,
        SupplierCacheService suppliers, RfqNotificationService notifications, IConfiguration config,
        ILogger<BillToPoMatcherService> log)
    {
        _sp = sp; _mail = mail; _sli = sli; _suppliers = suppliers; _notifications = notifications;
        _config = config; _log = log;
    }

    private bool   Enabled         => _config.GetValue("BillMatcher:Enabled", true);
    private bool   Auto            => _config.GetValue("BillMatcher:Auto", true);    // auto-apply on a confident single match (supplier+amount, or supplier-ref = 100%)
    private int    IntervalMinutes => Math.Max(1, _config.GetValue("BillMatcher:IntervalMinutes", 15));
    private decimal AmountTolerance => _config.GetValue("BillMatcher:AmountTolerance", 0.01m);

    public Task StartAsync(CancellationToken ct)
    {
        if (!Enabled) { _log.LogInformation("[BillMatcher] disabled"); return Task.CompletedTask; }
        _timer = new Timer(async _ =>
        {
            try { await RunOnceAsync(apply: Auto, CancellationToken.None); }
            catch (Exception ex) { _log.LogWarning(ex, "[BillMatcher] cycle failed"); }
        }, null, TimeSpan.FromSeconds(75), TimeSpan.FromMinutes(IntervalMinutes));   // delay so caches warm
        return Task.CompletedTask;
    }

    public Task StopAsync(CancellationToken ct) { _timer?.Dispose(); _kickTimer?.Dispose(); return Task.CompletedTask; }
    public void Dispose() { _timer?.Dispose(); _kickTimer?.Dispose(); }

    /// <summary>Coalesces post-classification triggers into one run a few seconds after the last one,
    /// so a freshly-classified bill is matched + flips its PO to Payment-required within seconds.</summary>
    public void KickSoon()
    {
        if (!Enabled) return;
        lock (_kickLock)
        {
            _kickTimer?.Dispose();
            _kickTimer = new Timer(async _ =>
            {
                try { await RunOnceAsync(apply: Auto, CancellationToken.None); }
                catch (Exception ex) { _log.LogWarning(ex, "[BillMatcher] kick failed"); }
            }, null, TimeSpan.FromSeconds(3), Timeout.InfiniteTimeSpan);
        }
    }

    public record BillMatch(string MailItemId, string Subject, string? BillSupplier, string? Amount,
        string? SupplierRef, string? BillPoNumber, string MatchedPo, string MatchedPoSpItemId, string Via, bool Applied);
    public record BillMatchResult(int Scanned, int Matched, int Applied, int SkippedReceipts, bool Auto,
        List<BillMatch> Matches, List<string> Ambiguous);

    /// <summary>One matching pass. apply=false → suggest only (no writes). apply=true → write
    /// PaymentStatus=Required + the bill pointer for confident single matches.</summary>
    public async Task<BillMatchResult> RunOnceAsync(bool apply, CancellationToken ct)
    {
        // Candidate POs: not already Required/Paid. (Bills can arrive late, so no stale filter.)
        var pos = (await _sp.ReadPurchaseOrdersAsync())
            .Where(p => !string.Equals(p.PaymentStatus, "Required", StringComparison.OrdinalIgnoreCase)
                     && !string.Equals(p.PaymentStatus, "Paid",     StringComparison.OrdinalIgnoreCase))
            .ToList();
        var byPoNumber = pos.Where(p => !string.IsNullOrWhiteSpace(p.PoNumber))
            .GroupBy(p => p.PoNumber!.Trim().ToUpperInvariant())
            .ToDictionary(g => g.Key, g => g.ToList());

        // SLI-derived lookups (in-memory cache): QuoteReference -> RFQs, and (RFQ|supplier) -> total.
        var sli = _sli.TryGet();
        if (sli is null) { await _sli.PopulateAsync(false, ct); sli = _sli.TryGet() ?? new(); }
        var quoteRefToRfqs = new Dictionary<string, HashSet<string>>(StringComparer.Ordinal);   // normRef -> rfqIds
        var rfqSupplierTotal = new Dictionary<string, decimal>(StringComparer.OrdinalIgnoreCase);  // "rfq|supplier" -> sum
        foreach (var r in sli)
        {
            var rfq = (S(r, "RFQ_ID") ?? S(r, "RFQ_x005F_ID") ?? "").Trim();
            if (rfq.Length == 0) continue;
            var sup = S(r, "SupplierName") ?? "";
            var qref = NormRef(S(r, "QuoteReference"));
            if (qref.Length >= 3)
            {
                if (!quoteRefToRfqs.TryGetValue(qref, out var set)) quoteRefToRfqs[qref] = set = new(StringComparer.OrdinalIgnoreCase);
                set.Add(rfq);
            }
            if (decimal.TryParse(S(r, "TotalPrice"), out var tp) && tp > 0)
            {
                var key = $"{rfq}|{sup}";
                rfqSupplierTotal[key] = (rfqSupplierTotal.TryGetValue(key, out var cur) ? cur : 0m) + tp;
            }
        }

        int scanned = 0, matched = 0, applied = 0, skippedReceipts = 0;
        var matches   = new List<BillMatch>();
        var ambiguous = new List<string>();

        foreach (var c in _mail.GetCurrents())
        {
            if (!string.Equals(c.CategoryPath, BillCategory, StringComparison.OrdinalIgnoreCase)) continue;
            if (!_mail.TryGetItem(c.MailItemId, out var item)) continue;
            if (string.Equals(item.Direction, "out", StringComparison.OrdinalIgnoreCase)) continue;
            if (item.Completed) continue;
            scanned++;

            // Skip auth/receipt confirmations misfiled under bills — the payment already happened, so
            // they must not flag PaymentStatus=Required.
            if (LooksLikeAuthOrReceipt(item.Subject, item.FromAddress))
            {
                skippedReceipts++;
                _log.LogDebug("[BillMatcher] skipped auth/receipt \"{Subj}\" [{Id}]",
                    item.Subject.Length > 80 ? item.Subject[..80] : item.Subject, c.MailItemId);
                continue;
            }

            var billSupplier = _suppliers.ResolveSupplierName(c.SupplierName) ?? c.SupplierName;
            var subj = item.Subject.Length > 80 ? item.Subject[..80] : item.Subject;

            // Resolve a single candidate PO by the strongest signal that yields exactly one.
            var (po, via) = ResolveSinglePo(c, billSupplier, byPoNumber, pos, quoteRefToRfqs, rfqSupplierTotal, out var ambiguityReason);
            if (po is null)
            {
                if (ambiguityReason is not null) ambiguous.Add($"\"{subj}\" [{c.MailItemId}] - {ambiguityReason}");
                continue;
            }
            matched++;
            var note = $"Bill matched via {via}: \"{subj}\" supplier={billSupplier} amount={c.Amount} ref={c.SupplierReference} [{c.MailItemId}]";
            if (apply)
            {
                await _sp.UpdatePurchaseOrderBillMatchAsync(po.SpItemId, c.MailItemId, c.Amount, c.SupplierReference, note);
                applied++;
                po.PaymentStatus  = "Required";
                po.BillMailItemId = c.MailItemId;
                _notifications.NotifyPoStatus(po);   // live re-colour + payment icon on the Ordered board
                // Don't match this PO again in this pass.
                if (!string.IsNullOrWhiteSpace(po.PoNumber) && byPoNumber.TryGetValue(po.PoNumber!.Trim().ToUpperInvariant(), out var l)) l.Remove(po);
                pos.Remove(po);
            }
            matches.Add(new BillMatch(c.MailItemId, subj, billSupplier, c.Amount, c.SupplierReference,
                c.PoNumber, po.PoNumber ?? "", po.SpItemId, via, apply));
        }

        _log.LogInformation("[BillMatcher] pass: scanned={S} matched={M} applied={A} skippedReceipts={R} (auto={Auto})",
            scanned, matched, applied, skippedReceipts, apply);
        return new BillMatchResult(scanned, matched, applied, skippedReceipts, apply, matches, ambiguous);
    }

    /// <summary>Returns the single best PO for a bill, or null with a reason when none/ambiguous.</summary>
    private PurchaseOrderRecord? ResolveSinglePoCore(MailClassRow c, string? billSupplier,
        Dictionary<string, List<PurchaseOrderRecord>> byPoNumber, List<PurchaseOrderRecord> pos,
        Dictionary<string, HashSet<string>> quoteRefToRfqs, Dictionary<string, decimal> rfqSupplierTotal,
        out string via, out string? ambiguity)
    {
        via = ""; ambiguity = null;

        // 1) Our PO# printed on the bill (subject/body/PDF) — exact, strongest.
        foreach (var pn in CandidatePoNumbers(c))
            if (byPoNumber.TryGetValue(pn, out var l) && l.Count == 1) { via = "po-number"; return l[0]; }

        // 2) supplierReference == a captured SLI QuoteReference -> that RFQ's PO (supplier-scoped).
        var nref = NormRef(c.SupplierReference);
        if (nref.Length >= 3 && quoteRefToRfqs.TryGetValue(nref, out var rfqs))
        {
            var cands = pos.Where(p => rfqs.Contains(p.RfqId) && SupplierMatches(billSupplier, p.SupplierName)).Distinct().ToList();
            if (cands.Count == 1) { via = "supplier-ref"; return cands[0]; }
            if (cands.Count > 1) { ambiguity = $"supplier-ref '{c.SupplierReference}' matched {cands.Count} POs"; return null; }
        }

        // 3) exact amount == PO total (SLI sum), scoped by fuzzy supplier.
        if (decimal.TryParse(c.Amount, out var billAmt) && billAmt > 0)
        {
            var cands = pos.Where(p =>
                    SupplierMatches(billSupplier, p.SupplierName)
                    && rfqSupplierTotal.TryGetValue($"{p.RfqId}|{p.SupplierName}", out var tot)
                    && Math.Abs(tot - billAmt) <= AmountTolerance)
                .ToList();
            if (cands.Count == 1) { via = "amount"; return cands[0]; }
            if (cands.Count > 1) { ambiguity = $"amount {c.Amount} matched {cands.Count} POs for supplier '{billSupplier}'"; return null; }
        }

        ambiguity ??= "no confident PO match (no PO#, supplier-ref, or exact-amount hit)";
        return null;
    }

    private (PurchaseOrderRecord? Po, string Via) ResolveSinglePo(MailClassRow c, string? billSupplier,
        Dictionary<string, List<PurchaseOrderRecord>> byPoNumber, List<PurchaseOrderRecord> pos,
        Dictionary<string, HashSet<string>> quoteRefToRfqs, Dictionary<string, decimal> rfqSupplierTotal,
        out string? ambiguity)
    {
        var po = ResolveSinglePoCore(c, billSupplier, byPoNumber, pos, quoteRefToRfqs, rfqSupplierTotal, out var via, out ambiguity);
        return (po, via);
    }

    private static IEnumerable<string> CandidatePoNumbers(MailClassRow c)
    {
        var set = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        var pn = c.PoNumber?.Trim();
        if (!string.IsNullOrEmpty(pn))
        {
            foreach (Match m in HskPoRx.Matches(pn)) set.Add(m.Value.ToUpperInvariant());
            if (pn.StartsWith("HSK-PO", StringComparison.OrdinalIgnoreCase)) set.Add(pn.ToUpperInvariant());
        }
        return set;
    }

    private bool SupplierMatches(string? billSupplier, string? poSupplier)
    {
        if (string.IsNullOrWhiteSpace(billSupplier) || string.IsNullOrWhiteSpace(poSupplier)) return false;
        var a = _suppliers.ResolveSupplierName(billSupplier) ?? billSupplier;
        var b = _suppliers.ResolveSupplierName(poSupplier)   ?? poSupplier;
        if (string.Equals(a, b, StringComparison.OrdinalIgnoreCase)) return true;
        var na = NormRef(a); var nb = NormRef(b);
        return na.Length > 2 && nb.Length > 2 && (na.Contains(nb) || nb.Contains(na));
    }

    private static string NormRef(string? s) => string.IsNullOrWhiteSpace(s) ? "" : NonAlnum.Replace(s, "").ToUpperInvariant();
    private static string? S(Dictionary<string, object?> d, string k) => d.TryGetValue(k, out var v) ? v?.ToString() : null;
}
