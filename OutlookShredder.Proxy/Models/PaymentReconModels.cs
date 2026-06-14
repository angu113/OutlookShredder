using System.Text.Json.Serialization;

namespace OutlookShredder.Proxy.Models;

/// <summary>
/// A payment transaction normalized from either OpenBravo or the merchant processor, for
/// reconciliation. Amounts are SIGNED (refunds/chargebacks negative) and dates are date-only.
/// </summary>
public sealed class PaymentTxn
{
    public string   Source    { get; init; } = "";   // "ob" | "processor"
    public DateTime Date      { get; init; }          // date-only semantics
    public decimal  Amount    { get; init; }          // signed
    public string?  Last4     { get; init; }
    public string?  AuthCode  { get; init; }
    public string?  TxnId     { get; init; }
    public string?  CardType  { get; init; }          // card brand (Heartland: AMEX/DISC/MC/V)
    public string?  PayType   { get; init; }          // OB Payment Method (Credit & Debit Cards / Check / Cash / ACH …)
    public string?  BatchId   { get; init; }
    public string?  Reference { get; init; }          // OB: Received From (business partner); processor: any ref
    public string?  SourceDoc { get; init; }          // OB payment document no.
    public string?  HskRef    { get; init; }          // OB HSK-SI / HSK-SO from the payment Description
    public string   RawKey    { get; init; } = "";    // stable per-row identity for audit
}

/// <summary>
/// PossibleMatch = amounts close but not exact (typo-explainable). Misclassified = a card-side gap that
/// matches an OB cash/check/ACH entry (wrong payment type). NeverCharged = OB card with no GP charge and
/// no other-type match. Informational = an OB non-card payment shown for completeness (not card-reconciled).
/// </summary>
[JsonConverter(typeof(JsonStringEnumConverter))]
public enum ReconRowStatus
{
    Matched, PossibleMatch, MissingInOb, MissingInProcessor,
    Misclassified, NeverCharged, Informational, AmountMismatch, Ambiguous
}

/// <summary>One reconciliation outcome: a matched pair, a one-sided gap, or an ambiguity.</summary>
public sealed class ReconRow
{
    public string         RowId       { get; set; } = "";   // = DiscrepancyKey for non-matched rows
    public ReconRowStatus Status      { get; set; }
    public PaymentTxn?    Ob          { get; set; }
    public PaymentTxn?    Processor   { get; set; }
    public decimal?       AmountDelta { get; set; }         // ob - processor when both present
    public string?        MatchedVia  { get; set; }         // "txnId" | "auth" | "amt+date+last4" | "amt+date"
    public string?        Note        { get; set; }
    public string?        Resolution  { get; set; }         // open | resolved | ignored (discrepancies; null for Matched)
}

/// <summary>Per-day (optionally per-batch) total comparison — the aggregate safety net.</summary>
public sealed class BatchTotalRow
{
    public DateTime Date           { get; set; }
    public string?  BatchId        { get; set; }
    public decimal  ObTotal        { get; set; }
    public decimal  ProcessorTotal { get; set; }
    public decimal  Delta          { get; set; }            // ob - processor
    public bool     Balanced       { get; set; }
}

public sealed class ReconCounts
{
    public int Matched            { get; set; }
    public int PossibleMatch      { get; set; }
    public int MissingInOb        { get; set; }
    public int MissingInProcessor { get; set; }
    public int Misclassified      { get; set; }
    public int NeverCharged       { get; set; }
    public int Informational      { get; set; }
    public int AmountMismatch     { get; set; }
    public int Ambiguous          { get; set; }
}

/// <summary>OB payment-type subtotal (e.g. Credit &amp; Debit Cards / Check / Cash / ACH) for the summary view.</summary>
public sealed class PayTypeSubtotal
{
    public string  PayType { get; set; } = "";
    public int     Count   { get; set; }
    public decimal Total   { get; set; }
}

/// <summary>The full result of one reconciliation run (cached + persisted as JSON).</summary>
public sealed class ReconRunResult
{
    public DateTime            RunAt          { get; set; }
    public DateTime?           TargetDate     { get; set; }   // the single business day reconciled
    public string?             ObSource       { get; set; }
    public string?             ProcessorSource{ get; set; }
    public int                 ObCount        { get; set; }   // OB card rows on the target day
    public int                 ProcessorCount { get; set; }   // Heartland rows on the target day
    public List<ReconRow>        Rows        { get; set; } = [];
    public List<BatchTotalRow>   BatchTotals { get; set; } = [];
    public List<PayTypeSubtotal> Subtotals   { get; set; } = [];   // OB payment detail subtotaled by type
    public ReconCounts           Counts      { get; set; } = new();
}

/// <summary>
/// A discrepancy whose resolution state is tracked durably in SharePoint. The DiscrepancyKey is
/// run-independent (derived from identity fields only) so a resolved item never resurfaces on re-run.
/// </summary>
public sealed class ReconDiscrepancy
{
    public string         SpItemId       { get; set; } = "";
    public string         DiscrepancyKey { get; set; } = "";
    public string         Status         { get; set; } = "open";  // open | resolved | ignored
    public ReconRowStatus Kind           { get; set; }
    public DateTime       Date           { get; set; }
    public decimal        Amount         { get; set; }
    public decimal?       AmountDelta    { get; set; }
    public string?        Last4          { get; set; }
    public string?        AuthCode       { get; set; }
    public string?        Reference      { get; set; }
    public string?        Note           { get; set; }
    public string?        ResolvedBy     { get; set; }
    public DateTime?      ResolvedAt     { get; set; }
}
