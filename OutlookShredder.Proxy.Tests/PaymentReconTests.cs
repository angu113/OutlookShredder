using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.Logging.Abstractions;
using OutlookShredder.Proxy.Models;
using OutlookShredder.Proxy.Services;
using Xunit;

namespace OutlookShredder.Proxy.Tests;

// Behaviour tests for the ShadowCat Payment Reconciliation core: the OB + Heartland CSV parsers
// (column-alias mapping, signed refunds, BOM, date-only) and the pure matcher (each outcome incl.
// ambiguity + a balanced-but-unmatched day) and the run-independent discrepancy key. Pure in/out.
public class PaymentReconTests
{
    private static readonly ReconMatchOptions Opts = new() { AmountTolerance = 0.01m, DateToleranceDays = 1 };
    private static readonly HeartlandColumnMap Map = new();

    private static PaymentTxn Ob(decimal amt, string date, string? last4 = null, string? auth = null,
        string? txn = null, string? doc = null, string? payType = null) =>
        new() { Source = "ob", Amount = amt, Date = DateTime.Parse(date).Date, Last4 = last4, AuthCode = auth, TxnId = txn, SourceDoc = doc, PayType = payType };

    private static PaymentTxn Pr(decimal amt, string date, string? last4 = null, string? auth = null,
        string? txn = null, string? reference = null) =>
        new() { Source = "processor", Amount = amt, Date = DateTime.Parse(date).Date, Last4 = last4, AuthCode = auth, TxnId = txn, Reference = reference };

    // ── parsers ───────────────────────────────────────────────────────────────

    [Fact]
    public void ObParser_real_columns_with_multiline_description_and_paytype()
    {
        // Real OB "ExportedData" header; row 1 has a MULTI-LINE quoted Description (embedded newlines).
        var csv =
            "Organization,Document No.,Description,Payment Date,Received From,Payment Method,Deposit To,Amount,Generated Credit,Used Credit,Card Number,Status\n" +
            "\"020803 Hackensack Store\",\"020803-ARR1025057\",\"Invoice No.: HSK-SI1023249\nOrder No.: HSK-SO1034199\n\",\"06-13-2026\",\"Sal Pellicone\",\"Credit & Debit Cards\",\"Bank\",65.4,0,0,\"\",\"Deposited not Cleared\"\n" +
            "\"020803 Hackensack Store\",\"020803-ARR1025058\",\"Order No.: HSK-SO1036156\n\",\"06-13-2026\",\"Yadit Ramot\",\"Check\",\"Bank\",97.7,0,0,\"\",\"Deposited not Cleared\"\n";
        var rows = ObPaymentsCsvParser.Parse(csv);

        Assert.Equal(2, rows.Count);                 // multi-line description did not split the record
        Assert.Equal("ob", rows[0].Source);
        Assert.Equal(65.4m, rows[0].Amount);
        Assert.Equal("Credit & Debit Cards", rows[0].PayType);
        Assert.Equal("Sal Pellicone", rows[0].Reference);
        Assert.Equal("020803-ARR1025057", rows[0].SourceDoc);
        Assert.Equal(new DateTime(2026, 6, 13), rows[0].Date);  // 06-13-2026 (m-d-yyyy)
        Assert.Null(rows[0].Last4);                   // OB carries no card number
        Assert.Equal("HSK-SI1023249", rows[0].HskRef);          // invoice preferred over order
        Assert.Equal("Check", rows[1].PayType);
        Assert.Equal("HSK-SO1036156", rows[1].HskRef);          // only an order on this row

        Assert.True(ObPaymentsCsvParser.IsCardMethod("Credit & Debit Cards"));
        Assert.False(ObPaymentsCsvParser.IsCardMethod("Check"));
    }

    [Fact]
    public void Possible_match_flags_small_amount_difference()
    {
        // 0.40 off — within the $1 possible band, outside the 0.01 exact tolerance -> PossibleMatch.
        var near = PaymentMatcher.Match([Ob(100.00m, "2026-06-10")], [Pr(100.40m, "2026-06-10")], Opts);
        var nrow = Assert.Single(near.Rows);
        Assert.Equal(ReconRowStatus.PossibleMatch, nrow.Status);
        Assert.Equal(-0.40m, nrow.AmountDelta);
        Assert.Equal(1, near.Counts.PossibleMatch);

        // 10 off on a 100 amount — beyond both bands (max($1, 5%)=$5) -> genuine gaps, not possible.
        var far = PaymentMatcher.Match([Ob(100.00m, "2026-06-10")], [Pr(110.00m, "2026-06-10")], Opts);
        Assert.Equal(0, far.Counts.PossibleMatch);
        Assert.Equal(1, far.Counts.MissingInProcessor);
        Assert.Equal(1, far.Counts.MissingInOb);
    }

    [Fact]
    public void HeartlandParser_uses_alias_map_and_strips_bom()
    {
        var csv = "﻿Transaction Date,Transaction Amount,Card Number,Authorization Code,Transaction ID,Card Type\n" +
                  "2026-06-10,100.00,XXXXXXXXXXXX1234,A1B2C3,TXN-9,VISA\n";
        var rows = HeartlandCsvParser.Parse(csv, Map);

        Assert.Single(rows);
        Assert.Equal("processor", rows[0].Source);
        Assert.Equal(100.00m, rows[0].Amount);
        Assert.Equal("1234", rows[0].Last4);
        Assert.Equal("TXN-9", rows[0].TxnId);
        Assert.Equal(new DateTime(2026, 6, 10), rows[0].Date);
    }

    [Fact]
    public void HeartlandParser_real_transaction_columns()
    {
        // Real header + a sale and a (parenthesised) return from the Merchant Batch Download export.
        var csv =
            "Merchantnumber_1,MerchantName,BatchNumber,CardType,CardNum,Expirationdate,TerminalID,TerminalDescription,SaleReturn,Trandate,AuthNumber,POSEntrySource,Dbt_Ind,Tran_time,Close_date,Amount,Textbox17\n" +
            "650000011899971,METAL SUPERMARKETS HACKENSACK,000184,AMEX,371664*****1016,0531,NULL,,S,06/08/2026,817481,O, ,15:36:01,06/08/2026,\"$3,743.96\",\"$75,719.54\"\n" +
            "650000011899971,METAL SUPERMARKETS HACKENSACK,000184,AMEX,379291*****2008, ,NULL,,R,06/08/2026,,T, ,14:27:27,06/08/2026,($211.33),\"$75,719.54\"\n";
        var rows = HeartlandCsvParser.Parse(csv, Map);

        Assert.Equal(2, rows.Count);
        var sale = rows[0];
        Assert.Equal(3743.96m, sale.Amount);        // "$3,743.96" (quoted, comma) -> 3743.96
        Assert.Equal("1016", sale.Last4);           // masked PAN 371664*****1016 -> 1016
        Assert.Equal("817481", sale.AuthCode);
        Assert.Equal("AMEX", sale.CardType);
        Assert.Equal("000184", sale.BatchId);
        Assert.Equal(new DateTime(2026, 6, 8), sale.Date);   // Close_date anchor

        var ret = rows[1];
        Assert.Equal(-211.33m, ret.Amount);         // ($211.33) -> negative (return)
        Assert.Equal("2008", ret.Last4);
        Assert.Null(ret.AuthCode);                  // empty auth on the return
    }

    [Fact]
    public void Parser_throws_when_required_columns_absent()
        => Assert.Throws<Exception>(() => HeartlandCsvParser.Parse("Foo,Bar\n1,2\n", Map));

    [Fact]
    public void Classifier_identifies_file_kinds_by_content()
    {
        var ob = "Organization,Document No.,Description,Payment Date,Received From,Payment Method,Deposit To,Amount,Card Number,Status\n" +
                 "x,P1,d,06-13-2026,Cust,Credit & Debit Cards,Bank,10,,ok\n";
        var hl = "Merchantnumber_1,MerchantName,BatchNumber,CardType,CardNum,Expirationdate,TerminalID,TerminalDescription,SaleReturn,Trandate,AuthNumber,POSEntrySource,Dbt_Ind,Tran_time,Close_date,Amount,Textbox17\n" +
                 "1,M,000184,V,4111******1111,0531,,,S,06/13/2026,A1,O, ,10:00:00,06/13/2026,10.00,100\n";
        var si = "Invoice Date,Document No.,Business Partner,Total Gross Amount,Total Paid,Payment Terms\n06/13/2026,INV1,ACME,100,0,Net 30\n";

        Assert.Equal(CsvKind.PaymentIn,    CsvClassifier.Classify(ob));
        Assert.Equal(CsvKind.Heartland,    CsvClassifier.Classify(hl));
        Assert.Equal(CsvKind.SalesInvoice, CsvClassifier.Classify(si));
        Assert.Equal(CsvKind.Unknown,      CsvClassifier.Classify("a,b,c\n1,2,3\n"));
    }

    // ── matcher ────────────────────────────────────────────────────────────────

    [Fact]
    public void Match_by_txnId_is_matched()
    {
        var r = PaymentMatcher.Match([Ob(100m, "2026-06-10", txn: "T1")], [Pr(100m, "2026-06-10", txn: "T1")], Opts);
        var row = Assert.Single(r.Rows);
        Assert.Equal(ReconRowStatus.Matched, row.Status);
        Assert.Equal("txnId", row.MatchedVia);
        Assert.Equal(1, r.Counts.Matched);
    }

    [Fact]
    public void Match_by_auth_with_amount_diff_is_amount_mismatch()
    {
        var r = PaymentMatcher.Match([Ob(100m, "2026-06-10", auth: "A9")], [Pr(100.50m, "2026-06-10", auth: "A9")], Opts);
        var row = Assert.Single(r.Rows);
        Assert.Equal(ReconRowStatus.AmountMismatch, row.Status);
        Assert.Equal(-0.50m, row.AmountDelta);   // ob - processor
        Assert.Equal(1, r.Counts.AmountMismatch);
    }

    [Fact]
    public void Match_by_amount_date_last4_then_amount_date()
    {
        // last4 present both sides -> matched via the stronger tier
        var a = PaymentMatcher.Match([Ob(75m, "2026-06-10", last4: "1111")], [Pr(75m, "2026-06-10", last4: "1111")], Opts);
        Assert.Equal("amt+date+last4", Assert.Single(a.Rows).MatchedVia);

        // no last4 -> matched via amount+date, and date within tolerance (+1 day)
        var b = PaymentMatcher.Match([Ob(75m, "2026-06-10")], [Pr(75m, "2026-06-11")], Opts);
        var row = Assert.Single(b.Rows);
        Assert.Equal(ReconRowStatus.Matched, row.Status);
        Assert.Equal("amt+date", row.MatchedVia);
    }

    [Fact]
    public void One_sided_rows_are_gaps()
    {
        var r = PaymentMatcher.Match([Ob(10m, "2026-06-10")], [Pr(20m, "2026-06-10")], Opts);
        Assert.Equal(1, r.Counts.MissingInProcessor); // OB 10 has no processor row
        Assert.Equal(1, r.Counts.MissingInOb);        // processor 20 has no OB row
    }

    [Fact]
    public void Duplicate_amount_same_day_without_ids_is_ambiguous()
    {
        var ob   = new[] { Ob(50m, "2026-06-10"), Ob(50m, "2026-06-10") };
        var proc = new[] { Pr(50m, "2026-06-10"), Pr(50m, "2026-06-10") };
        var r = PaymentMatcher.Match(ob, proc, Opts);
        Assert.Equal(4, r.Counts.Ambiguous);
        Assert.Equal(0, r.Counts.Matched);            // never a false match
    }

    [Fact]
    public void Refund_matches_refund_not_a_sale()
    {
        // a -100 refund and a +100 sale on the same day must NOT cross-match (signed amounts)
        var ob   = new[] { Ob(-100m, "2026-06-10", txn: "R1"), Ob(100m, "2026-06-10", txn: "S1") };
        var proc = new[] { Pr(-100m, "2026-06-10", txn: "R1"), Pr(100m, "2026-06-10", txn: "S1") };
        var r = PaymentMatcher.Match(ob, proc, Opts);
        Assert.Equal(2, r.Counts.Matched);
        Assert.Equal(0, r.Counts.AmountMismatch);
    }

    [Fact]
    public void Batch_totals_balance_even_when_rows_dont_match()
    {
        // OB has 100 + 50; processor reports a single 150 deposit — per-row unmatched, day balances.
        var ob   = new[] { Ob(100m, "2026-06-10"), Ob(50m, "2026-06-10") };
        var proc = new[] { Pr(150m, "2026-06-10") };
        var r = PaymentMatcher.Match(ob, proc, Opts);

        var batch = Assert.Single(r.BatchTotals);
        Assert.True(batch.Balanced);
        Assert.Equal(0m, batch.Delta);
        Assert.Equal(0, r.Counts.Matched); // per-row none matched
    }

    [Fact]
    public void CrossType_gp_only_matching_ob_other_is_misclassified()
    {
        // GP card $213.25 with no OB card payment, but an OB Check of $213.25 same day -> misclassified.
        var r = PaymentMatcher.Match([], [Pr(213.25m, "2026-06-13")], Opts);
        PaymentMatcher.ApplyCrossTypeInsight(r, [Ob(213.25m, "2026-06-13", payType: "Check")], Opts);

        Assert.Equal(1, r.Counts.Misclassified);
        Assert.Equal(0, r.Counts.MissingInOb);
        Assert.Equal(0, r.Counts.Informational);   // the check was consumed into the misclassified row
        Assert.Contains("Check", Assert.Single(r.Rows).Note);
    }

    [Fact]
    public void CrossType_gp_only_with_no_other_match_is_missing_in_ob()
    {
        var r = PaymentMatcher.Match([], [Pr(54.12m, "2026-06-13")], Opts);
        PaymentMatcher.ApplyCrossTypeInsight(r, [Ob(999m, "2026-06-13", payType: "Cash")], Opts);

        Assert.Equal(1, r.Counts.MissingInOb);     // GP charge, no OB payment -> create it
        Assert.Equal(1, r.Counts.Informational);   // the unrelated cash payment is just listed
        Assert.Equal(0, r.Counts.Misclassified);
    }

    [Fact]
    public void CrossType_ob_card_only_with_no_other_match_is_never_charged()
    {
        var r = PaymentMatcher.Match([Ob(745.31m, "2026-06-13")], [], Opts);
        PaymentMatcher.ApplyCrossTypeInsight(r, [], Opts);

        Assert.Equal(1, r.Counts.NeverCharged);    // OB card, no GP, no other-type match
        Assert.Equal(0, r.Counts.MissingInProcessor);
    }

    [Fact]
    public void Run_reconciles_only_the_target_business_day()
    {
        var svc = new PaymentReconciliationService(new ConfigurationBuilder().Build(),
            NullLogger<PaymentReconciliationService>.Instance);
        var ob =
            "Document No.,Description,Payment Date,Received From,Payment Method,Amount\n" +
            "P1,d,06-12-2026,Cust A,Credit & Debit Cards,100.00\n" +
            "P2,d,06-13-2026,Cust B,Credit & Debit Cards,200.00\n";
        var hl =
            "BatchNumber,CardType,CardNum,SaleReturn,Trandate,AuthNumber,Close_date,Amount\n" +
            "000184,V,4111******1111,S,06/12/2026,A1,06/12/2026,100.00\n" +
            "000185,V,4222******2222,S,06/13/2026,A2,06/13/2026,200.00\n";

        var r = svc.Run(ob, hl);

        Assert.Equal(new DateTime(2026, 6, 13), r.TargetDate);  // latest OB day
        Assert.Equal(1, r.Counts.Matched);                      // only the 06/13 pair
        Assert.Equal(0, r.Counts.MissingInOb);                  // 06/12 excluded, not flagged
        Assert.Equal(0, r.Counts.MissingInProcessor);
        Assert.Equal(1, r.ObCount);
        Assert.Equal(1, r.ProcessorCount);
    }

    [Fact]
    public void DiscrepancyKey_is_run_independent_and_status_specific()
    {
        var row1 = PaymentMatcher.Match([Ob(10m, "2026-06-10", doc: "D1")], [], Opts).Rows[0];
        var row2 = PaymentMatcher.Match([Ob(10m, "2026-06-10", doc: "D1")], [], Opts).Rows[0];
        Assert.False(string.IsNullOrEmpty(row1.RowId));
        Assert.Equal(row1.RowId, row2.RowId);  // same identity -> same key across runs

        var missingInOb = PaymentMatcher.Match([], [Pr(10m, "2026-06-10")], Opts).Rows[0];
        Assert.NotEqual(row1.RowId, missingInOb.RowId); // different kind -> different key
    }
}
