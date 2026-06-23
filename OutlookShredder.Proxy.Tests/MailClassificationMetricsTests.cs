using OutlookShredder.Proxy.Services;
using Xunit;

namespace OutlookShredder.Proxy.Tests;

// Pure-function tests for MailEvalService.ComputeMetrics.
// No model calls, no SP reads — entirely deterministic.
public class MailClassificationMetricsTests
{
    private static EvalResultItem R(string gold, string predicted, double confidence = 0.9, bool? match = null)
        => new()
        {
            MailItemId = gold + predicted,
            Gold       = gold,
            Predicted  = predicted,
            Confidence = confidence,
            Match      = match ?? string.Equals(gold, predicted, StringComparison.OrdinalIgnoreCase),
            Provider   = "claude",
        };

    private static readonly DateTimeOffset T0 = new(2026, 1, 1, 0, 0, 0, TimeSpan.Zero);
    private static readonly DateTimeOffset T1 = new(2026, 1, 1, 0, 1, 0, TimeSpan.Zero);

    [Fact]
    public void Perfect_predictions_produce_100pct_accuracy_and_P_R_F1_1()
    {
        var items = new List<EvalResultItem>
        {
            R("Supplier/RFQ Responses", "Supplier/RFQ Responses"),
            R("Customer/Orders",        "Customer/Orders"),
            R("Corporate/Receipts",     "Corporate/Receipts"),
        };
        var report = MailEvalService.ComputeMetrics(items, T0, T1);

        Assert.Equal(3, report.TotalItems);
        Assert.Equal(3, report.CorrectItems);
        Assert.Equal(1.0, report.OverallAccuracy, 6);

        var rfq = report.ByLeaf.First(m => m.Leaf == "Supplier/RFQ Responses");
        Assert.Equal(1.0, rfq.Precision, 6);
        Assert.Equal(1.0, rfq.Recall,    6);
        Assert.Equal(1.0, rfq.F1,        6);
    }

    [Fact]
    public void Confusion_matrix_captures_misclassifications()
    {
        // Two Order Confirmations predicted as RFQ Responses (the dangerous confusion pair).
        var items = new List<EvalResultItem>
        {
            R("Supplier/Order Confirmations", "Supplier/RFQ Responses"),
            R("Supplier/Order Confirmations", "Supplier/RFQ Responses"),
            R("Supplier/RFQ Responses",       "Supplier/RFQ Responses"),
        };
        var report = MailEvalService.ComputeMetrics(items, T0, T1);

        Assert.Equal(1, report.CorrectItems);

        // gold=Order Confirmations → predicted=RFQ Responses: 2 cases
        Assert.True(report.Confusion.ContainsKey("Supplier/Order Confirmations"));
        Assert.Equal(2, report.Confusion["Supplier/Order Confirmations"]["Supplier/RFQ Responses"]);

        // Precision for RFQ Responses: TP=1, FP=2 → 1/3
        var rfq = report.ByLeaf.First(m => m.Leaf == "Supplier/RFQ Responses");
        Assert.Equal(1.0 / 3.0, rfq.Precision, 6);
        Assert.Equal(1.0,       rfq.Recall,     6);
    }

    [Fact]
    public void Calibration_buckets_accuracy_matches_actual_correct_rate()
    {
        // All predictions at 0.85 confidence (bucket 8 = 0.8-0.9): 2 correct, 1 wrong.
        var items = new List<EvalResultItem>
        {
            R("A", "A", 0.85),
            R("A", "A", 0.85),
            R("A", "B", 0.85),
        };
        var report = MailEvalService.ComputeMetrics(items, T0, T1);
        var bucket = report.Calibration.First(b => b.LowConfidence == 0.8);

        Assert.Equal(3, bucket.Count);
        Assert.Equal(2, bucket.Correct);
        Assert.NotNull(bucket.Accuracy);
        Assert.Equal(2.0 / 3.0, bucket.Accuracy!.Value, 6);
    }

    [Fact]
    public void Empty_predictions_do_not_crash()
    {
        var report = MailEvalService.ComputeMetrics([], T0, T1);
        Assert.Equal(0, report.TotalItems);
        Assert.Equal(0.0, report.OverallAccuracy);
        Assert.Empty(report.ByLeaf);
    }

    [Fact]
    public void Provider_stats_tally_correctly()
    {
        var items = new List<EvalResultItem>
        {
            new() { Gold="A", Predicted="A", Match=true,  Provider="claude",  Confidence=0.9 },
            new() { Gold="B", Predicted="B", Match=true,  Provider="claude",  Confidence=0.9 },
            new() { Gold="C", Predicted="A", Match=false, Provider="gemini",  Confidence=0.7 },
        };
        var report = MailEvalService.ComputeMetrics(items, T0, T1);
        var claude = report.ByProvider.First(p => p.Provider == "claude");
        var gemini = report.ByProvider.First(p => p.Provider == "gemini");

        Assert.Equal(2, claude.Count);  Assert.Equal(2, claude.Correct);
        Assert.Equal(1, gemini.Count);  Assert.Equal(0, gemini.Correct);
    }

    [Fact]
    public void Report_timestamps_are_preserved()
    {
        var report = MailEvalService.ComputeMetrics([], T0, T1);
        Assert.Equal(T0, report.StartedAt);
        Assert.Equal(T1, report.FinishedAt);
    }
}
