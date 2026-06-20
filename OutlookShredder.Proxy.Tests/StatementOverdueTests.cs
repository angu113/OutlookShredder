using OutlookShredder.Proxy.Models;
using OutlookShredder.Proxy.Services;
using Xunit;

namespace OutlookShredder.Proxy.Tests;

// Behaviour tests for the proxy-side overdue/terms logic that drives the ShadowCat scheduled emails
// (StatementOverdue). Mirrors the client StatementTerms semantics: "Net N" -> N days; Immediate/blank -> 0.
public class StatementOverdueTests
{
    private static CustomerStatementDto Stmt(string name, string terms, params (string Date, decimal Amt)[] invs) =>
        new()
        {
            CustomerName = name,
            Terms        = terms,
            Invoices     = invs.Select(i => new InvoiceLineDto { InvoiceDate = i.Date, Amount = i.Amt }).ToList(),
        };

    [Theory]
    [InlineData("Net 30 Days", 30)]
    [InlineData("Net30", 30)]
    [InlineData("NET45", 45)]
    [InlineData("Immediate", 0)]
    [InlineData("COD", 0)]
    [InlineData("", 0)]
    [InlineData(null, 0)]
    public void ParseDays_parses_terms(string? terms, int expected) =>
        Assert.Equal(expected, StatementOverdue.ParseDays(terms));

    [Theory]
    [InlineData("Net 30 Days", TermsBucket.Net, true)]
    [InlineData("Net 30 Days", TermsBucket.Immediate, false)]
    [InlineData("Immediate", TermsBucket.Immediate, true)]
    [InlineData("", TermsBucket.Immediate, true)]
    [InlineData("COD", TermsBucket.Net, false)]
    public void InBucket_classifies(string? terms, TermsBucket bucket, bool expected) =>
        Assert.Equal(expected, StatementOverdue.InBucket(terms, bucket));

    [Fact]
    public void Overdue_counts_only_invoices_past_terms()
    {
        var asOf = new DateTime(2026, 6, 20);
        // Net 30: 5/11 (40 days) is past terms; 5/31 (20 days) is not.
        var s = Stmt("Acme", "Net 30 Days", ("2026-05-11", 100m), ("2026-05-31", 50m));
        Assert.Equal(100m, StatementOverdue.Overdue(s, asOf));
    }

    [Fact]
    public void Immediate_overdue_is_anything_dated_before_today()
    {
        var asOf = new DateTime(2026, 6, 20);
        var s = Stmt("Acme", "Immediate", ("2026-06-19", 75m), ("2026-06-20", 25m)); // today's invoice not overdue
        Assert.Equal(75m, StatementOverdue.Overdue(s, asOf));
    }

    [Fact]
    public void OverdueRows_filters_by_bucket_and_positive_overdue_sorted_by_name()
    {
        var asOf = new DateTime(2026, 6, 20);
        var statements = new List<CustomerStatementDto>
        {
            Stmt("Zeta Net",  "Net 30 Days", ("2026-05-01", 200m)),   // net, overdue
            Stmt("Alpha Net", "Net 30 Days", ("2026-06-19", 200m)),   // net, NOT overdue (recent)
            Stmt("Beta Imm",  "Immediate",   ("2026-06-01", 300m)),   // immediate, overdue
        };

        var net = StatementOverdue.OverdueRows(statements, TermsBucket.Net, asOf);
        var netRow = Assert.Single(net);
        Assert.Equal("Zeta Net", netRow.Customer);

        var imm = StatementOverdue.OverdueRows(statements, TermsBucket.Immediate, asOf);
        var immRow = Assert.Single(imm);
        Assert.Equal("Beta Imm", immRow.Customer);
        Assert.Equal(300m, immRow.Overdue);
    }

    [Fact]
    public void OverdueRows_null_statements_is_empty() =>
        Assert.Empty(StatementOverdue.OverdueRows(null, TermsBucket.Net, DateTime.Today));
}
