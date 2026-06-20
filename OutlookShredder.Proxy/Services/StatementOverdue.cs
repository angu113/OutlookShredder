using System.Globalization;
using System.Text.RegularExpressions;
using OutlookShredder.Proxy.Models;

namespace OutlookShredder.Proxy.Services;

public enum TermsBucket { Immediate, Net }

/// <summary>One overdue customer row for the ShadowCat overdue emails (mirrors the client OverdueRow).</summary>
public sealed record OverdueRow(string Customer, string Terms, decimal Due, decimal Overdue);

/// <summary>
/// Proxy-side payment-terms + overdue logic, mirroring the client <c>StatementTerms</c> so the scheduled
/// overdue emails compute exactly as the Generate Statement tab does: "Net N" -> N days; Immediate / blank /
/// unparseable -> 0 (due-on-receipt). Overdue = invoices dated before (asOf - termDays).
/// </summary>
public static class StatementOverdue
{
    private static readonly Regex NetRx = new(@"net\s*(\d+)", RegexOptions.IgnoreCase | RegexOptions.Compiled);

    public static int ParseDays(string? terms)
    {
        if (string.IsNullOrWhiteSpace(terms)) return 0;
        var m = NetRx.Match(terms);
        return m.Success && int.TryParse(m.Groups[1].Value, out var d) ? d : 0;
    }

    public static bool InBucket(string? terms, TermsBucket bucket) =>
        bucket == TermsBucket.Immediate ? ParseDays(terms) == 0 : ParseDays(terms) > 0;

    public static decimal DueTotal(CustomerStatementDto s) => s.Invoices.Sum(i => i.Amount);

    public static decimal Overdue(CustomerStatementDto s, DateTime asOf)
    {
        var cutoff = asOf.Date.AddDays(-ParseDays(s.Terms));
        return s.Invoices
            .Where(i => DateTime.TryParse(i.InvoiceDate, CultureInfo.InvariantCulture,
                            DateTimeStyles.None, out var d) && d.Date < cutoff)
            .Sum(i => i.Amount);
    }

    /// <summary>Customers in the given terms bucket with a past-terms balance (Overdue &gt; 0), ordered by name.</summary>
    public static List<OverdueRow> OverdueRows(
        IReadOnlyList<CustomerStatementDto>? statements, TermsBucket bucket, DateTime asOf)
    {
        if (statements is null) return [];
        return statements
            .Where(s => InBucket(s.Terms, bucket))
            .Select(s => new OverdueRow(s.CustomerName, s.Terms, DueTotal(s), Overdue(s, asOf)))
            .Where(r => r.Overdue > 0)
            .OrderBy(r => r.Customer, StringComparer.OrdinalIgnoreCase)
            .ToList();
    }

    public static string Label(string? terms) => string.IsNullOrWhiteSpace(terms) ? "Immediate" : terms.Trim();
}
