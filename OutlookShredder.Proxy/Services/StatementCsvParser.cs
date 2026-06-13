using System.Globalization;
using System.Text;
using System.Text.RegularExpressions;
using OutlookShredder.Proxy.Models;

namespace OutlookShredder.Proxy.Services;

/// <summary>
/// Parses the OpenBravo "Export to Spreadsheet" CSV for Sales Invoices into
/// CustomerStatementDtos with the same filtering logic as the Shredder StatementBuilder:
/// excludes credit memos (TotalGrossAmount &lt;= 0), excludes fully/overpaid lines (Amount &lt;= 0),
/// and excludes customers with no outstanding balance.
/// </summary>
public static class StatementCsvParser
{
    private static readonly Regex NetDaysRx = new(@"Net\s+(\d+)\s+Days?", RegexOptions.IgnoreCase);

    public static List<CustomerStatementDto> Parse(string csvContent)
    {
        var lines = csvContent.Split('\n');
        if (lines.Length < 2) return [];

        var headers = ParseRow(lines[0]);
        int idxDate  = IndexOf(headers, "Invoice Date");
        int idxDoc   = IndexOf(headers, "Document No.");
        int idxBp    = IndexOf(headers, "Business Partner");
        int idxGross = IndexOf(headers, "Total Gross Amount");
        int idxPaid  = IndexOf(headers, "Total Paid");
        int idxTerms = IndexOf(headers, "Payment Terms");

        if (idxDate < 0 || idxDoc < 0 || idxBp < 0 || idxGross < 0 || idxPaid < 0)
            throw new Exception(
                $"CSV missing required columns (found: {string.Join(", ", headers)}). " +
                "Expected: Invoice Date, Document No., Business Partner, Total Gross Amount, Total Paid.");

        var raw = new List<(string Bp, DateTime Date, string DocNo, decimal Gross, decimal Paid, string Terms)>();

        for (int i = 1; i < lines.Length; i++)
        {
            var line = lines[i].Trim('\r', ' ');
            if (string.IsNullOrEmpty(line)) continue;
            var cols = ParseRow(line);

            if (!decimal.TryParse(Col(cols, idxGross), NumberStyles.Any, CultureInfo.InvariantCulture, out decimal gross)) continue;
            if (gross <= 0) continue; // exclude credit memos

            if (!DateTime.TryParse(Col(cols, idxDate), out DateTime date)) continue;

            decimal.TryParse(Col(cols, idxPaid), NumberStyles.Any, CultureInfo.InvariantCulture, out decimal paid);

            raw.Add((
                Bp:    Col(cols, idxBp).Trim(),
                Date:  date,
                DocNo: Col(cols, idxDoc).Trim(),
                Gross: gross,
                Paid:  paid,
                Terms: idxTerms >= 0 ? Col(cols, idxTerms).Trim() : ""
            ));
        }

        return raw
            .GroupBy(r => r.Bp)
            .Select(g =>
            {
                var invoices = g
                    .OrderBy(r => r.Date)
                    .Select(r => new InvoiceLineDto
                    {
                        InvoiceDate   = r.Date.ToString("yyyy-MM-dd"),
                        InvoiceNumber = r.DocNo,
                        DueDate       = ParseDueDate(r.Date, r.Terms)?.ToString("yyyy-MM-dd"),
                        Amount        = r.Gross - r.Paid,
                    })
                    .Where(l => l.Amount > 0)   // exclude fully / overpaid
                    .ToList();

                var terms = g
                    .Select(r => r.Terms)
                    .Where(t => t.Length > 0)
                    .Distinct()
                    .OrderBy(t => t)
                    .FirstOrDefault() ?? "";

                return new CustomerStatementDto
                {
                    CustomerName = g.Key,
                    Terms        = terms,
                    Invoices     = invoices,
                };
            })
            .Where(s => s.Invoices.Count > 0)   // exclude zero-balance customers
            .OrderBy(s => s.CustomerName)
            .ToList();
    }

    private static DateTime? ParseDueDate(DateTime invoiceDate, string terms)
    {
        var m = NetDaysRx.Match(terms ?? "");
        if (m.Success && int.TryParse(m.Groups[1].Value, out int days))
            return invoiceDate.AddDays(days);
        return null;
    }

    private static int IndexOf(string[] headers, string name) =>
        Array.FindIndex(headers, h => h.Trim().Equals(name, StringComparison.OrdinalIgnoreCase));

    private static string Col(string[] cols, int idx) =>
        idx >= 0 && idx < cols.Length ? cols[idx] : "";

    // RFC 4180 CSV row parser — handles quoted fields containing commas and escaped quotes.
    private static string[] ParseRow(string line)
    {
        var result = new List<string>();
        int i = 0;
        while (i <= line.Length)
        {
            if (i == line.Length) { result.Add(""); break; }

            if (line[i] == '"')
            {
                i++;
                var sb = new StringBuilder();
                while (i < line.Length)
                {
                    if (line[i] == '"' && i + 1 < line.Length && line[i + 1] == '"')
                    { sb.Append('"'); i += 2; }
                    else if (line[i] == '"')
                    { i++; break; }
                    else
                    { sb.Append(line[i++]); }
                }
                result.Add(sb.ToString());
                if (i < line.Length && line[i] == ',') i++;
            }
            else
            {
                int start = i;
                while (i < line.Length && line[i] != ',') i++;
                result.Add(line[start..i].TrimEnd('\r'));
                if (i < line.Length) i++;
            }
        }
        return [.. result];
    }
}
