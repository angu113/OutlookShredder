using System.Globalization;
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
        var lines = CsvRowReader.SplitLines(csvContent);
        if (lines.Length < 2) return [];

        var headers = CsvRowReader.ParseRow(lines[0]);
        int idxDate  = CsvRowReader.IndexOf(headers, "Invoice Date");
        int idxDoc   = CsvRowReader.IndexOf(headers, "Document No.");
        int idxBp    = CsvRowReader.IndexOf(headers, "Business Partner");
        int idxGross = CsvRowReader.IndexOf(headers, "Total Gross Amount");
        int idxPaid  = CsvRowReader.IndexOf(headers, "Total Paid");
        int idxTerms = CsvRowReader.IndexOf(headers, "Payment Terms");
        int idxPo    = CsvRowReader.IndexOf(headers, "Customer PO Number");

        if (idxDate < 0 || idxDoc < 0 || idxBp < 0 || idxGross < 0 || idxPaid < 0)
            throw new Exception(
                $"CSV missing required columns (found: {string.Join(", ", headers)}). " +
                "Expected: Invoice Date, Document No., Business Partner, Total Gross Amount, Total Paid.");

        var raw = new List<(string Bp, DateTime Date, string DocNo, string Po, decimal Gross, decimal Paid, string Terms)>();

        for (int i = 1; i < lines.Length; i++)
        {
            var line = lines[i].Trim('\r', ' ');
            if (string.IsNullOrEmpty(line)) continue;
            var cols = CsvRowReader.ParseRow(line);

            if (!decimal.TryParse(CsvRowReader.Col(cols, idxGross), NumberStyles.Any, CultureInfo.InvariantCulture, out decimal gross)) continue;
            if (gross <= 0) continue; // exclude credit memos

            if (!DateTime.TryParse(CsvRowReader.Col(cols, idxDate), out DateTime date)) continue;

            decimal.TryParse(CsvRowReader.Col(cols, idxPaid), NumberStyles.Any, CultureInfo.InvariantCulture, out decimal paid);

            raw.Add((
                Bp:    CsvRowReader.Col(cols, idxBp).Trim(),
                Date:  date,
                DocNo: CsvRowReader.Col(cols, idxDoc).Trim(),
                Po:    idxPo >= 0 ? CsvRowReader.Col(cols, idxPo).Trim() : "",
                Gross: gross,
                Paid:  paid,
                Terms: idxTerms >= 0 ? CsvRowReader.Col(cols, idxTerms).Trim() : ""
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
                        CustomerPO    = r.Po,
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
}
