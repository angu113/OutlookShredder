namespace OutlookShredder.Proxy.Services;

/// <summary>The kind of export a CSV is, identified by its header columns (not its filename).</summary>
public enum CsvKind { Unknown, SalesInvoice, PaymentIn, Heartland }

/// <summary>
/// Identifies an exported CSV by its CONTENT (header columns), so the right parser/flow handles it.
/// The OB Sales-Invoice and Payment-In grid exports are both named <c>ExportedData*.csv</c> and land in
/// the same Downloads folder watched by overlapping watchers — content classification is what keeps a
/// payments run from accidentally consuming an invoice file (or vice versa).
/// </summary>
public static class CsvClassifier
{
    public static CsvKind Classify(string csvContent)
    {
        var records = CsvRowReader.ParseRecords(csvContent);
        if (records.Count == 0) return CsvKind.Unknown;

        var headers = records[0].Select(h => h.Trim().ToLowerInvariant()).ToHashSet();
        bool Has(string name) => headers.Contains(name);

        // Heartland transaction export.
        if (Has("batchnumber") && Has("authnumber") && (Has("close_date") || Has("trandate")))
            return CsvKind.Heartland;

        // OB Payment In grid export.
        if (Has("payment method") && Has("payment date") && Has("amount") && Has("received from"))
            return CsvKind.PaymentIn;

        // OB Sales Invoice grid export.
        if (Has("invoice date") && Has("total gross amount"))
            return CsvKind.SalesInvoice;

        return CsvKind.Unknown;
    }
}
