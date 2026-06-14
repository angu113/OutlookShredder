using System.Text.RegularExpressions;
using OutlookShredder.Proxy.Models;

namespace OutlookShredder.Proxy.Services;

/// <summary>
/// Parses the OpenBravo "Payment In" grid export ("ExportedData" CSV) into <see cref="PaymentTxn"/>s.
/// Columns: Organization, Document No., Description (MULTI-LINE), Payment Date, Received From,
/// Payment Method, Deposit To, Amount, …, Card Number (empty — OB carries no card detail), Status.
/// Parses ALL payment methods (Credit &amp; Debit Cards / Check / Cash / ACH …) so the caller can
/// subtotal per type; reconciliation filters to the card rows. Uses the record parser because the
/// Description cell contains embedded newlines. Amounts are SIGNED.
/// </summary>
public static class ObPaymentsCsvParser
{
    private static readonly string[] DateAliases   = ["Payment Date", "Date", "Document Date"];
    private static readonly string[] AmountAliases = ["Amount", "Payment Amount", "Paid"];
    private static readonly string[] DocAliases    = ["Document No.", "Document No", "Payment No."];
    private static readonly string[] PartnerAlias  = ["Received From", "Business Partner", "Customer"];
    private static readonly string[] MethodAliases = ["Payment Method", "Pay Method", "Tender"];
    private static readonly string[] CardAliases   = ["Card Number", "Card No.", "Last 4"];
    private static readonly string[] DescAliases   = ["Description", "Payment Description"];

    // HSK reference embedded in the Description, e.g. "Invoice No.: HSK-SI1024569 / Order No.: HSK-SO1036155".
    private static readonly Regex HskRx = new(@"HSK-[A-Z]{2}\d+", RegexOptions.IgnoreCase | RegexOptions.Compiled);

    public static List<PaymentTxn> Parse(string csvContent)
    {
        var records = CsvRowReader.ParseRecords(csvContent);
        if (records.Count < 2) return [];

        var headers = records[0];
        int iDate = CsvRowReader.IndexOfAny(headers, DateAliases);
        int iAmt  = CsvRowReader.IndexOfAny(headers, AmountAliases);
        int iDoc  = CsvRowReader.IndexOfAny(headers, DocAliases);
        int iBp   = CsvRowReader.IndexOfAny(headers, PartnerAlias);
        int iMeth = CsvRowReader.IndexOfAny(headers, MethodAliases);
        int iCard = CsvRowReader.IndexOfAny(headers, CardAliases);
        int iDesc = CsvRowReader.IndexOfAny(headers, DescAliases);

        if (iDate < 0 || iAmt < 0)
            throw new Exception(
                $"OB payments CSV missing required Date/Amount columns (found: {string.Join(", ", headers)}).");

        var list = new List<PaymentTxn>();
        for (int i = 1; i < records.Count; i++)
        {
            var cols = records[i];
            if (!CsvRowReader.TryParseDate(CsvRowReader.Col(cols, iDate), out var date)) continue;
            if (!CsvRowReader.TryParseAmount(CsvRowReader.Col(cols, iAmt), out var amt)) continue;

            list.Add(new PaymentTxn
            {
                Source    = "ob",
                Date      = date,
                Amount    = amt,
                Last4     = CsvRowReader.Last4(CsvRowReader.Col(cols, iCard)),
                PayType   = CsvRowReader.NullIfEmpty(CsvRowReader.Col(cols, iMeth)),
                Reference = CsvRowReader.NullIfEmpty(CsvRowReader.Col(cols, iBp)),
                SourceDoc = CsvRowReader.NullIfEmpty(CsvRowReader.Col(cols, iDoc)),
                HskRef    = ExtractHsk(CsvRowReader.Col(cols, iDesc)),
                RawKey    = $"ob|{i}|{date:yyyyMMdd}|{amt}",
            });
        }
        return list;
    }

    /// <summary>Pulls the HSK reference from the Description, preferring the invoice (HSK-SI) over the order (HSK-SO).</summary>
    private static string? ExtractHsk(string? description)
    {
        if (string.IsNullOrWhiteSpace(description)) return null;
        var hits = HskRx.Matches(description).Select(m => m.Value.ToUpperInvariant()).Distinct().ToList();
        if (hits.Count == 0) return null;
        return hits.FirstOrDefault(h => h.StartsWith("HSK-SI")) ?? hits[0];
    }

    /// <summary>True for OB payment methods that settle through the card processor (Heartland).</summary>
    public static bool IsCardMethod(string? payType) =>
        payType is not null &&
        (payType.Contains("card", StringComparison.OrdinalIgnoreCase) ||
         payType.Contains("credit", StringComparison.OrdinalIgnoreCase) ||
         payType.Contains("debit", StringComparison.OrdinalIgnoreCase));
}
