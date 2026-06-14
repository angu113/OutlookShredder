using OutlookShredder.Proxy.Models;

namespace OutlookShredder.Proxy.Services;

/// <summary>
/// Header-alias map for the Heartland / Global Payments exports, bound from config
/// (<c>ShadowRecon:Heartland:Columns</c>). Heartland exposes TWO reporting methods (two CSV layouts);
/// the alias arrays are the UNION across both so a single parser handles either — whichever header is
/// present wins (<see cref="CsvRowReader.IndexOfAny"/>). The exact headers / amount-sign / refund
/// convention are UNKNOWN until we capture real exports — defaults below are best-guess placeholders.
/// Finalizing these (and whether the second method is transaction-level vs batch/deposit-level) after
/// the first real capture is the explicit first build step, and it is a config edit, not a code change.
/// </summary>
public sealed class HeartlandColumnMap
{
    // Defaults match the real Heartland transaction-detail export (the "Merchant Batch Download" CSV):
    //   ...,BatchNumber,CardType,CardNum,...,SaleReturn,Trandate,AuthNumber,...,Close_date,Amount,Textbox17
    // "Closed" is the chosen anchor date (Close_date), with Trandate as fallback. Generic aliases are
    // kept too so a differently-shaped export still maps. Amount uses parentheses for returns (signed).
    public string[] Date      { get; set; } = ["Close_date", "Trandate", "Transaction Date", "Settlement Date", "Date"];
    public string[] Amount    { get; set; } = ["Amount", "Transaction Amount", "Net Amount", "Total Amount"];
    public string[] Last4     { get; set; } = ["CardNum", "Card Number", "Last 4", "Account Number"];
    public string[] AuthCode  { get; set; } = ["AuthNumber", "Authorization Code", "Auth Code", "Approval Code"];
    public string[] TxnId     { get; set; } = ["Transaction ID", "Reference Number", "ARN", "Trans ID"];
    public string[] CardType  { get; set; } = ["CardType", "Card Type", "Card Brand", "Network"];
    public string[] BatchId   { get; set; } = ["BatchNumber", "Batch", "Batch Number", "Batch ID"];
    public string[] Reference { get; set; } = ["Reference", "Invoice", "Customer Reference", "Purchase ID"];

    /// <summary>When true, the parsed amount sign is trusted (returns already negative / parenthesised).</summary>
    public bool RefundsAreNegative { get; set; } = true;
}

/// <summary>
/// Parses a Heartland / Global Payments transaction (or batch/deposit) export into
/// <see cref="PaymentTxn"/>s using a column-alias map that spans both Heartland report formats.
/// Amounts kept SIGNED. See <see cref="HeartlandColumnMap"/> re: finalizing aliases after real capture.
/// </summary>
public static class HeartlandCsvParser
{
    public static List<PaymentTxn> Parse(string csvContent, HeartlandColumnMap map)
    {
        var records = CsvRowReader.ParseRecords(csvContent);
        if (records.Count < 2) return [];

        var headers = records[0];
        int iDate = CsvRowReader.IndexOfAny(headers, map.Date);
        int iAmt  = CsvRowReader.IndexOfAny(headers, map.Amount);
        int iL4   = CsvRowReader.IndexOfAny(headers, map.Last4);
        int iAuth = CsvRowReader.IndexOfAny(headers, map.AuthCode);
        int iTxn  = CsvRowReader.IndexOfAny(headers, map.TxnId);
        int iCard = CsvRowReader.IndexOfAny(headers, map.CardType);
        int iBat  = CsvRowReader.IndexOfAny(headers, map.BatchId);
        int iRef  = CsvRowReader.IndexOfAny(headers, map.Reference);

        if (iDate < 0 || iAmt < 0)
            throw new Exception(
                "Heartland CSV missing required Date/Amount columns " +
                $"(found: {string.Join(", ", headers)}). Adjust ShadowRecon:Heartland:Columns aliases.");

        var list = new List<PaymentTxn>();
        for (int i = 1; i < records.Count; i++)
        {
            var cols = records[i];
            if (!CsvRowReader.TryParseDate(CsvRowReader.Col(cols, iDate), out var date)) continue;
            if (!CsvRowReader.TryParseAmount(CsvRowReader.Col(cols, iAmt), out var amt)) continue;

            list.Add(new PaymentTxn
            {
                Source    = "processor",
                Date      = date,
                Amount    = amt,
                Last4     = CsvRowReader.Last4(CsvRowReader.Col(cols, iL4)),
                AuthCode  = CsvRowReader.NullIfEmpty(CsvRowReader.Col(cols, iAuth)),
                TxnId     = CsvRowReader.NullIfEmpty(CsvRowReader.Col(cols, iTxn)),
                CardType  = CsvRowReader.NullIfEmpty(CsvRowReader.Col(cols, iCard)),
                BatchId   = CsvRowReader.NullIfEmpty(CsvRowReader.Col(cols, iBat)),
                Reference = CsvRowReader.NullIfEmpty(CsvRowReader.Col(cols, iRef)),
                RawKey    = $"proc|{i}|{date:yyyyMMdd}|{amt}",
            });
        }
        return list;
    }
}
