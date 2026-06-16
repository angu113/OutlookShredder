using System.Text;
using System.Text.RegularExpressions;
using UglyToad.PdfPig;

namespace OutlookShredder.Proxy.Services;

/// <summary>
/// Deterministically pulls the customer sales-order numbers (HSK-SO…) that OpenBravo prints in the
/// "Customer / Sales Order #" column of a procure-to-order Purchase Order PDF. Regex, not AI: the
/// column format is fixed (HSK-SO + 7 digits), so extraction is cheap and exact. Stock-replenishment
/// POs have no such column and yield an empty list (a meaningful signal, not a failure).
/// </summary>
public static class PoSalesOrderExtractor
{
    // HSK-SO with any/no separator and optional spacing; >=5 digits to avoid false positives.
    private static readonly Regex SoRx = new(@"HSK[\s_\-]*SO[\s]*(\d{5,})",
        RegexOptions.IgnoreCase | RegexOptions.Compiled);

    /// <summary>Distinct canonical "HSK-SO#######" numbers in first-seen order. Pure (unit-tested).</summary>
    public static IReadOnlyList<string> FromText(string? text)
    {
        if (string.IsNullOrEmpty(text)) return [];
        var seen   = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        var result = new List<string>();
        foreach (Match m in SoRx.Matches(text))
        {
            var so = "HSK-SO" + m.Groups[1].Value;
            if (seen.Add(so)) result.Add(so);
        }
        return result;
    }

    /// <summary>Extracts the PO PDF's text via PdfPig then returns its sales-order numbers. Never throws —
    /// a parse failure yields an empty list (logged), so PO ingestion is never blocked by SO capture.</summary>
    public static IReadOnlyList<string> FromPdf(byte[] pdfBytes, ILogger? log = null)
    {
        try
        {
            using var doc = PdfDocument.Open(pdfBytes);
            var sb = new StringBuilder();
            foreach (var p in doc.GetPages()) sb.Append(p.Text).Append('\n');
            return FromText(sb.ToString());
        }
        catch (Exception ex)
        {
            log?.LogWarning(ex, "[PO-SO] PDF text extraction failed");
            return [];
        }
    }
}
