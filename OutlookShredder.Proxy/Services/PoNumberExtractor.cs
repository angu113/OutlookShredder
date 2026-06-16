using System.Text;
using System.Text.RegularExpressions;
using UglyToad.PdfPig;

namespace OutlookShredder.Proxy.Services;

/// <summary>
/// Reads the HSK-PO document number from a "Purchase Order - Material" PDF body when the filename
/// carries none (a bare "PurchaseOrder.pdf"). OpenBravo prints the PO number immediately under the
/// title, e.g. "PURCHASE ORDER - MATERIAL HSK-PO1006243…", so a regex on the extracted text is exact
/// and beats the AI for this fixed-format field.
/// </summary>
public static class PoNumberExtractor
{
    // HSK-PO with any/no separator and optional spacing; >=5 digits (PO numbers are 7).
    private static readonly Regex PoRx = new(@"HSK[\s_\-]*PO[\s]*(\d{5,})",
        RegexOptions.IgnoreCase | RegexOptions.Compiled);

    /// <summary>Canonical "HSK-PO#######" (first occurrence — the header), or null. Pure (unit-tested).</summary>
    public static string? FromText(string? text)
    {
        if (string.IsNullOrEmpty(text)) return null;
        var m = PoRx.Match(text);
        return m.Success ? "HSK-PO" + m.Groups[1].Value : null;
    }

    /// <summary>Extracts the PO PDF's text via PdfPig then returns its PO number. Never throws —
    /// a parse failure yields null (logged), so recognition simply falls back to its other sources.</summary>
    public static string? FromPdf(byte[] pdfBytes, ILogger? log = null)
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
            log?.LogWarning(ex, "[PO-NUM] PDF text extraction failed");
            return null;
        }
    }
}
