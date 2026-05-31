using System.Text.RegularExpressions;

namespace OutlookShredder.Proxy.Services;

/// <summary>
/// Pulls strong business identifiers out of an email (subject + body) so conversations that share one
/// can be suggested as a "project" (Layer 2). Deliberately broad — noise is filtered downstream because
/// a suggestion only fires when the SAME token appears in two or more distinct conversations, which
/// random tokens rarely do.
///
/// Captures:
///   • Booking / BL / container / order numbers — 2–6 letters + 6+ digits (e.g. QGD2700416, TAOSE2604049,
///     CMAU7933240). One pattern covers ISO-6346 containers and most carrier/booking refs.
///   • Our PO/SO refs — HSK-PO…/HSK-SO….
///   • Labelled values — the token after MBL/HBL/BL/CTNR/Booking/Container/Quote/Ref/Order/PO/SO
///     (catches dashed/odd values the bare pattern misses).
/// </summary>
public static partial class MailReferenceExtractor
{
    [GeneratedRegex(@"\b[A-Z]{2,6}\d{6,}\b")]                              private static partial Regex StrongRx();
    [GeneratedRegex(@"\bHSK-(?:PO|SO)\d+\b")]                             private static partial Regex HskRx();
    [GeneratedRegex(@"\b(?:MBL|HBL|BL|CTNR|CNTR|CONTAINER|BOOKING|QUOTE|REF|ORDER|PO|SO)\b\s*[:#]?\s*([A-Z0-9][A-Z0-9\-]{4,})")]
    private static partial Regex LabeledRx();

    public static List<string> Extract(string? subject, string? body)
    {
        var text = ((subject ?? "") + "\n" + (body ?? "")).ToUpperInvariant();
        if (text.Length > 20000) text = text[..20000];   // cap — refs live up top, avoid scanning huge bodies

        var refs = new HashSet<string>(StringComparer.Ordinal);
        foreach (Match m in StrongRx().Matches(text)) refs.Add(m.Value);
        foreach (Match m in HskRx().Matches(text))    refs.Add(m.Value);
        foreach (Match m in LabeledRx().Matches(text))
        {
            var v = m.Groups[1].Value.Trim('-');
            if (v.Length >= 5 && v.Any(char.IsDigit)) refs.Add(v);
        }
        return refs.ToList();
    }
}
