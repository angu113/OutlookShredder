using Microsoft.Extensions.Logging;
using Microsoft.VisualBasic.FileIO;

namespace OutlookShredder.Proxy.Services;

/// <summary>
/// Parses ERP CSV exports for Business Partners and Contacts.
/// No SharePoint dependency — pure parsing and dedup logic.
/// </summary>
public sealed class CustomerImportService(ILogger<CustomerImportService> log)
{
    public sealed record BpRow(string Name, string PopupMessage);
    public sealed record ContactRow(string CustomerName, string ContactName, string Phone);
    /// <summary>A row that was filtered out during parsing and needs manual review.</summary>
    public sealed record SkippedItem(string Name, string Reason);
    public sealed record ParseResult<T>(
        IReadOnlyList<T>           Rows,
        IReadOnlyList<string>      Warnings,
        IReadOnlyList<SkippedItem> Skipped);

    // Phrases that flag a BP name as an ERP system artefact.
    // "Dupe" is short and may appear in real company names — all matches are written
    // to the per-run skipped review file so they can be inspected manually.
    private static readonly string[] InvalidPhrases =
        ["duplicate", "do not use", "dupe"];

    // ── Business Partners (ExportedData (2).csv style) ───────────────────────

    public ParseResult<BpRow> ParsePartners(string csv)
    {
        var warnings = new List<string>();
        var skipped  = new List<SkippedItem>();
        var rows     = new List<BpRow>();
        var lines    = ReadCsv(csv);

        if (lines.Count < 2)
        {
            warnings.Add("Partners CSV: no data rows found");
            return new(rows, warnings, skipped);
        }

        var hdr      = lines[0];
        int nameIdx  = ColIndex(hdr, "Name");
        int popupIdx = ColIndex(hdr, "Popup Message");

        if (nameIdx < 0 || popupIdx < 0)
        {
            warnings.Add($"Partners CSV: required columns not found. Header: {string.Join(" | ", hdr)}");
            return new(rows, warnings, skipped);
        }

        var seen = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

        for (int i = 1; i < lines.Count; i++)
        {
            var cols = lines[i];
            if (cols.Length <= Math.Max(nameIdx, popupIdx)) continue;

            var name = cols[nameIdx].Trim();
            if (string.IsNullOrWhiteSpace(name)) continue;

            var matchedPhrase = InvalidPhrases.FirstOrDefault(
                p => name.Contains(p, StringComparison.OrdinalIgnoreCase));

            if (matchedPhrase is not null)
            {
                var reason = $"name contains '{matchedPhrase}'";
                var msg    = $"Row {i + 1}: BP '{name}' skipped — {reason}";
                log.LogWarning("[CustImport] {Msg}", msg);
                warnings.Add(msg);
                skipped.Add(new SkippedItem(name, reason));
                continue;
            }

            if (!seen.Add(name))
            {
                var msg = $"Row {i + 1}: duplicate business partner '{name}' — skipped";
                log.LogWarning("[CustImport] {Msg}", msg);
                warnings.Add(msg);
                // Duplicate-within-file is not added to skipped — it's a data quality issue,
                // not an ambiguous phrase match that needs human review.
                continue;
            }

            rows.Add(new BpRow(name, cols[popupIdx].Trim()));
        }

        return new(rows, warnings, skipped);
    }

    // ── Contacts (ExportedData (1).csv style) ────────────────────────────────

    public ParseResult<ContactRow> ParseContacts(string csv)
    {
        var warnings = new List<string>();
        var skipped  = new List<SkippedItem>();
        var raw      = new List<(string Bp, string Contact, string Phone)>();
        var lines    = ReadCsv(csv);

        if (lines.Count < 2)
        {
            warnings.Add("Contacts CSV: no data rows found");
            return new([], warnings, skipped);
        }

        var hdr        = lines[0];
        int bpIdx      = ColIndex(hdr, "Customer Name");
        int contactIdx = ColIndex(hdr, "Contact Name");
        int phoneIdx   = ColIndex(hdr, "Phone");

        if (bpIdx < 0 || contactIdx < 0 || phoneIdx < 0)
        {
            warnings.Add($"Contacts CSV: required columns not found. Header: {string.Join(" | ", hdr)}");
            return new([], warnings, skipped);
        }

        for (int i = 1; i < lines.Count; i++)
        {
            var cols   = lines[i];
            int maxIdx = Math.Max(Math.Max(bpIdx, contactIdx), phoneIdx);
            if (cols.Length <= maxIdx) continue;

            var bp      = cols[bpIdx].Trim();
            var contact = cols[contactIdx].Trim();
            var rawPh   = cols[phoneIdx].Trim();

            if (string.IsNullOrWhiteSpace(bp) || string.IsNullOrWhiteSpace(contact)) continue;

            var phone = NormalizePhone(rawPh);
            if (phone is null)
            {
                if (!string.IsNullOrWhiteSpace(rawPh))
                {
                    var msg = $"Row {i + 1}: invalid phone '{rawPh}' for {bp}/{contact} — skipped";
                    log.LogWarning("[CustImport] {Msg}", msg);
                    warnings.Add(msg);
                    // Surface every phone-formatting reject for review so the bad ERP phone can be cleaned
                    // up — flows to the _skipped_review_*.csv file + the response's skippedForReview count.
                    skipped.Add(new SkippedItem($"{bp} / {contact}", $"unloadable phone format: '{rawPh}'"));
                }
                continue;
            }

            raw.Add((bp, contact, phone));
        }

        return new(Deduplicate(raw), warnings, skipped);
    }

    // ── Dedup: per-BP, prefer phones not shared across multiple contacts ─────

    private static List<ContactRow> Deduplicate(
        List<(string Bp, string Contact, string Phone)> raw)
    {
        var result = new List<ContactRow>();

        foreach (var bpGroup in raw.GroupBy(r => r.Bp, StringComparer.OrdinalIgnoreCase))
        {
            // Collect distinct phones per contact name within this BP
            var byContact = bpGroup
                .GroupBy(r => r.Contact, StringComparer.OrdinalIgnoreCase)
                .ToDictionary(
                    g => g.Key,
                    g => g.Select(r => r.Phone).Distinct().ToHashSet(),
                    StringComparer.OrdinalIgnoreCase);

            // Count how many distinct contacts each phone appears for (within this BP)
            var phoneContactCount = new Dictionary<string, int>(StringComparer.Ordinal);
            foreach (var phones in byContact.Values)
                foreach (var ph in phones)
                    phoneContactCount[ph] = phoneContactCount.GetValueOrDefault(ph) + 1;

            // Phones shared by 2+ contacts at this BP = company/general line
            var shared = phoneContactCount
                .Where(kv => kv.Value >= 2)
                .Select(kv => kv.Key)
                .ToHashSet(StringComparer.Ordinal);

            foreach (var (contact, phones) in byContact)
            {
                // Prefer the contact's own unique numbers; fall back to shared only if nothing else
                var preferred = phones.Except(shared).ToList();
                var keep      = preferred.Count > 0 ? preferred : [.. phones];

                foreach (var ph in keep)
                    result.Add(new ContactRow(bpGroup.Key, contact, ph));
            }
        }

        return result;
    }

    // ── Phone normalisation ──────────────────────────────────────────────────

    /// <summary>
    /// Strips all non-digit characters. Drops a leading country code '1' if 11 digits.
    /// Returns a 10-digit string, or null if the result is not exactly 10 digits.
    /// </summary>
    public static string? NormalizePhone(string? raw)
    {
        if (string.IsNullOrWhiteSpace(raw)) return null;
        var d = new string(raw.Where(char.IsDigit).ToArray());
        if (d.Length == 11 && d[0] == '1') d = d[1..];
        return d.Length == 10 ? d : null;
    }

    // ── CSV helpers ──────────────────────────────────────────────────────────

    private static int ColIndex(string[] hdr, string name) =>
        Array.FindIndex(hdr, h => h.Trim().Equals(name, StringComparison.OrdinalIgnoreCase));

    private static List<string[]> ReadCsv(string content)
    {
        var result = new List<string[]>();
        using var p = new TextFieldParser(new StringReader(content));
        p.TextFieldType = FieldType.Delimited;
        p.SetDelimiters(",");
        p.HasFieldsEnclosedInQuotes = true;
        p.TrimWhiteSpace = false;
        while (!p.EndOfData)
        {
            try   { var f = p.ReadFields(); if (f is not null) result.Add(f); }
            catch { /* skip malformed lines */ }
        }
        return result;
    }
}
