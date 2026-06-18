using System.Globalization;
using Microsoft.Extensions.Logging;
using Microsoft.VisualBasic.FileIO;

namespace OutlookShredder.Proxy.Services;

/// <summary>
/// Parses ERP CSV exports for Business Partners and Contacts.
/// No SharePoint dependency — pure parsing and dedup logic.
/// </summary>
public sealed class CustomerImportService(ILogger<CustomerImportService> log)
{
    public sealed record BpRow(string Name, string PopupMessage, bool Active = true);

    /// <summary>Parses an ERP "Active" cell. Accepts true/false, yes/no, y/n, 1/0 (any case).
    /// Blank / missing / unrecognised ⇒ <c>true</c> (active) — we never flip a customer inactive from
    /// absent data; only an explicit negative deactivates.</summary>
    internal static bool ParseActiveFlag(string? raw) =>
        (raw?.Trim().ToLowerInvariant()) switch
        {
            "false" or "no" or "n" or "0" or "inactive" => false,
            _                                           => true,
        };
    public sealed record ContactRow(string CustomerName, string ContactName, string Phone);
    /// <summary>One row of the rich "Customer Info" master export. <see cref="Fields"/> maps a
    /// SharePoint internal column name (from <see cref="CustomerInfoSchema"/>) to the raw CSV cell
    /// value; only columns that carried a non-empty value are present.</summary>
    public sealed record CustomerInfoRow(string Name, IReadOnlyDictionary<string, string?> Fields);
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
        int activeIdx = ColIndex(hdr, "Active");   // optional — older partner exports omit it (=> all active)

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

            var active = activeIdx < 0 || activeIdx >= cols.Length
                ? true                                        // column absent ⇒ active
                : ParseActiveFlag(cols[activeIdx]);
            rows.Add(new BpRow(name, cols[popupIdx].Trim(), active));
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

    // ── Customer Info (ExportedData (5).csv style — the rich BP master) ───────

    /// <summary>
    /// Parses the "Customer Info" master export (Business Partner + ~22 enrichment columns:
    /// Payment Terms, On Hold, credit limits, sales/margin stats, etc.). The Business Partner
    /// name is the cross-file match key (same key the partner load writes to the Customers list).
    /// Dedups by name (first kept); blank-name rows are dropped silently (export footer/totals);
    /// duplicate names and unparseable numeric/boolean cells are surfaced for the review file.
    /// </summary>
    public ParseResult<CustomerInfoRow> ParseCustomerInfo(string csv)
    {
        var warnings = new List<string>();
        var skipped  = new List<SkippedItem>();
        var rows     = new List<CustomerInfoRow>();
        var lines    = ReadCsv(csv);

        if (lines.Count < 2)
        {
            warnings.Add("Customer Info CSV: no data rows found");
            return new(rows, warnings, skipped);
        }

        var hdr     = lines[0];
        int nameIdx = ColIndex(hdr, CustomerInfoSchema.NameCsvHeader);
        if (nameIdx < 0)
        {
            warnings.Add($"Customer Info CSV: '{CustomerInfoSchema.NameCsvHeader}' column not found. Header: {string.Join(" | ", hdr)}");
            return new(rows, warnings, skipped);
        }

        // Resolve each enrichment column's CSV index once. Missing columns are tolerated (logged) so
        // an export that drops/renames a field still loads everything else.
        var colIdx = new List<(CustomerInfoSchema.Col Col, int Idx)>();
        foreach (var c in CustomerInfoSchema.Columns)
        {
            int idx = ColIndex(hdr, c.Csv);
            if (idx < 0) warnings.Add($"Customer Info CSV: optional column '{c.Csv}' not present — skipped");
            colIdx.Add((c, idx));
        }

        var seen = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        for (int i = 1; i < lines.Count; i++)
        {
            var cols = lines[i];
            if (cols.Length <= nameIdx) continue;

            var name = cols[nameIdx].Trim();
            if (string.IsNullOrWhiteSpace(name)) continue;   // export footer / total rows — drop quietly

            if (!seen.Add(name))
            {
                var reason = "duplicate Business Partner in Customer Info file (first kept)";
                var msg    = $"Row {i + 1}: {reason} — '{name}'";
                log.LogWarning("[CustImport] {Msg}", msg);
                warnings.Add(msg);
                skipped.Add(new SkippedItem(name, reason));
                continue;
            }

            var fields = new Dictionary<string, string?>(StringComparer.Ordinal);
            foreach (var (col, idx) in colIdx)
            {
                if (idx < 0 || idx >= cols.Length) continue;
                var raw = cols[idx].Trim();
                if (raw.Length == 0) continue;

                // Surface (but don't drop the row over) numeric/boolean cells we can't parse —
                // the field is left unset; everything else on the row still loads.
                if (col.Kind != CustomerInfoSchema.Kind.Text &&
                    CustomerInfoSchema.ToTyped(raw, col.Kind) is null)
                {
                    skipped.Add(new SkippedItem(name, $"unparseable {col.Kind} for '{col.Csv}': '{raw}'"));
                    continue;
                }

                fields[col.Sp] = raw;
            }

            rows.Add(new CustomerInfoRow(name, fields));
        }

        return new(rows, warnings, skipped);
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

/// <summary>
/// The single source of truth for the "Customer Info" enrichment columns: maps each CSV header to a
/// SharePoint internal column name + storage kind. Drives all four touch points so they never drift:
/// CSV parsing (<see cref="CustomerImportService.ParseCustomerInfo"/>), SP list provisioning
/// (<c>EnsureCustomerListsAsync</c>), change detection (<see cref="Canon"/>), and writes
/// (<see cref="ToTyped"/>). SP internal names are deliberately clean identifiers (no spaces / % / +-/?)
/// so Graph accepts them as column names. The Business Partner name is the match key (the Customers
/// list <c>Title</c>) and is intentionally NOT in <see cref="Columns"/>.
/// </summary>
public static class CustomerInfoSchema
{
    public enum Kind { Text, Number, Boolean }
    /// <summary><paramref name="Write"/>=false means the column is provisioned + parsed (so it can be
    /// read, e.g. for the inactive-candidate filter) but the Customer Info enrichment never WRITES it —
    /// used for <c>Active</c>, which is owned by the Business Partner (partners) load.</summary>
    public sealed record Col(string Csv, string Sp, Kind Kind, bool Write = true);

    public const string NameCsvHeader = "Business Partner";

    public static readonly Col[] Columns =
    [
        // Active is owned by the partners (BP master) load — provisioned + parsed here (for IsInactive),
        // but NOT written by enrichment, so the last-running customerinfo load can't override it.
        new("Active",                     "Active",             Kind.Boolean, Write: false),
        new("Business Partner Category",  "BpCategory",         Kind.Text),
        new("Payment Terms",              "PaymentTerms",       Kind.Text),
        new("On Hold",                    "OnHold",             Kind.Boolean),
        new("Payment Method",             "PaymentMethod",      Kind.Text),
        new("Contact",                    "PrimaryContact",     Kind.Text),
        new("Phone Number",               "ContactPhone",       Kind.Text),
        new("Email",                      "ContactEmail",       Kind.Text),
        new("Credit Line Limit",          "CreditLineLimit",    Kind.Number),
        new("Credit Limit +/-",           "CreditAvailable",    Kind.Number),
        new("Auto Invoice",               "AutoInvoice",        Kind.Boolean),
        new("Auto Statement",             "AutoStatement",      Kind.Boolean),
        new("Current",                    "CurrentBalance",     Kind.Number),
        new("Win %",                      "WinPct",             Kind.Number),
        new("Win % Transactions",         "WinPctTransactions", Kind.Number),
        new("Sales Last 6 Months",        "SalesLast6Mo",       Kind.Number),
        new("Sales Last 12 Months",       "SalesLast12Mo",      Kind.Number),
        new("Average Invoice Value",      "AvgInvoiceValue",    Kind.Number),
        new("Average Margin Invoice",     "AvgMarginPct",       Kind.Number),
        new("Tax Exempt",                 "TaxExempt",          Kind.Boolean),
        new("How Did You Hear About Us?", "HowDidYouHear",      Kind.Text),
        new("Margin Type",                "MarginType",         Kind.Text),
    ];

    /// <summary>Timestamp column stamped on each enrichment write (so a stale record is visible).</summary>
    public const string UpdatedAtColumn = "CustomerInfoUpdatedAt";

    /// <summary>The ERP "Active" flag column — a logical-deletion marker: false/NO = inactive (hidden in
    /// ERP views). Existing records are still enriched (and marked active/inactive); but an *unmatched*
    /// inactive row is dead data and must NOT be surfaced as a candidate to add.</summary>
    public const string ActiveColumn = "Active";

    /// <summary>True when the row's Active flag is explicitly false. Blank / missing / true ⇒ treated as active.</summary>
    public static bool IsInactive(IReadOnlyDictionary<string, string?> fields) =>
        fields.TryGetValue(ActiveColumn, out var a) && Canon(a, Kind.Boolean) == "false";

    /// <summary>SharePoint column kind string consumed by <c>EnsureListColumnsAsync</c>.</summary>
    public static string SpType(Kind k) => k switch
    {
        Kind.Number  => "number",
        Kind.Boolean => "boolean",
        _            => "text",
    };

    /// <summary>Canonical comparable form, so "only write when changed" ignores formatting noise
    /// (e.g. "150000" vs "150000.0", "True" vs "true"). Blank → "".</summary>
    public static string Canon(string? raw, Kind kind)
    {
        if (string.IsNullOrWhiteSpace(raw)) return "";
        var s = raw.Trim();
        return kind switch
        {
            Kind.Boolean => bool.TryParse(s, out var b) ? (b ? "true" : "false") : s.ToLowerInvariant(),
            // "0.####" strips trailing-zero scale so "150000" == "150000.0" and 50.49697… rounds to 50.497.
            Kind.Number  => decimal.TryParse(s, NumberStyles.Any, CultureInfo.InvariantCulture, out var d)
                              ? Math.Round(d, 4).ToString("0.####", CultureInfo.InvariantCulture)
                              : s,
            _            => s,
        };
    }

    /// <summary>The typed value to store in SharePoint, or null when blank / unparseable.</summary>
    public static object? ToTyped(string? raw, Kind kind)
    {
        if (string.IsNullOrWhiteSpace(raw)) return null;
        var s = raw.Trim();
        return kind switch
        {
            Kind.Boolean => bool.TryParse(s, out var b) ? b : null,
            Kind.Number  => decimal.TryParse(s, NumberStyles.Any, CultureInfo.InvariantCulture, out var d)
                              ? (double)Math.Round(d, 4) : null,
            _            => (object?)s,
        };
    }
}
