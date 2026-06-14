using System.Globalization;
using System.Text;

namespace OutlookShredder.Proxy.Services;

/// <summary>
/// Shared RFC-4180 CSV helpers used by the statement / OB-payments / Elavon parsers.
/// Extracted from StatementCsvParser so the three parsers share one implementation rather
/// than drifting copies.
/// </summary>
public static class CsvRowReader
{
    /// <summary>Splits content into raw lines, stripping a leading UTF-8 BOM (U+FEFF) if present.</summary>
    public static string[] SplitLines(string content)
    {
        if (!string.IsNullOrEmpty(content)) content = content.TrimStart('﻿');
        return content.Split('\n');
    }

    /// <summary>First header index equal (case-insensitive, trimmed) to <paramref name="name"/>, else -1.</summary>
    public static int IndexOf(string[] headers, string name) =>
        Array.FindIndex(headers, h => h.Trim().Equals(name, StringComparison.OrdinalIgnoreCase));

    /// <summary>First header index matching ANY of the alias names (case-insensitive, trimmed), else -1.</summary>
    public static int IndexOfAny(string[] headers, IEnumerable<string> names)
    {
        foreach (var n in names)
        {
            int idx = IndexOf(headers, n);
            if (idx >= 0) return idx;
        }
        return -1;
    }

    /// <summary>Safe column access — empty string when the index is out of range or negative.</summary>
    public static string Col(string[] cols, int idx) =>
        idx >= 0 && idx < cols.Length ? cols[idx] : "";

    public static string? NullIfEmpty(string? s) =>
        string.IsNullOrWhiteSpace(s) ? null : s.Trim();

    /// <summary>Digits only (e.g. card "**** 1234" -> "1234"); returns the last 4 when longer. Null when none.</summary>
    public static string? Last4(string? s)
    {
        if (string.IsNullOrEmpty(s)) return null;
        var digits = new string(s.Where(char.IsDigit).ToArray());
        if (digits.Length == 0) return null;
        return digits.Length > 4 ? digits[^4..] : digits;
    }

    /// <summary>
    /// Parses a money value, culture-invariant: strips currency symbols / thousands separators /
    /// whitespace, and treats parentheses as negative (e.g. "($12.34)" -> -12.34). Keeps the sign.
    /// </summary>
    public static bool TryParseAmount(string? s, out decimal amount)
    {
        amount = 0m;
        if (string.IsNullOrWhiteSpace(s)) return false;
        var t = s.Trim();
        bool parenNeg = t.StartsWith('(') && t.EndsWith(')');
        if (parenNeg) t = t[1..^1];
        var sb = new StringBuilder(t.Length);
        foreach (var ch in t)
        {
            if (char.IsDigit(ch) || ch == '.' || ch == '-') sb.Append(ch);
            // drop $, commas, spaces, currency codes, etc.
        }
        if (sb.Length == 0) return false;
        if (!decimal.TryParse(sb.ToString(), NumberStyles.Any, CultureInfo.InvariantCulture, out amount))
            return false;
        if (parenNeg) amount = -amount;
        return true;
    }

    private static readonly string[] DateFormats =
        ["MM/dd/yyyy", "M/d/yyyy", "MM-dd-yyyy", "M-d-yyyy", "yyyy-MM-dd", "MM/dd/yyyy HH:mm:ss"];

    /// <summary>Date parse tolerant of US m/d and m-d layouts (OB/Heartland) and ISO; invariant culture.</summary>
    public static bool TryParseDate(string? s, out DateTime date)
    {
        date = default;
        if (string.IsNullOrWhiteSpace(s)) return false;
        s = s.Trim();
        if (DateTime.TryParse(s, CultureInfo.InvariantCulture, DateTimeStyles.None, out date)) { date = date.Date; return true; }
        if (DateTime.TryParseExact(s, DateFormats, CultureInfo.InvariantCulture, DateTimeStyles.None, out date)) { date = date.Date; return true; }
        return false;
    }

    /// <summary>
    /// Full RFC-4180 parser over the whole content: splits into records on newlines that are OUTSIDE
    /// quotes, so quoted fields may contain commas, escaped quotes, AND embedded newlines (the OB
    /// "ExportedData" payment export has multi-line Description cells). Strips a leading BOM.
    /// </summary>
    public static List<string[]> ParseRecords(string content)
    {
        var records = new List<string[]>();
        if (string.IsNullOrEmpty(content)) return records;
        if (content[0] == '﻿') content = content[1..];

        var field = new StringBuilder();
        var row = new List<string>();
        bool inQuotes = false, rowHasData = false;
        int i = 0;
        while (i < content.Length)
        {
            char c = content[i];
            if (inQuotes)
            {
                if (c == '"' && i + 1 < content.Length && content[i + 1] == '"') { field.Append('"'); i += 2; continue; }
                if (c == '"') { inQuotes = false; i++; continue; }
                field.Append(c); i++; continue;
            }
            switch (c)
            {
                case '"': inQuotes = true; rowHasData = true; i++; break;
                case ',': row.Add(field.ToString()); field.Clear(); rowHasData = true; i++; break;
                case '\r': i++; break;
                case '\n':
                    row.Add(field.ToString()); field.Clear();
                    if (rowHasData) records.Add([.. row]);
                    row.Clear(); rowHasData = false; i++;
                    break;
                default: field.Append(c); rowHasData = true; i++; break;
            }
        }
        if (rowHasData || field.Length > 0) { row.Add(field.ToString()); records.Add([.. row]); }
        return records;
    }

    /// <summary>RFC 4180 row parser — handles quoted fields containing commas and escaped quotes.</summary>
    public static string[] ParseRow(string line)
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
