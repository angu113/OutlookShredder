using System.Collections.Concurrent;
using System.Text.RegularExpressions;
using Microsoft.Extensions.Logging;
using PdfSharp.Pdf;
using PdfSharp.Pdf.IO;
using OutlookShredder.Proxy.Services.Drawing;
using PigDoc = UglyToad.PdfPig.PdfDocument;

namespace OutlookShredder.Proxy.Services;

/// <summary>
/// Display-time enrichment: scans a picking slip for <c>FAB:</c> shop notes, turns the canonical
/// text after the anchor into a dimensioned drawing (the same engine as the Drawing tab), and
/// appends each drawing as a page to the slip. Replaces the old keyword-based callout page.
///
/// Generated drawings are cached in-memory keyed by their canonical text, so a slip viewed
/// repeatedly — or two slips with the same FAB note — re-use one render.
/// </summary>
internal static class PickingSlipFabAppender
{
    // "FAB: (2) U 4 x 2 x 36, 16ga CRS, finish outside"  ->  qty=2, desc="U 4 x 2 ..."
    private static readonly Regex FabRx =
        new(@"FAB:\s*(?:\((\d+)\)\s*)?(.+)$", RegexOptions.IgnoreCase | RegexOptions.Compiled);

    private static readonly ConcurrentDictionary<string, byte[]> _cache =
        new(StringComparer.OrdinalIgnoreCase);

    /// <summary>
    /// Returns the slip with a drawing page appended per distinct FAB note. On any failure
    /// (no notes, unparseable text, render/append error) the original bytes are returned unchanged.
    /// </summary>
    public static byte[] AppendFabDrawings(byte[] slipBytes, ILogger? log = null)
    {
        List<string> descs;
        try
        {
            using var pig = PigDoc.Open(slipBytes);
            descs = ExtractFabDescs(ExtractRows(pig));
        }
        catch (Exception ex)
        {
            log?.LogWarning(ex, "[FAB] text scan failed — returning slip unchanged");
            return slipBytes;
        }

        var distinct = descs.Distinct(StringComparer.OrdinalIgnoreCase).ToList();
        if (distinct.Count == 0) return slipBytes;

        PickingSlipEnricher.EnsureFontResolver();   // drawings embed Arial, same as enrichment
        var drawings = distinct.Select(d => RenderFab(d, log)).Where(b => b is not null).Cast<byte[]>().ToList();
        if (drawings.Count == 0) return slipBytes;

        try
        {
            using var ms = new MemoryStream();
            ms.Write(slipBytes, 0, slipBytes.Length);
            ms.Position = 0;
            using var outDoc = PdfReader.Open(ms, PdfDocumentOpenMode.Modify);
            foreach (var d in drawings)
            {
                using var dms = new MemoryStream(d);
                using var dDoc = PdfReader.Open(dms, PdfDocumentOpenMode.Import);
                for (int i = 0; i < dDoc.PageCount; i++)
                {
                    var added = outDoc.AddPage(dDoc.Pages[i]);
                    // Drawings are landscape; rotate 90° so they fill the portrait picking-slip page.
                    added.Rotate = (added.Rotate + 90) % 360;
                }
            }
            using var outMs = new MemoryStream();
            outDoc.Save(outMs);
            log?.LogInformation("[FAB] appended {N} drawing page(s) to picking slip", drawings.Count);
            return outMs.ToArray();
        }
        catch (Exception ex)
        {
            log?.LogWarning(ex, "[FAB] append failed — returning slip unchanged");
            return slipBytes;
        }
    }

    private static byte[]? RenderFab(string desc, ILogger? log)
    {
        if (_cache.TryGetValue(desc, out var cached)) return cached;
        try
        {
            var spec = DrawingTextParser.Parse(desc);
            var fp = FlatPattern.Develop(spec);
            var pdf = DrawingPdfRenderer.Render(fp);
            _cache[desc] = pdf;
            return pdf;
        }
        catch (Exception ex)
        {
            log?.LogWarning(ex, "[FAB] could not render note '{Desc}'", desc);
            return null;
        }
    }

    /// <summary>
    /// Pulls FAB note descriptions out of the page rows, stitching wrapped continuations back on.
    /// The FAB cell often wraps in the narrow Product column (e.g. "… finish" on one row, "outside"
    /// on the next); a continuation stays in the FAB cell's column (left edge ≳ the FAB row's) and
    /// isn't a new FAB note, a blank, or a new line-item (which reaches the far-left MSPC column).
    /// </summary>
    internal static List<string> ExtractFabDescs(List<(string Text, double Left)> rows)
    {
        var descs = new List<string>();
        for (int i = 0; i < rows.Count; i++)
        {
            var m = FabRx.Match(rows[i].Text);
            if (!m.Success) continue;
            var desc = m.Groups[2].Value.Trim();

            double fabLeft = rows[i].Left;
            for (int j = i + 1; j < rows.Count; j++)
            {
                var next = rows[j];
                if (string.IsNullOrWhiteSpace(next.Text)) break;
                if (FabRx.IsMatch(next.Text)) break;
                if (next.Left < fabLeft - 12.0) break;   // reaches a left column (MSPC) → new line-item
                desc += " " + next.Text.Trim();
                i = j;                                    // consume the continuation
            }

            if (desc.Length >= 3) descs.Add(desc);
        }
        return descs;
    }

    /// <summary>Groups every page's words into text rows by baseline (top-to-bottom), each with the
    /// row's left edge so wrapped FAB continuations can be stitched back on by column.</summary>
    private static List<(string Text, double Left)> ExtractRows(PigDoc doc)
    {
        var result = new List<(string, double)>();
        foreach (var page in doc.GetPages())
        {
            var rows = page.GetWords()
                .GroupBy(w => (int)Math.Round(w.BoundingBox.Bottom / 2.0))   // ~2pt baseline tolerance
                .OrderByDescending(g => g.Key);
            foreach (var row in rows)
            {
                var ordered = row.OrderBy(w => w.BoundingBox.Left).ToList();
                result.Add((string.Join(" ", ordered.Select(w => w.Text)), ordered.Min(w => w.BoundingBox.Left)));
            }
        }
        return result;
    }
}
