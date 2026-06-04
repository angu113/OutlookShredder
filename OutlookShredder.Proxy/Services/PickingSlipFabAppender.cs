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
        var descs = new List<string>();
        try
        {
            using var pig = PigDoc.Open(slipBytes);
            foreach (var line in ExtractLines(pig))
            {
                var m = FabRx.Match(line);
                if (!m.Success) continue;
                var desc = m.Groups[2].Value.Trim();
                if (desc.Length >= 3) descs.Add(desc);
            }
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

    /// <summary>Groups a page's words into text lines by baseline, top-to-bottom.</summary>
    private static IEnumerable<string> ExtractLines(PigDoc doc)
    {
        foreach (var page in doc.GetPages())
        {
            var rows = page.GetWords()
                .GroupBy(w => (int)Math.Round(w.BoundingBox.Bottom / 2.0))   // ~2pt baseline tolerance
                .OrderByDescending(g => g.Key);
            foreach (var row in rows)
                yield return string.Join(" ", row.OrderBy(w => w.BoundingBox.Left).Select(w => w.Text));
        }
    }
}
