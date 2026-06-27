using System.Collections.Concurrent;
using System.Text.RegularExpressions;
using Microsoft.Extensions.Logging;
using PdfSharp.Drawing;
using PdfSharp.Pdf;
using PdfSharp.Pdf.IO;
using OutlookShredder.Proxy.Services.Drawing;
using PigDoc = UglyToad.PdfPig.PdfDocument;

namespace OutlookShredder.Proxy.Services;

/// <summary>A FAB note's order quantity (the <c>(N)</c> prefix, default 1) plus its description text.</summary>
public sealed record FabNote(int Qty, string Desc);

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

    private static readonly Regex _wsRx = new(@"\s+", RegexOptions.Compiled);

    private static readonly ConcurrentDictionary<string, byte[]> _cache =
        new(StringComparer.OrdinalIgnoreCase);

    /// <summary>
    /// Returns the slip with a drawing page appended per distinct FAB note. On any failure
    /// (no notes, unparseable text, render/append error) the original bytes are returned unchanged.
    /// When <paramref name="customerName"/> is supplied it is passed to the renderer so column
    /// drawings receive their fab-context BOM header.
    /// </summary>
    public static byte[] AppendFabDrawings(byte[] slipBytes, ILogger? log = null, string? customerName = null)
    {
        var distinct = GetFabDescs(slipBytes, log);
        if (distinct.Count == 0) return slipBytes;

        PickingSlipEnricher.EnsureFontResolver();   // drawings embed Arial, same as enrichment
        // Letter each page (A, B, C…) by deduped order. The SAME letter is stamped beside the FAB: note
        // in the slip body, so a part's note and its drawing page share one letter. Map slug -> letter.
        var slugToLetter = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        for (int i = 0; i < distinct.Count; i++)
        {
            var slug = DevelopSlug(distinct[i]);
            if (!string.IsNullOrEmpty(slug)) slugToLetter.TryAdd(slug!, FabLetter(i));
        }

        var drawings = distinct.Select((d, i) => RenderFab(d, log, customerName, FabLetter(i)))
            .Where(b => b is not null).Cast<byte[]>().ToList();
        if (drawings.Count == 0) return slipBytes;

        // Where each FAB: note sits in the slip body — for the in-place letter stamps.
        List<FabAnchor> anchors;
        try { using var pig = PigDoc.Open(slipBytes); anchors = ExtractFabAnchors(pig); }
        catch { anchors = new List<FabAnchor>(); }

        try
        {
            using var ms = new MemoryStream();
            ms.Write(slipBytes, 0, slipBytes.Length);
            ms.Position = 0;
            using var outDoc = PdfReader.Open(ms, PdfDocumentOpenMode.Modify);
            // Stamp the item letter beside each FAB: note BEFORE appending drawings, so the anchor page
            // indices still line up with the slip's original pages.
            StampFabLetters(outDoc, anchors, slugToLetter, log);
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

    /// <summary>
    /// Scans a slip and returns its distinct FAB note descriptions, deduped the same way the drawing
    /// append uses (bracket/column stitch → whitespace-norm → by-part-slug). Returns an empty list on
    /// any read failure or when the slip carries no FAB notes. Shared by the drawing-append path and
    /// the DXF-generation endpoint so both see exactly one entry per distinct part.
    /// </summary>
    public static List<string> GetFabDescs(byte[] slipBytes, ILogger? log = null)
    {
        try
        {
            using var pig = PigDoc.Open(slipBytes);
            var descs = ExtractFabDescs(ExtractRows(pig));
            return DedupeBySlug(DedupeDescs(descs), log);
        }
        catch (Exception ex)
        {
            log?.LogWarning(ex, "[FAB] text scan failed");
            return new List<string>();
        }
    }

    /// <summary>
    /// As <see cref="GetFabDescs"/>, but each surviving note keeps its order quantity (the <c>(N)</c>
    /// prefix). Used by the DXF endpoint so each part can be labelled "xN". Empty on read failure / no
    /// notes.
    /// </summary>
    public static List<FabNote> GetFabNotes(byte[] slipBytes, ILogger? log = null)
    {
        try
        {
            using var pig = PigDoc.Open(slipBytes);
            var notes = ExtractFabNotes(ExtractRows(pig));
            return DedupeNotesBySlug(DedupeNotes(notes), log);
        }
        catch (Exception ex)
        {
            log?.LogWarning(ex, "[FAB] text scan failed");
            return new List<FabNote>();
        }
    }

    /// <summary>Quantity-aware twin of <see cref="DedupeDescs"/> — collapses whitespace-identical notes.</summary>
    internal static List<FabNote> DedupeNotes(IEnumerable<FabNote> notes)
    {
        var seen   = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        var result = new List<FabNote>();
        foreach (var n in notes)
            if (seen.Add(_wsRx.Replace(n.Desc, " ").Trim()))
                result.Add(n);
        return result;
    }

    /// <summary>Quantity-aware twin of <see cref="DedupeBySlug"/> — one note per part slug, keeping the
    /// longest (least-clipped) description and its quantity.</summary>
    internal static List<FabNote> DedupeNotesBySlug(List<FabNote> notes, ILogger? log = null)
    {
        var best  = new Dictionary<string, FabNote>(StringComparer.OrdinalIgnoreCase);
        var order = new List<string>();
        var loose = new List<FabNote>();
        foreach (var n in notes)
        {
            string? slug = null;
            try { slug = FlatPattern.Develop(DrawingTextParser.Parse(n.Desc)).Cut.Part; } catch { }
            if (string.IsNullOrEmpty(slug)) { loose.Add(n); continue; }

            if (!best.TryGetValue(slug, out var cur)) { best[slug] = n; order.Add(slug); }
            else
            {
                if (n.Desc.Length > cur.Desc.Length) best[slug] = n;
                log?.LogInformation("[FAB] dedup: collapsed duplicate echo for slug '{Slug}'", slug);
            }
        }
        var result = order.Select(s => best[s]).ToList();
        result.AddRange(loose);
        return result;
    }

    /// <summary>Item letter for the FAB note at <paramref name="index"/> in deduped order: A…Z, then AA, AB…</summary>
    internal static string FabLetter(int index) =>
        index < 26
            ? ((char)('A' + index)).ToString()
            : $"{(char)('A' + index / 26 - 1)}{(char)('A' + index % 26)}";

    private static byte[]? RenderFab(string desc, ILogger? log, string? customerName = null, string? itemTag = null)
    {
        // Cache key includes customerName (column BOM header) and the item letter (drawn in the title block).
        var cacheKey = $"{desc}|{customerName}|{itemTag}";
        if (_cache.TryGetValue(cacheKey, out var cached)) return cached;
        try
        {
            var spec = DrawingTextParser.Parse(desc);
            var fp = FlatPattern.Develop(spec);
            var pdf = DrawingPdfRenderer.Render(fp, polishBilingual: false, customerName: customerName, itemTag: itemTag);
            _cache[cacheKey] = pdf;
            return pdf;
        }
        catch (Exception ex)
        {
            log?.LogWarning(ex, "[FAB] could not render note '{Desc}'", desc);
            return null;
        }
    }

    /// <summary>
    /// Pulls FAB note descriptions out of the page rows. A note authored with a terminator —
    /// <c>FAB: (1) [ … ]</c> — is captured between the brackets, stitching across wrapped rows until
    /// the closing <c>]</c>; that is layout-independent, so the inline line-item copy and the
    /// footer/special-instructions copy OpenBravo prints for the same note come out identical (and
    /// dedupe to one drawing), and a wrapped tail (e.g. an "edge 1.5/1.5" value) can't be dropped.
    /// A note without brackets falls back to the legacy column heuristic: the FAB cell often wraps in
    /// the narrow Product column (e.g. "… finish" on one row, "outside" on the next); a continuation
    /// stays in the FAB cell's column (left edge ≳ the FAB row's) and isn't a new FAB note, a blank,
    /// or a new line-item (which reaches the far-left MSPC column).
    /// </summary>
    internal static List<string> ExtractFabDescs(List<(string Text, double Left)> rows)
        => ExtractFabNotes(rows).Select(n => n.Desc).ToList();

    /// <summary>As <see cref="ExtractFabDescs"/>, but also captures the order quantity from the note's
    /// <c>(N)</c> prefix (defaults to 1) so the DXF can label each part "xN".</summary>
    internal static List<FabNote> ExtractFabNotes(List<(string Text, double Left)> rows)
    {
        var texts = rows.Select(r => r.Text).ToList();
        var lefts = rows.Select(r => r.Left).ToList();
        var notes = new List<FabNote>();
        for (int i = 0; i < rows.Count; i++)
        {
            var m = FabRx.Match(texts[i]);
            if (!m.Success) continue;
            int qty  = int.TryParse(m.Groups[1].Value, out var q) && q > 0 ? q : 1;
            var desc = StitchFabDesc(texts, lefts, i, m.Groups[2].Value.Trim(), out int consumed);
            i = consumed;                                 // skip the rows folded into this note
            if (desc.Length >= 3) notes.Add(new FabNote(qty, desc));
        }
        return notes;
    }

    /// <summary>
    /// Stitches one FAB note that begins at <paramref name="startIndex"/> into its full description,
    /// mirroring the rules in <see cref="ExtractFabNotes"/>: a bracket-bounded note (<c>FAB: [ … ]</c>)
    /// captures everything to the matching <c>]</c> (across wrapped rows, cap 10); a bracket-less note
    /// folds in continuation rows that stay in the FAB cell's column. <paramref name="lastConsumed"/>
    /// is the last row index folded in (== startIndex when nothing wrapped).
    /// </summary>
    private static string StitchFabDesc(IReadOnlyList<string> texts, IReadOnlyList<double> lefts,
        int startIndex, string firstDesc, out int lastConsumed)
    {
        lastConsumed = startIndex;
        var desc = firstDesc;
        if (desc.StartsWith("["))
        {
            desc = desc.Substring(1);
            int close = desc.IndexOf(']');
            for (int j = startIndex + 1; close < 0 && j < texts.Count && j <= startIndex + 10; j++)
            {
                if (string.IsNullOrWhiteSpace(texts[j])) break;
                if (FabRx.IsMatch(texts[j])) break;
                desc += " " + texts[j].Trim();
                lastConsumed = j;
                close = desc.IndexOf(']');
            }
            if (close >= 0) desc = desc.Substring(0, close);
            return desc.Trim();
        }

        double fabLeft = lefts[startIndex];
        for (int j = startIndex + 1; j < texts.Count; j++)
        {
            if (string.IsNullOrWhiteSpace(texts[j])) break;
            if (FabRx.IsMatch(texts[j])) break;
            if (lefts[j] < fabLeft - 12.0) break;   // reaches a left column (MSPC) → new line-item
            desc += " " + texts[j].Trim();
            lastConsumed = j;
        }
        return desc;
    }

    /// <summary>Develops a FAB note description to its part slug (e.g. <c>flitch_109x7.25</c>), or null
    /// if it doesn't parse/develop. The slug keys the item letter to both the drawing page and the note.</summary>
    private static string? DevelopSlug(string desc)
    {
        try { return FlatPattern.Develop(DrawingTextParser.Parse(desc)).Cut.Part; } catch { return null; }
    }

    /// <summary>
    /// Collapses notes that are identical after whitespace normalisation to a single entry, keeping
    /// the first occurrence. OpenBravo prints each FAB note twice (inline line-item + footer); with
    /// bracket-bounded capture both copies normalise identically, so this removes the duplicate
    /// drawing. Case- and whitespace-insensitive (different wrap points introduce differing spaces).
    /// </summary>
    internal static List<string> DedupeDescs(IEnumerable<string> descs)
    {
        var seen   = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        var result = new List<string>();
        foreach (var d in descs)
            if (seen.Add(_wsRx.Replace(d, " ").Trim()))
                result.Add(d);
        return result;
    }

    /// <summary>
    /// Collapses notes that develop to the SAME part (slug, e.g. <c>flitch_109x7.25</c>) to one drawing,
    /// keeping the LONGEST description. OpenBravo echoes each FAB note twice — inline in the line-items
    /// and again in the special-instructions footer — and the two can wrap/clip differently (the footer
    /// box clips "edge" to "edg", so that copy fails to parse the edge value and renders the default
    /// margin). Text matching can't see through that; part identity can. Keeping the longest capture
    /// keeps the least-clipped copy, so the surviving drawing has the fully-specified geometry.
    /// Descriptions that don't develop are passed through unchanged (RenderFab logs the parse failure).
    /// </summary>
    internal static List<string> DedupeBySlug(List<string> descs, ILogger? log = null)
    {
        var best  = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        var order = new List<string>();   // slugs, first-seen order
        var loose = new List<string>();   // un-developable — keep as-is
        foreach (var d in descs)
        {
            string? slug = null;
            try { slug = FlatPattern.Develop(DrawingTextParser.Parse(d)).Cut.Part; } catch { }
            if (string.IsNullOrEmpty(slug)) { loose.Add(d); continue; }

            if (!best.TryGetValue(slug, out var cur)) { best[slug] = d; order.Add(slug); }
            else
            {
                if (d.Length > cur.Length) best[slug] = d;
                log?.LogInformation("[FAB] dedup: collapsed duplicate echo for slug '{Slug}'", slug);
            }
        }
        var result = order.Select(s => best[s]).ToList();
        result.AddRange(loose);
        return result;
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

    /// <summary>Where a FAB: note sits in the slip body: its page, the FAB: row's left edge and vertical
    /// span (PDF points, bottom-left origin), and the note's developed part slug. Used to stamp the item
    /// letter beside the note. Slug is empty when the note didn't develop (e.g. a clipped footer echo).</summary>
    private sealed record FabAnchor(int PageIndex, double Left, double Top, double Bottom, string Slug);

    /// <summary>Finds every FAB: note occurrence in the slip with its page + position, so the item letter
    /// can be stamped beside it. Mirrors the row grouping in <see cref="ExtractRows"/> but keeps the FAB
    /// row's bounding box and page; the note is stitched (via <see cref="StitchFabDesc"/>) and developed
    /// to a slug so the same letter lands on the note and its drawing page.</summary>
    private static List<FabAnchor> ExtractFabAnchors(PigDoc doc)
    {
        var anchors = new List<FabAnchor>();
        var pages = doc.GetPages().ToList();
        for (int p = 0; p < pages.Count; p++)
        {
            var rows = pages[p].GetWords()
                .GroupBy(w => (int)Math.Round(w.BoundingBox.Bottom / 2.0))   // ~2pt baseline tolerance
                .OrderByDescending(g => g.Key)
                .Select(g =>
                {
                    var ws = g.OrderBy(w => w.BoundingBox.Left).ToList();
                    return (
                        Text:   string.Join(" ", ws.Select(w => w.Text)),
                        Left:   ws.Min(w => w.BoundingBox.Left),
                        Top:    ws.Max(w => w.BoundingBox.Top),
                        Bottom: ws.Min(w => w.BoundingBox.Bottom));
                })
                .ToList();

            var texts = rows.Select(r => r.Text).ToList();
            var lefts = rows.Select(r => r.Left).ToList();
            for (int i = 0; i < rows.Count; i++)
            {
                var m = FabRx.Match(texts[i]);
                if (!m.Success) continue;
                var desc = StitchFabDesc(texts, lefts, i, m.Groups[2].Value.Trim(), out _);
                if (desc.Length < 3) continue;
                anchors.Add(new FabAnchor(p, rows[i].Left, rows[i].Top, rows[i].Bottom, DevelopSlug(desc) ?? ""));
            }
        }
        return anchors;
    }

    /// <summary>Draws the item letter (black circled glyph) just left of each FAB: note in the slip body,
    /// keyed by part slug so it matches the letter on the note's drawing page. Best-effort per anchor —
    /// a draw failure (or an undeveloped/clipped echo with no slug) is skipped, never fatal.</summary>
    private static void StampFabLetters(PdfDocument outDoc, List<FabAnchor> anchors,
        Dictionary<string, string> slugToLetter, ILogger? log)
    {
        // One letter per note: OpenBravo prints each FAB note more than once (inline line-item +
        // special-instructions footer), so a slug can have several anchors. Stamp only the first
        // occurrence (reading order) — otherwise the same letter lands two-plus times on the slip.
        var stamped = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        foreach (var a in anchors)
        {
            if (string.IsNullOrEmpty(a.Slug) || !slugToLetter.TryGetValue(a.Slug, out var letter)) continue;
            if (!stamped.Add(a.Slug)) continue;          // already stamped this note's letter
            if (a.PageIndex < 0 || a.PageIndex >= outDoc.PageCount) continue;
            try
            {
                var page = outDoc.Pages[a.PageIndex];
                using var gfx = XGraphics.FromPdfPage(page, XGraphicsPdfPageOptions.Append);
                double ph = page.Height.Point;
                const double r = 7, gap = 4;
                double cy = ph - (a.Top + a.Bottom) / 2.0;     // FAB: row centre, flipped to top-left origin
                double cx = a.Left - gap - r;                  // just left of the FAB: text
                if (cx < r + 2) cx = a.Left + gap + r;         // no room on the left → sit just right of it
                var box = new XRect(cx - r, cy - r, 2 * r, 2 * r);
                gfx.DrawEllipse(new XPen(XColors.Black, 1.2), box);
                gfx.DrawString(letter, new XFont("Arial", 9, XFontStyleEx.Bold), XBrushes.Black, box, XStringFormats.Center);
            }
            catch (Exception ex) { log?.LogWarning(ex, "[FAB] could not stamp letter on slip page {Page}", a.PageIndex); }
        }
    }
}
