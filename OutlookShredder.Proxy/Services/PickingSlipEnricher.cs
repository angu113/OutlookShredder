using System.Text.RegularExpressions;
using PdfSharp.Drawing;
using PdfSharp.Pdf.IO;
using UglyToad.PdfPig.Content;
using PigDoc  = UglyToad.PdfPig.PdfDocument;
using SharpDoc = PdfSharp.Pdf.PdfDocument;

namespace OutlookShredder.Proxy.Services;

/// <summary>
/// Enriches picking-slip PDFs:
///   1. Bolds shop-comment lines in-place (white rect + bold redraw).
///   2. Appends a callout page with geometric drawings for bend/cut keywords.
/// </summary>
internal static class PickingSlipEnricher
{
    // ── Keyword detection ────────────────────────────────────────────────────

    private static readonly Regex _cutToRegex =
        new(@"\bcut\s+to\b", RegexOptions.IgnoreCase | RegexOptions.Compiled);

    private static readonly Regex _bendRegex =
        new(@"\bbend\s+([ULJZ])\b", RegexOptions.IgnoreCase | RegexOptions.Compiled);

    // ── Public entry point ───────────────────────────────────────────────────

    public static byte[] Enrich(byte[] pdfBytes)
    {
        var blocks = ParseCommentBlocks(pdfBytes);
        if (blocks.Count == 0) return pdfBytes;

        // Collect all keywords across all blocks.
        var keywords = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        foreach (var b in blocks)
            foreach (var kw in b.Keywords)
                keywords.Add(kw);

        var result = BoldComments(pdfBytes, blocks);

        if (keywords.Count > 0)
            result = AppendCalloutPage(result, keywords);

        return result;
    }

    // ── Structure parsing (PdfPig) ───────────────────────────────────────────

    // MSPC lines: letters-slash-digits, e.g. HP/375, HTSQ/22120, AF6061/2502500
    // The MSPC and product name share the same visual row, so text may continue
    // after the code — we only check the start of the grouped line.
    private static readonly Regex _mspcRegex =
        new(@"^[A-Z]{1,8}/\d{2,8}", RegexOptions.Compiled);

    // Balance / order line — starts with "B: " followed by a digit or letter
    private static readonly Regex _bLineRegex =
        new(@"^B:\s", RegexOptions.Compiled);

    // Cut instruction lines: S1:, B1:, OC1:, TB:, SC:, C1: etc.
    // B: (no digit) is the order/balance line and is handled separately above.
    private static readonly Regex _cutInstructionRegex =
        new(@"^(?:S\d+:|B\d+:|OC\d*:|TB\d*:|SC\d*:|C\d+:)", RegexOptions.IgnoreCase | RegexOptions.Compiled);

    private sealed record CommentBlock(
        int PageIndex,
        IReadOnlyList<TextLine> Lines,
        IReadOnlyList<string> Keywords);

    private sealed record TextLine(string Text, double X, double Y, double Width, double Height);

    private static List<CommentBlock> ParseCommentBlocks(byte[] pdfBytes)
    {
        var result = new List<CommentBlock>();

        using var doc = PigDoc.Open(pdfBytes);

        foreach (var page in doc.GetPages())
        {
            int pageIdx = page.Number - 1;
            var rawLines = GroupIntoLines(page.GetWords().ToList());

            // State machine per page:
            //   Preamble   — scanning for an MSPC line
            //   InProduct  — saw MSPC; waiting for the B: balance/order line
            //   AfterBLine — saw B: line; any non-instruction line is a shop comment
            var state = ParseState.Preamble;
            var commentLines = new List<TextLine>();

            foreach (var line in rawLines)
            {
                var text = line.Text.Trim();
                if (string.IsNullOrWhiteSpace(text)) continue;

                switch (state)
                {
                    case ParseState.Preamble:
                        if (_mspcRegex.IsMatch(text))
                            state = ParseState.InProduct;
                        break;

                    case ParseState.InProduct:
                        if (_mspcRegex.IsMatch(text))
                        {
                            // Back-to-back MSPC (e.g. service line after product) — keep state
                        }
                        else if (_bLineRegex.IsMatch(text))
                        {
                            state = ParseState.AfterBLine;
                        }
                        break;

                    case ParseState.AfterBLine:
                        if (_mspcRegex.IsMatch(text))
                        {
                            MaybeAddBlock(result, pageIdx, commentLines);
                            commentLines = [];
                            state = ParseState.InProduct;
                        }
                        else if (_bLineRegex.IsMatch(text))
                        {
                            // Additional B: line — not a comment
                        }
                        else if (_cutInstructionRegex.IsMatch(text))
                        {
                            // First cut instruction ends the comment zone for this product.
                            // Flush and move to DoneWithProduct so service lines that follow
                            // (e.g. "LASER CUTTING Laser Cutting") are not mistaken for comments.
                            MaybeAddBlock(result, pageIdx, commentLines);
                            commentLines = [];
                            state = ParseState.DoneWithProduct;
                        }
                        else
                        {
                            commentLines.Add(line);
                        }
                        break;

                    case ParseState.DoneWithProduct:
                        // Ignore everything — service lines, extra cut rows, dividers —
                        // until the next MSPC with a slash format starts a new product block.
                        if (_mspcRegex.IsMatch(text))
                            state = ParseState.InProduct;
                        break;
                }
            }

            MaybeAddBlock(result, pageIdx, commentLines);
        }

        return result;
    }

    private static void MaybeAddBlock(List<CommentBlock> result, int pageIdx, List<TextLine> lines)
    {
        if (lines.Count == 0) return;

        var keywords = new List<string>();
        foreach (var l in lines)
        {
            if (_cutToRegex.IsMatch(l.Text))
                keywords.Add("cut to");
            foreach (Match m in _bendRegex.Matches(l.Text))
                keywords.Add($"bend {m.Groups[1].Value.ToUpperInvariant()}");
        }

        result.Add(new CommentBlock(pageIdx, lines, keywords));
    }

    private enum ParseState { Preamble, InProduct, AfterBLine, DoneWithProduct }

    // ── Word → line grouping ─────────────────────────────────────────────────

    private static List<TextLine> GroupIntoLines(List<Word> words)
    {
        if (words.Count == 0) return [];

        // Sort top-to-bottom (PdfPig Y is bottom-up, so descending Y = top-first)
        var sorted = words.OrderByDescending(w => w.BoundingBox.Bottom).ToList();

        var lineGroups = new List<List<Word>>();
        List<Word>? current = null;
        double currentY = double.NaN;

        foreach (var word in sorted)
        {
            double y = word.BoundingBox.Bottom;
            if (current is null || Math.Abs(y - currentY) > 4)
            {
                current = [word];
                lineGroups.Add(current);
                currentY = y;
            }
            else
            {
                current.Add(word);
            }
        }

        var result = new List<TextLine>();
        foreach (var group in lineGroups)
        {
            // Sort left-to-right within the line
            var ordered = group.OrderBy(w => w.BoundingBox.Left).ToList();
            var text = string.Join(" ", ordered.Select(w => w.Text));

            double left   = ordered.Min(w => w.BoundingBox.Left);
            double bottom = ordered.Min(w => w.BoundingBox.Bottom);
            double right  = ordered.Max(w => w.BoundingBox.Right);
            double top    = ordered.Max(w => w.BoundingBox.Top);

            result.Add(new TextLine(text, left, bottom, right - left, top - bottom));
        }

        return result;
    }

    // ── Bold overlay (PdfSharp) ──────────────────────────────────────────────

    private static byte[] BoldComments(byte[] pdfBytes, List<CommentBlock> blocks)
    {
        using var ms = new MemoryStream(pdfBytes);
        var doc = PdfReader.Open(ms, PdfDocumentOpenMode.Modify);

        var boxPen = new XPen(XColors.Black, 0.75) { DashStyle = XDashStyle.Dash };

        foreach (var block in blocks)
        {
            var page = doc.Pages[block.PageIndex];
            double pageH = page.Height.Point;

            using var gfx = XGraphics.FromPdfPage(page);
            var boldFont = new XFont("Arial", 11, XFontStyleEx.Bold);

            // Compute the union bounding box for the whole block so we can draw one
            // enclosing rectangle around all comment lines in this product block.
            double bLeft   = block.Lines.Min(l => l.X) - 4;
            double bBottom = block.Lines.Min(l => l.Y);
            double bRight  = block.Lines.Max(l => l.X + l.Width) + 4;
            double bTop    = block.Lines.Max(l => l.Y + l.Height);

            foreach (var line in block.Lines)
            {
                // Convert PdfPig bottom-left coords to PdfSharp top-left
                double psY = pageH - line.Y - line.Height;

                // White out original text
                var rect = new XRect(line.X - 2, psY - 1, line.Width + 4, line.Height + 3);
                gfx.DrawRectangle(XBrushes.White, rect);

                // Redraw bold
                var textRect = new XRect(line.X, psY, line.Width + 20, line.Height + 4);
                gfx.DrawString(line.Text, boldFont, XBrushes.Black,
                    textRect, XStringFormats.TopLeft);
            }

            // Draw enclosing box around the entire comment block (PdfSharp top-left origin)
            double boxX = bLeft;
            double boxY = pageH - bTop - 3;
            double boxW = bRight - bLeft;
            double boxH = bTop - bBottom + 6;
            gfx.DrawRectangle(boxPen, new XRect(boxX, boxY, boxW, boxH));
        }

        using var outMs = new MemoryStream();
        doc.Save(outMs);
        return outMs.ToArray();
    }

    // ── Callout page (PdfSharp) ──────────────────────────────────────────────

    private static byte[] AppendCalloutPage(byte[] pdfBytes, ISet<string> keywords)
    {
        using var ms = new MemoryStream(pdfBytes);
        var doc = PdfReader.Open(ms, PdfDocumentOpenMode.Modify);

        var page = doc.AddPage();
        page.Width  = XUnit.FromInch(8.5);
        page.Height = XUnit.FromInch(11);

        using var gfx = XGraphics.FromPdfPage(page);

        var titleFont  = new XFont("Arial", 14, XFontStyleEx.Bold);
        var labelFont  = new XFont("Arial", 10, XFontStyleEx.Bold);
        var dimFont    = new XFont("Arial",  8, XFontStyleEx.Regular);

        gfx.DrawString("Shop Callouts", titleFont, XBrushes.Black,
            new XRect(0, 20, page.Width.Point, 24), XStringFormats.TopCenter);

        double x = 60;
        double y = 70;
        double cellW = 180;
        double cellH = 200;
        int col = 0;

        // Ordered for consistent layout
        var ordered = new[] { "cut to", "bend U", "bend L", "bend J", "bend Z" }
            .Where(k => keywords.Contains(k))
            .ToList();

        foreach (var kw in ordered)
        {
            DrawCallout(gfx, labelFont, dimFont, kw, x + col * (cellW + 30), y);
            col++;
            if (col == 3) { col = 0; y += cellH + 20; }
        }

        using var outMs = new MemoryStream();
        doc.Save(outMs);
        return outMs.ToArray();
    }

    private static void DrawCallout(
        XGraphics gfx, XFont labelFont, XFont dimFont,
        string keyword, double ox, double oy)
    {
        // Label
        gfx.DrawString(keyword.ToUpperInvariant(), labelFont, XBrushes.Black,
            new XRect(ox, oy, 160, 16), XStringFormats.TopLeft);

        double sy = oy + 22; // shape origin Y

        var pen = new XPen(XColors.Black, 2);

        switch (keyword.ToLowerInvariant())
        {
            case "cut to":
                DrawCutTo(gfx, pen, dimFont, ox, sy);
                break;
            case "bend u":
                DrawBendU(gfx, pen, ox, sy);
                break;
            case "bend l":
                DrawBendL(gfx, pen, ox, sy);
                break;
            case "bend j":
                DrawBendJ(gfx, pen, ox, sy);
                break;
            case "bend z":
                DrawBendZ(gfx, pen, ox, sy);
                break;
        }
    }

    // ── Individual shape drawings ────────────────────────────────────────────

    private static void DrawCutTo(XGraphics gfx, XPen pen, XFont dimFont, double ox, double oy)
    {
        // Horizontal bar
        gfx.DrawLine(pen, ox, oy + 20, ox + 120, oy + 20);
        gfx.DrawLine(pen, ox, oy + 30, ox + 120, oy + 30);

        // Cut marks (angled lines through the bar)
        var cutPen = new XPen(XColors.Red, 1.5);
        double[] cx = [ox + 40, ox + 80];
        foreach (var x in cx)
        {
            gfx.DrawLine(cutPen, x - 8, oy + 10, x + 8, oy + 40);
        }

        // Dimension arrows
        DrawDimArrow(gfx, pen, dimFont, ox, oy + 50, ox + 40, oy + 50, "L1");
        DrawDimArrow(gfx, pen, dimFont, ox + 40, oy + 50, ox + 80, oy + 50, "L2");
        DrawDimArrow(gfx, pen, dimFont, ox + 80, oy + 50, ox + 120, oy + 50, "L3");
    }

    private static void DrawBendU(XGraphics gfx, XPen pen, double ox, double oy)
    {
        // U-channel cross section
        // Left leg
        gfx.DrawLine(pen, ox + 20, oy, ox + 20, oy + 60);
        // Bottom web
        gfx.DrawLine(pen, ox + 20, oy + 60, ox + 100, oy + 60);
        // Right leg
        gfx.DrawLine(pen, ox + 100, oy + 60, ox + 100, oy);

        // Inner offset (thickness indication)
        var thinPen = new XPen(XColors.Gray, 0.75);
        gfx.DrawLine(thinPen, ox + 28, oy + 2, ox + 28, oy + 52);
        gfx.DrawLine(thinPen, ox + 28, oy + 52, ox + 92, oy + 52);
        gfx.DrawLine(thinPen, ox + 92, oy + 52, ox + 92, oy + 2);

        // Dimension labels
        var labelPen = new XPen(XColors.DimGray, 0.5);
        gfx.DrawString("flange", new XFont("Arial", 7, XFontStyleEx.Regular), XBrushes.DimGray,
            new XRect(ox, oy - 12, 40, 12), XStringFormats.TopLeft);
        gfx.DrawString("flange", new XFont("Arial", 7, XFontStyleEx.Regular), XBrushes.DimGray,
            new XRect(ox + 88, oy - 12, 40, 12), XStringFormats.TopLeft);
        gfx.DrawString("web", new XFont("Arial", 7, XFontStyleEx.Regular), XBrushes.DimGray,
            new XRect(ox + 48, oy + 62, 30, 12), XStringFormats.TopLeft);
    }

    private static void DrawBendL(XGraphics gfx, XPen pen, double ox, double oy)
    {
        // L-angle cross section
        // Vertical leg
        gfx.DrawLine(pen, ox + 20, oy, ox + 20, oy + 70);
        // Horizontal leg
        gfx.DrawLine(pen, ox + 20, oy + 70, ox + 90, oy + 70);

        // Thickness lines
        var thinPen = new XPen(XColors.Gray, 0.75);
        gfx.DrawLine(thinPen, ox + 28, oy + 2, ox + 28, oy + 62);
        gfx.DrawLine(thinPen, ox + 28, oy + 62, ox + 90, oy + 62);

        // 90° arc at corner
        gfx.DrawArc(new XPen(XColors.DimGray, 0.5),
            ox + 20, oy + 50, 14, 14, 180, -90);

        gfx.DrawString("leg", new XFont("Arial", 7, XFontStyleEx.Regular), XBrushes.DimGray,
            new XRect(ox + 2, oy + 30, 20, 12), XStringFormats.TopLeft);
        gfx.DrawString("leg", new XFont("Arial", 7, XFontStyleEx.Regular), XBrushes.DimGray,
            new XRect(ox + 52, oy + 72, 20, 12), XStringFormats.TopLeft);
    }

    private static void DrawBendJ(XGraphics gfx, XPen pen, double ox, double oy)
    {
        // J-hook: tall vertical, short bottom hook to one side
        // Outer profile
        gfx.DrawLine(pen, ox + 40, oy, ox + 40, oy + 70);      // outer vertical
        gfx.DrawArc(pen, ox + 20, oy + 50, 40, 40, 0, 180);     // bottom curve outer
        gfx.DrawLine(pen, ox + 20, oy + 70, ox + 20, oy + 40);  // short inner leg

        // Inner offset
        var thinPen = new XPen(XColors.Gray, 0.75);
        gfx.DrawLine(thinPen, ox + 32, oy + 2, ox + 32, oy + 56);
        gfx.DrawArc(thinPen, ox + 26, oy + 52, 24, 24, 0, 180);
        gfx.DrawLine(thinPen, ox + 26, oy + 64, ox + 26, oy + 40);

        gfx.DrawString("hook", new XFont("Arial", 7, XFontStyleEx.Regular), XBrushes.DimGray,
            new XRect(ox + 4, oy + 80, 30, 12), XStringFormats.TopLeft);
    }

    private static void DrawBendZ(XGraphics gfx, XPen pen, double ox, double oy)
    {
        // Z (or S offset): top flange → offset web → bottom flange
        gfx.DrawLine(pen, ox + 60, oy + 10, ox + 110, oy + 10);  // top flange
        gfx.DrawLine(pen, ox + 60, oy + 18, ox + 110, oy + 18);

        gfx.DrawLine(pen, ox + 68, oy + 18, ox + 42, oy + 52);   // offset web outer
        gfx.DrawLine(pen, ox + 60, oy + 18, ox + 34, oy + 52);   // offset web inner

        gfx.DrawLine(pen, ox + 10, oy + 52, ox + 60, oy + 52);   // bottom flange
        gfx.DrawLine(pen, ox + 10, oy + 60, ox + 60, oy + 60);

        gfx.DrawString("offset", new XFont("Arial", 7, XFontStyleEx.Regular), XBrushes.DimGray,
            new XRect(ox + 30, oy + 28, 40, 12), XStringFormats.TopLeft);
    }

    // ── Dimension arrow helper ───────────────────────────────────────────────

    private static void DrawDimArrow(
        XGraphics gfx, XPen pen, XFont font,
        double x1, double y, double x2, double y2, string label)
    {
        var dimPen = new XPen(XColors.DimGray, 0.75);
        gfx.DrawLine(dimPen, x1, y, x2, y2);

        // Arrowheads
        double ah = 5;
        gfx.DrawLine(dimPen, x1, y, x1 + ah, y - ah / 2);
        gfx.DrawLine(dimPen, x1, y, x1 + ah, y + ah / 2);
        gfx.DrawLine(dimPen, x2, y2, x2 - ah, y2 - ah / 2);
        gfx.DrawLine(dimPen, x2, y2, x2 - ah, y2 + ah / 2);

        double midX = (x1 + x2) / 2;
        gfx.DrawString(label, font, XBrushes.DimGray,
            new XRect(midX - 15, y - 14, 30, 12), XStringFormats.TopCenter);
    }
}
