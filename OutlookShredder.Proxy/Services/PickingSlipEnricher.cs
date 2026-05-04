using System.Text.RegularExpressions;
using Microsoft.Extensions.Logging;
using PdfSharp.Drawing;
using PdfSharp.Pdf;
using PdfSharp.Pdf.IO;
using UglyToad.PdfPig.Content;
using PigDoc   = UglyToad.PdfPig.PdfDocument;
using SharpDoc = PdfSharp.Pdf.PdfDocument;

namespace OutlookShredder.Proxy.Services;

/// <summary>
/// Enriches picking-slip PDFs:
///   1. Stamps the customer name in a box at the top-left of page 1.
///   2. Bolds shop-comment lines in-place (white rect + bold redraw).
///   3. Stamps the rep's first name over the "Customer Rep:" / "Store #:" area.
///   4. Appends a callout page with geometric drawings for bend/cut keywords.
///
/// EnrichPickingSlip does all four in a single PdfPig read + single PdfSharp write.
/// </summary>
internal static class PickingSlipEnricher
{
    // ── Font resolver ─────────────────────────────────────────────────────────

    private static readonly object _fontLock = new();
    private static bool _fontResolverSet;

    private static void EnsureFontResolver()
    {
        if (_fontResolverSet) return;
        lock (_fontLock)
        {
            if (_fontResolverSet) return;
            PdfSharp.Fonts.GlobalFontSettings.FontResolver = new ArialFontResolver();
            _fontResolverSet = true;
        }
    }

    private sealed class ArialFontResolver : PdfSharp.Fonts.IFontResolver
    {
        private static readonly string FontsDir =
            Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Windows), "Fonts");

        public PdfSharp.Fonts.FontResolverInfo? ResolveTypeface(string familyName, bool isBold, bool isItalic)
        {
            if (!familyName.Equals("Arial", StringComparison.OrdinalIgnoreCase)) return null;
            var face = (isBold, isItalic) switch
            {
                (true,  true)  => "arialbi",
                (true,  false) => "arialbd",
                (false, true)  => "ariali",
                _              => "arial",
            };
            return new PdfSharp.Fonts.FontResolverInfo(face);
        }

        public byte[]? GetFont(string faceName) =>
            new[] { $"{faceName}.ttf", $"{faceName}.ttc", $"{faceName}.otf" }
                .Select(f => Path.Combine(FontsDir, f))
                .Where(File.Exists)
                .Select(File.ReadAllBytes)
                .FirstOrDefault();
    }

    // ── Keyword detection ─────────────────────────────────────────────────────

    private static readonly Regex _cutToRegex =
        new(@"\bcut\s+to\b", RegexOptions.IgnoreCase | RegexOptions.Compiled);

    private static readonly Regex _bendRegex =
        new(@"\bbend\s+([ULJZ])\b", RegexOptions.IgnoreCase | RegexOptions.Compiled);

    private static readonly Regex _mspcRegex =
        new(@"^[A-Z][A-Z0-9]{0,9}/\d{2,8}", RegexOptions.Compiled);

    private static readonly Regex _bLineRegex =
        new(@"^B:\s", RegexOptions.Compiled);

    private static readonly Regex _cutInstructionRegex =
        new(@"^(?:S\d+:|B\d+:|OC\d*:|TB\d*:|SC\d*:|C\d+:)", RegexOptions.IgnoreCase | RegexOptions.Compiled);

    private static readonly Regex _footerMarkerRegex =
        new(@"^(TOTAL[\s:]+\d|Description\s*\(Special|Contact\s*:\s*\w|Delivery\s+Services)",
            RegexOptions.IgnoreCase | RegexOptions.Compiled);

    private static readonly Regex _serviceLineRegex =
        new(@"^[A-Z]{3,}(?:\s+[A-Z]{2,})*(?:\s+[A-Z][a-z]|\s*$)", RegexOptions.Compiled);

    // ── Data records ─────────────────────────────────────────────────────────

    private sealed record CommentBlock(
        int PageIndex,
        IReadOnlyList<TextLine> Lines,
        IReadOnlyList<string> Keywords);

    private sealed record TextLine(string Text, double X, double Y, double Width, double Height);

    private enum ParseState { Preamble, InProduct, AfterBLine, DoneWithProduct }

    // ── Combined entry point (1 PdfPig read + 1 PdfSharp write) ──────────────

    /// <summary>
    /// Applies all picking-slip enrichments in two PDF passes (one read, one write).
    /// Returns the enriched bytes and the ship-to customer name extracted from the PDF
    /// (null when the Ship To label is not found).
    /// </summary>
    public static (byte[] Bytes, string? ShipToName) EnrichPickingSlip(
        byte[] pdfBytes, string? knownCustomerName = null, ILogger? log = null)
    {
        EnsureFontResolver();

        // ── PdfPig pass: extract everything ──────────────────────────────────
        string? shipToName   = null;
        string? repFirstName = null;
        double  coverL = double.MaxValue, coverB = double.MaxValue;
        double  coverR = double.MinValue, coverT = double.MinValue;
        double  repLineH = 10;
        List<CommentBlock> blocks;

        using (var pigDoc = PigDoc.Open(pdfBytes))
        {
            var page1Words = pigDoc.GetPage(1).GetWords().ToList();

            // 1. Ship-to customer name (left column, line below "Ship" word)
            shipToName = ExtractShipToNameFromWords(page1Words, log);

            // 2. Rep name + cover rect (for white-out + re-stamp)
            var page1Lines = GroupIntoLines(page1Words);
            foreach (var line in page1Lines)
            {
                bool isRepLine   = line.Text.IndexOf("Customer Rep:", StringComparison.OrdinalIgnoreCase) >= 0;
                bool isStoreLine = line.Text.IndexOf("Store #",       StringComparison.OrdinalIgnoreCase) >= 0;
                if (isRepLine)
                {
                    int sep  = line.Text.IndexOf("Customer Rep:", StringComparison.OrdinalIgnoreCase) + "Customer Rep:".Length;
                    var after = line.Text[sep..].Trim();
                    if (!string.IsNullOrWhiteSpace(after))
                        repFirstName = after.Split(' ', StringSplitOptions.RemoveEmptyEntries)[0];
                    repLineH = line.Height;
                }
                if (isRepLine || isStoreLine)
                {
                    coverL = Math.Min(coverL, line.X);
                    coverB = Math.Min(coverB, line.Y);
                    coverR = Math.Max(coverR, line.X + line.Width);
                    coverT = Math.Max(coverT, line.Y + line.Height);
                }
            }
            log?.LogInformation("[PSE] rep='{Rep}' cover=({L:F1},{B:F1})–({R:F1},{T:F1})",
                repFirstName ?? "(null)", coverL, coverB, coverR, coverT);

            // 3. Comment blocks from all pages
            blocks = ParseCommentBlocksFromDoc(pigDoc, log);
        }

        // Effective customer name (PDF takes priority over AI extraction)
        var customerName = !string.IsNullOrWhiteSpace(shipToName) ? shipToName : knownCustomerName;
        bool hasCustomerName = !string.IsNullOrWhiteSpace(customerName);
        bool hasRepName      = repFirstName is not null && coverL != double.MaxValue;

        var keywords = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        foreach (var b in blocks)
            foreach (var kw in b.Keywords)
                keywords.Add(kw);

        // Skip write pass if nothing to do
        if (!hasCustomerName && !hasRepName && blocks.Count == 0)
            return (pdfBytes, shipToName);

        // ── PdfSharp pass: apply all modifications in one shot ────────────────
        using var ms  = new MemoryStream(pdfBytes);
        var doc       = PdfReader.Open(ms, PdfDocumentOpenMode.Modify);

        if (hasCustomerName)
            StampCustomerNameOnDoc(doc, customerName!, log);

        if (blocks.Count > 0)
            BoldCommentsOnDoc(doc, blocks);

        if (hasRepName)
            StampRepNameOnDoc(doc, repFirstName!, coverL, coverB, coverR, coverT, repLineH, log);

        if (keywords.Count > 0)
            AppendCalloutPageToDoc(doc, keywords);

        using var outMs = new MemoryStream();
        doc.Save(outMs);
        return (outMs.ToArray(), shipToName);
    }

    // ── PdfPig extraction helpers ─────────────────────────────────────────────

    private static string? ExtractShipToNameFromWords(List<Word> words, ILogger? log)
    {
        const double leftColMax = 200.0;

        var shipWord = words
            .Where(w => w.Text.Equals("Ship", StringComparison.OrdinalIgnoreCase)
                     && w.BoundingBox.Left < leftColMax)
            .OrderByDescending(w => w.BoundingBox.Bottom)
            .FirstOrDefault();

        if (shipWord is null) return null;

        double shipY = shipWord.BoundingBox.Bottom;
        log?.LogInformation("[PSE] 'Ship' label at X={X:F1} Y={Y:F1}", shipWord.BoundingBox.Left, shipY);

        var nameWords = words
            .Where(w =>
            {
                double drop = shipY - w.BoundingBox.Bottom;
                return drop is > 4 and < 25 && w.BoundingBox.Left < leftColMax;
            })
            .OrderBy(w => w.BoundingBox.Left)
            .ToList();

        if (nameWords.Count == 0) return null;

        var name = string.Join(" ", nameWords.Select(w => w.Text)).Trim();
        log?.LogInformation("[PSE] Ship-to name: '{Name}'", name);
        return string.IsNullOrWhiteSpace(name) ? null : name;
    }

    private static List<CommentBlock> ParseCommentBlocksFromDoc(PigDoc doc, ILogger? log)
    {
        var result = new List<CommentBlock>();
        foreach (var page in doc.GetPages())
        {
            int pageIdx = page.Number - 1;
            var rawLines = GroupIntoLines(page.GetWords().ToList());
            var state = ParseState.Preamble;
            var commentLines = new List<TextLine>();

            foreach (var line in rawLines)
            {
                var text = line.Text.Trim();
                if (string.IsNullOrWhiteSpace(text)) continue;

                log?.LogDebug("[PSE] p{Page} [{State}] {Text}", pageIdx, state, text);

                switch (state)
                {
                    case ParseState.Preamble:
                        if (_mspcRegex.IsMatch(text)) state = ParseState.InProduct;
                        break;

                    case ParseState.InProduct:
                        if (_mspcRegex.IsMatch(text)) { /* back-to-back MSPC — keep state */ }
                        else if (_bLineRegex.IsMatch(text)) state = ParseState.AfterBLine;
                        break;

                    case ParseState.AfterBLine:
                        if (_mspcRegex.IsMatch(text))
                        {
                            MaybeAddBlock(result, pageIdx, commentLines);
                            commentLines = [];
                            state = ParseState.InProduct;
                        }
                        else if (_bLineRegex.IsMatch(text)) { /* extra B: — not a comment */ }
                        else if (_cutInstructionRegex.IsMatch(text) || _serviceLineRegex.IsMatch(text)
                                 || _footerMarkerRegex.IsMatch(text))
                        {
                            MaybeAddBlock(result, pageIdx, commentLines);
                            commentLines = [];
                            state = ParseState.DoneWithProduct;
                        }
                        else
                        {
                            log?.LogInformation("[PSE] COMMENT p{Page}: {Text}", pageIdx, text);
                            commentLines.Add(line);
                        }
                        break;

                    case ParseState.DoneWithProduct:
                        if (_mspcRegex.IsMatch(text)) state = ParseState.InProduct;
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
            if (_cutToRegex.IsMatch(l.Text)) keywords.Add("cut to");
            foreach (Match m in _bendRegex.Matches(l.Text))
                keywords.Add($"bend {m.Groups[1].Value.ToUpperInvariant()}");
        }
        result.Add(new CommentBlock(pageIdx, lines, keywords));
    }

    // ── PdfSharp modification helpers (operate on an already-open SharpDoc) ──

    private static void StampCustomerNameOnDoc(SharpDoc doc, string customerName, ILogger? log)
    {
        var page = doc.Pages[0];
        const double boxX      = 18;
        const double boxY      = 18;
        const double boxWidth  = 180;
        const double boxHeight = 72;
        const double pad       = 6;
        double textAreaW = boxWidth  - pad * 2;
        double textAreaH = boxHeight - pad * 2;

        using var gfx = XGraphics.FromPdfPage(page);
        var (lines, font) = FitTextInBox(gfx, customerName, textAreaW, textAreaH);

        var boxRect = new XRect(boxX, boxY, boxWidth, boxHeight);
        gfx.DrawRectangle(XBrushes.White, boxRect);
        gfx.DrawRectangle(new XPen(XColors.Black, 0.75), boxRect);

        double lineH  = font.GetHeight();
        double blockH = lines.Count * lineH;
        double startY = boxY + pad + (textAreaH - blockH) / 2.0;
        foreach (var line in lines)
        {
            gfx.DrawString(line, font, XBrushes.Black,
                new XRect(boxX + pad, startY, textAreaW, lineH), XStringFormats.TopCenter);
            startY += lineH;
        }
    }

    private static void BoldCommentsOnDoc(SharpDoc doc, List<CommentBlock> blocks)
    {
        foreach (var block in blocks)
        {
            var page  = doc.Pages[block.PageIndex];
            double pageH = page.Height.Point;
            using var gfx = XGraphics.FromPdfPage(page);
            var boldFont = new XFont("Arial", 13, XFontStyleEx.Bold);

            foreach (var line in block.Lines)
            {
                double psY      = pageH - line.Y - line.Height;
                var    rect     = new XRect(line.X - 10, psY - 4, line.Width + 80, line.Height + 8);
                var    whiteBrush = new XSolidBrush(XColors.White);
                gfx.DrawRectangle(whiteBrush, rect);
                gfx.DrawRectangle(whiteBrush, rect);
                gfx.DrawString(line.Text, boldFont, XBrushes.Black,
                    new XRect(line.X, psY, line.Width + 20, line.Height + 4), XStringFormats.TopLeft);
            }
        }
    }

    private static void StampRepNameOnDoc(
        SharpDoc doc,
        string repFirstName,
        double coverL, double coverB, double coverR, double coverT,
        double repLineH,
        ILogger? log)
    {
        var    sharpPage = doc.Pages[0];
        double pageH     = sharpPage.Height.Point;
        const double pad = 4;
        double psTop     = pageH - coverT;
        double psHeight  = coverT - coverB;
        var    coverRect = new XRect(coverL - pad, psTop - pad,
                                     (coverR - coverL) + pad * 2, psHeight + pad * 2);

        using var gfx = XGraphics.FromPdfPage(sharpPage);
        gfx.DrawRectangle(XBrushes.White, coverRect);
        gfx.DrawRectangle(XBrushes.White, coverRect);

        double bodyPt    = Math.Clamp(Math.Round(repLineH / 1.2), 7, 14);
        var    font      = new XFont("Arial", bodyPt + 4, XFontStyleEx.Bold);
        const double shiftRight = 18;
        var textRect = new XRect(coverRect.X + shiftRight, coverRect.Y,
                                 coverRect.Width - shiftRight, coverRect.Height);
        gfx.DrawString(repFirstName, font, XBrushes.Black, textRect, XStringFormats.CenterLeft);
        log?.LogInformation("[PSE] Stamped rep '{Rep}'", repFirstName);
    }

    private static void AppendCalloutPageToDoc(SharpDoc doc, ISet<string> keywords)
    {
        var page = doc.AddPage();
        page.Width  = XUnit.FromInch(8.5);
        page.Height = XUnit.FromInch(11);
        using var gfx = XGraphics.FromPdfPage(page);

        var titleFont = new XFont("Arial", 14, XFontStyleEx.Bold);
        var labelFont = new XFont("Arial", 10, XFontStyleEx.Bold);
        var dimFont   = new XFont("Arial",  8, XFontStyleEx.Regular);

        gfx.DrawString("Shop Callouts", titleFont, XBrushes.Black,
            new XRect(0, 20, page.Width.Point, 24), XStringFormats.TopCenter);

        double x = 60, y = 70, cellW = 180, cellH = 200;
        int col = 0;
        var ordered = new[] { "cut to", "bend U", "bend L", "bend J", "bend Z" }
            .Where(k => keywords.Contains(k))
            .ToList();
        foreach (var kw in ordered)
        {
            DrawCallout(gfx, labelFont, dimFont, kw, x + col * (cellW + 30), y);
            col++;
            if (col == 3) { col = 0; y += cellH + 20; }
        }
    }

    // ── Text layout helpers ───────────────────────────────────────────────────

    private static (List<string> Lines, XFont Font) FitTextInBox(
        XGraphics gfx, string text, double maxW, double maxH)
    {
        for (double size = 16; size >= 7; size -= 1)
        {
            var font  = new XFont("Arial", size, XFontStyleEx.Bold);
            var lines = WrapWords(gfx, font, text, maxW);
            if (lines.Count * font.GetHeight() <= maxH)
                return (lines, font);
        }
        var fallback = new XFont("Arial", 7, XFontStyleEx.Bold);
        return (WrapWords(gfx, fallback, text, maxW), fallback);
    }

    private static List<string> WrapWords(XGraphics gfx, XFont font, string text, double maxW)
    {
        var lines   = new List<string>();
        var current = "";
        foreach (var word in text.Split(' ', StringSplitOptions.RemoveEmptyEntries))
        {
            var candidate = string.IsNullOrEmpty(current) ? word : current + " " + word;
            if (gfx.MeasureString(candidate, font).Width <= maxW)
                current = candidate;
            else
            {
                if (!string.IsNullOrEmpty(current)) lines.Add(current);
                current = word;
            }
        }
        if (!string.IsNullOrEmpty(current)) lines.Add(current);
        return lines.Count > 0 ? lines : [text];
    }

    // ── Word → line grouping ──────────────────────────────────────────────────

    private static List<TextLine> GroupIntoLines(List<Word> words)
    {
        if (words.Count == 0) return [];
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
            else current.Add(word);
        }

        var result = new List<TextLine>();
        foreach (var group in lineGroups)
        {
            var ordered = group.OrderBy(w => w.BoundingBox.Left).ToList();
            var text    = string.Join(" ", ordered.Select(w => w.Text));
            result.Add(new TextLine(
                text,
                ordered.Min(w => w.BoundingBox.Left),
                ordered.Min(w => w.BoundingBox.Bottom),
                ordered.Max(w => w.BoundingBox.Right) - ordered.Min(w => w.BoundingBox.Left),
                ordered.Max(w => w.BoundingBox.Top)   - ordered.Min(w => w.BoundingBox.Bottom)));
        }
        return result;
    }

    // ── Callout shapes ────────────────────────────────────────────────────────

    private static void DrawCallout(
        XGraphics gfx, XFont labelFont, XFont dimFont,
        string keyword, double ox, double oy)
    {
        gfx.DrawString(keyword.ToUpperInvariant(), labelFont, XBrushes.Black,
            new XRect(ox, oy, 160, 16), XStringFormats.TopLeft);
        double sy  = oy + 22;
        var    pen = new XPen(XColors.Black, 2);
        switch (keyword.ToLowerInvariant())
        {
            case "cut to": DrawCutTo(gfx, pen, dimFont, ox, sy); break;
            case "bend u": DrawBendU(gfx, pen, ox, sy);          break;
            case "bend l": DrawBendL(gfx, pen, ox, sy);          break;
            case "bend j": DrawBendJ(gfx, pen, ox, sy);          break;
            case "bend z": DrawBendZ(gfx, pen, ox, sy);          break;
        }
    }

    private static void DrawCutTo(XGraphics gfx, XPen pen, XFont dimFont, double ox, double oy)
    {
        gfx.DrawLine(pen, ox, oy + 20, ox + 120, oy + 20);
        gfx.DrawLine(pen, ox, oy + 30, ox + 120, oy + 30);
        var cutPen = new XPen(XColors.Red, 1.5);
        double[] cx = [ox + 40, ox + 80];
        foreach (var x in cx)
        {
            gfx.DrawLine(cutPen, x - 8, oy + 10, x + 8, oy + 40);
        }
        DrawDimArrow(gfx, pen, dimFont, ox,       oy + 50, ox + 40,  oy + 50, "L1");
        DrawDimArrow(gfx, pen, dimFont, ox + 40,  oy + 50, ox + 80,  oy + 50, "L2");
        DrawDimArrow(gfx, pen, dimFont, ox + 80,  oy + 50, ox + 120, oy + 50, "L3");
    }

    private static void DrawBendU(XGraphics gfx, XPen pen, double ox, double oy)
    {
        gfx.DrawLine(pen, ox + 20, oy, ox + 20, oy + 60);
        gfx.DrawLine(pen, ox + 20, oy + 60, ox + 100, oy + 60);
        gfx.DrawLine(pen, ox + 100, oy + 60, ox + 100, oy);
        var thinPen = new XPen(XColors.Gray, 0.75);
        gfx.DrawLine(thinPen, ox + 28, oy + 2, ox + 28, oy + 52);
        gfx.DrawLine(thinPen, ox + 28, oy + 52, ox + 92, oy + 52);
        gfx.DrawLine(thinPen, ox + 92, oy + 52, ox + 92, oy + 2);
        var labelFont = new XFont("Arial", 7, XFontStyleEx.Regular);
        gfx.DrawString("flange", labelFont, XBrushes.DimGray, new XRect(ox, oy - 12, 40, 12), XStringFormats.TopLeft);
        gfx.DrawString("flange", labelFont, XBrushes.DimGray, new XRect(ox + 88, oy - 12, 40, 12), XStringFormats.TopLeft);
        gfx.DrawString("web",    labelFont, XBrushes.DimGray, new XRect(ox + 48, oy + 62, 30, 12), XStringFormats.TopLeft);
    }

    private static void DrawBendL(XGraphics gfx, XPen pen, double ox, double oy)
    {
        gfx.DrawLine(pen, ox + 20, oy, ox + 20, oy + 70);
        gfx.DrawLine(pen, ox + 20, oy + 70, ox + 90, oy + 70);
        var thinPen = new XPen(XColors.Gray, 0.75);
        gfx.DrawLine(thinPen, ox + 28, oy + 2, ox + 28, oy + 62);
        gfx.DrawLine(thinPen, ox + 28, oy + 62, ox + 90, oy + 62);
        gfx.DrawArc(new XPen(XColors.DimGray, 0.5), ox + 20, oy + 50, 14, 14, 180, -90);
        var labelFont = new XFont("Arial", 7, XFontStyleEx.Regular);
        gfx.DrawString("leg", labelFont, XBrushes.DimGray, new XRect(ox + 2, oy + 30, 20, 12), XStringFormats.TopLeft);
        gfx.DrawString("leg", labelFont, XBrushes.DimGray, new XRect(ox + 52, oy + 72, 20, 12), XStringFormats.TopLeft);
    }

    private static void DrawBendJ(XGraphics gfx, XPen pen, double ox, double oy)
    {
        gfx.DrawLine(pen, ox + 40, oy, ox + 40, oy + 70);
        gfx.DrawArc(pen, ox + 20, oy + 50, 40, 40, 0, 180);
        gfx.DrawLine(pen, ox + 20, oy + 70, ox + 20, oy + 40);
        var thinPen = new XPen(XColors.Gray, 0.75);
        gfx.DrawLine(thinPen, ox + 32, oy + 2, ox + 32, oy + 56);
        gfx.DrawArc(thinPen, ox + 26, oy + 52, 24, 24, 0, 180);
        gfx.DrawLine(thinPen, ox + 26, oy + 64, ox + 26, oy + 40);
        gfx.DrawString("hook", new XFont("Arial", 7, XFontStyleEx.Regular), XBrushes.DimGray,
            new XRect(ox + 4, oy + 80, 30, 12), XStringFormats.TopLeft);
    }

    private static void DrawBendZ(XGraphics gfx, XPen pen, double ox, double oy)
    {
        gfx.DrawLine(pen, ox + 60, oy + 10, ox + 110, oy + 10);
        gfx.DrawLine(pen, ox + 60, oy + 18, ox + 110, oy + 18);
        gfx.DrawLine(pen, ox + 68, oy + 18, ox + 42, oy + 52);
        gfx.DrawLine(pen, ox + 60, oy + 18, ox + 34, oy + 52);
        gfx.DrawLine(pen, ox + 10, oy + 52, ox + 60, oy + 52);
        gfx.DrawLine(pen, ox + 10, oy + 60, ox + 60, oy + 60);
        gfx.DrawString("offset", new XFont("Arial", 7, XFontStyleEx.Regular), XBrushes.DimGray,
            new XRect(ox + 30, oy + 28, 40, 12), XStringFormats.TopLeft);
    }

    private static void DrawDimArrow(
        XGraphics gfx, XPen pen, XFont font,
        double x1, double y, double x2, double y2, string label)
    {
        var dimPen = new XPen(XColors.DimGray, 0.75);
        gfx.DrawLine(dimPen, x1, y, x2, y2);
        double ah = 5;
        gfx.DrawLine(dimPen, x1, y, x1 + ah, y - ah / 2);
        gfx.DrawLine(dimPen, x1, y, x1 + ah, y + ah / 2);
        gfx.DrawLine(dimPen, x2, y2, x2 - ah, y2 - ah / 2);
        gfx.DrawLine(dimPen, x2, y2, x2 - ah, y2 + ah / 2);
        gfx.DrawString(label, font, XBrushes.DimGray,
            new XRect((x1 + x2) / 2 - 15, y - 14, 30, 12), XStringFormats.TopCenter);
    }

    // ── Legacy public methods (kept for isolated call-sites) ──────────────────

    /// <summary>
    /// Extracts the ship-to customer name from the picking slip without modifying the PDF.
    /// Prefer EnrichPickingSlip for full enrichment to avoid redundant PDF opens.
    /// </summary>
    public static string? ExtractShipToName(byte[] pdfBytes, ILogger? log = null)
    {
        using var pigDoc = PigDoc.Open(pdfBytes);
        return ExtractShipToNameFromWords(pigDoc.GetPage(1).GetWords().ToList(), log);
    }

    /// <summary>Bolds comment lines and appends callout page. Legacy — use EnrichPickingSlip.</summary>
    public static byte[] Enrich(byte[] pdfBytes, ILogger? log = null)
    {
        List<CommentBlock> blocks;
        using (var pigDoc = PigDoc.Open(pdfBytes))
            blocks = ParseCommentBlocksFromDoc(pigDoc, log);

        if (blocks.Count == 0) return pdfBytes;

        var keywords = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        foreach (var b in blocks)
            foreach (var kw in b.Keywords)
                keywords.Add(kw);

        using var ms  = new MemoryStream(pdfBytes);
        var doc       = PdfReader.Open(ms, PdfDocumentOpenMode.Modify);
        BoldCommentsOnDoc(doc, blocks);
        if (keywords.Count > 0) AppendCalloutPageToDoc(doc, keywords);
        using var outMs = new MemoryStream();
        doc.Save(outMs);
        return outMs.ToArray();
    }

    /// <summary>Stamps the rep's first name over the Customer Rep / Store# area. Legacy — use EnrichPickingSlip.</summary>
    public static byte[] StampRepName(byte[] pdfBytes, ILogger? log = null)
    {
        string? repFirstName = null;
        double  coverL = double.MaxValue, coverB = double.MaxValue;
        double  coverR = double.MinValue, coverT = double.MinValue;
        double  repLineH = 10;

        using (var pigDoc = PigDoc.Open(pdfBytes))
        {
            var lines = GroupIntoLines(pigDoc.GetPage(1).GetWords().ToList());
            foreach (var line in lines)
            {
                bool isRepLine   = line.Text.IndexOf("Customer Rep:", StringComparison.OrdinalIgnoreCase) >= 0;
                bool isStoreLine = line.Text.IndexOf("Store #",       StringComparison.OrdinalIgnoreCase) >= 0;
                if (isRepLine)
                {
                    int sep  = line.Text.IndexOf("Customer Rep:", StringComparison.OrdinalIgnoreCase) + "Customer Rep:".Length;
                    var after = line.Text[sep..].Trim();
                    if (!string.IsNullOrWhiteSpace(after))
                        repFirstName = after.Split(' ', StringSplitOptions.RemoveEmptyEntries)[0];
                    repLineH = line.Height;
                }
                if (isRepLine || isStoreLine)
                {
                    coverL = Math.Min(coverL, line.X);
                    coverB = Math.Min(coverB, line.Y);
                    coverR = Math.Max(coverR, line.X + line.Width);
                    coverT = Math.Max(coverT, line.Y + line.Height);
                }
            }
        }

        log?.LogInformation("[PSE] StampRepName: rep='{Rep}' cover=({L:F1},{B:F1})–({R:F1},{T:F1})",
            repFirstName ?? "(null)", coverL, coverB, coverR, coverT);
        if (repFirstName is null || coverL == double.MaxValue) return pdfBytes;

        EnsureFontResolver();
        using var ms  = new MemoryStream(pdfBytes);
        var doc       = PdfReader.Open(ms, PdfDocumentOpenMode.Modify);
        StampRepNameOnDoc(doc, repFirstName, coverL, coverB, coverR, coverT, repLineH, log);
        using var outMs = new MemoryStream();
        doc.Save(outMs);
        return outMs.ToArray();
    }
}
