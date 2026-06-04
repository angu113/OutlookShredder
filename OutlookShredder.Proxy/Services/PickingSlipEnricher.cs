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
///   1. Stamps both header cells: left = customer name + attention + contact phone;
///      right = sales rep + order date + delivery method + carrier + PO# (if present).
///   2. Bolds shop-comment lines in-place (white rect + bold redraw).
///   3. Appends a callout page with geometric drawings for bend/cut keywords.
///
/// EnrichPickingSlip does all three in a single PdfPig read + single PdfSharp write.
/// </summary>
internal static class PickingSlipEnricher
{
    // ── Font resolver ─────────────────────────────────────────────────────────

    private static readonly object _fontLock = new();
    private static bool _fontResolverSet;

    /// <summary>
    /// Sets the global Arial font resolver exactly once. Shared with
    /// <see cref="ErpDocumentFooterService"/> so both PDF-write paths embed fonts identically.
    /// </summary>
    internal static void EnsureFontResolver()
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

    // Numeric size code after the slash may be a single digit (e.g. HR/1), so allow \d{1,8};
    // requiring 2+ digits silently dropped single-digit-MSPC products from detection.
    private static readonly Regex _mspcRegex =
        new(@"^[A-Z][A-Z0-9]{0,9}/\d{1,8}", RegexOptions.Compiled);

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

    /// <summary>All header fields extracted from page 1 of the picking slip.</summary>
    private sealed record SlipHeader(
        string? CustomerName,
        string? Attention,      // "Attention: "
        string? ContactPhone,   // "Contact Phone: "
        string? RepName,        // "Customer Rep: " (full name)
        string? OrderDate,      // "Order Date: "
        string? DeliveryMethod, // "Delivery Method: "
        string? Carrier,        // "Carrier: "
        string? PoNumber);      // "Customer Purchase Order #"

    // ── Combined entry point (1 PdfPig read + 1 PdfSharp write) ──────────────

    /// <summary>
    /// Applies all picking-slip enrichments in two PDF passes (one read, one write).
    /// Returns the enriched bytes, the ship-to customer name extracted from the PDF
    /// (null when the Ship To label is not found), and any matched processing-operation
    /// keywords found in B: shop-comment lines (used by the Trigger Prioritize auto-add rule).
    /// </summary>
    public static (byte[] Bytes, string? ShipToName, IReadOnlyList<string> ProcessOps) EnrichPickingSlip(
        byte[] pdfBytes,
        string? knownCustomerName = null,
        ILogger? log = null,
        IReadOnlyList<string>? processingKeywords = null)
    {
        EnsureFontResolver();

        // ── PdfPig pass: extract everything ──────────────────────────────────
        string? shipToName = null;
        SlipHeader header;
        double? hdrPsTop    = null;
        double? hdrPsHeight = null;
        List<CommentBlock> blocks;
        List<string> allLinesAcrossPages = [];
        int pdfPageCount;

        using (var pigDoc = PigDoc.Open(pdfBytes))
        {
            pdfPageCount   = pigDoc.NumberOfPages;
            var page1      = pigDoc.GetPage(1);
            double pigPageH = page1.Height;
            double pigPageW = page1.Width;
            var page1Words  = page1.GetWords().ToList();

            // 1. Ship-to customer name
            shipToName = ExtractShipToNameFromWords(page1Words, log);

            // 2. All labeled header fields + header cell bounds
            var page1Lines = GroupIntoLines(page1Words);
            (header, hdrPsTop, hdrPsHeight) = ExtractHeaderFields(
                page1Lines, pigPageW, pigPageH,
                shipToName ?? knownCustomerName, log);

            // 3. Comment blocks from all pages, plus the flat list of every text line on every page
            //    (used for the broader processing-keyword scan below — the keyword list intentionally
            //    targets words that appear as line-item product names, not only B: comments).
            blocks = ParseCommentBlocksFromDoc(pigDoc, log, out allLinesAcrossPages);
        }

        bool hasHeaderBounds = hdrPsTop.HasValue && hdrPsHeight is > 0;

        var keywords = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        foreach (var b in blocks)
            foreach (var kw in b.Keywords)
                keywords.Add(kw);

        // Processing-operation keyword scan over EVERY text line on every page.
        // Keywords (Laser Cutting / Bending / Welding / Drilling / Fabricating) typically appear
        // either as B: shop-comment lines OR as service-line-items like "FABRICATING SERVICES"
        // — broader scan keeps the rule simple and matches user expectation.
        var processOps = ScanProcessingKeywords(allLinesAcrossPages, processingKeywords, log);

        if (!hasHeaderBounds && blocks.Count == 0 && pdfPageCount <= 1)
            return (pdfBytes, shipToName, processOps);

        // ── PdfSharp pass: apply all modifications in one shot ────────────────
        using var ms  = new MemoryStream(pdfBytes);
        var doc       = PdfReader.Open(ms, PdfDocumentOpenMode.Modify);

        if (hasHeaderBounds)
        {
            double pageW = doc.Pages[0].Width.Point;
            StampLeftCellOnDoc(doc,  header, hdrPsTop!.Value, hdrPsHeight!.Value, pageW, log);
            StampRightCellOnDoc(doc, header, hdrPsTop!.Value, hdrPsHeight!.Value, pageW, log);
        }

        if (blocks.Count > 0)
            BoldCommentsOnDoc(doc, blocks);

        StampPageNumbersOnDoc(doc);

        if (keywords.Count > 0)
            AppendCalloutPageToDoc(doc, keywords);

        using var outMs = new MemoryStream();
        doc.Save(outMs);
        return (outMs.ToArray(), shipToName, processOps);
    }

    private static IReadOnlyList<string> ScanProcessingKeywords(
        IReadOnlyList<string> textLines,
        IReadOnlyList<string>? keywords,
        ILogger? log)
    {
        if (keywords is null || keywords.Count == 0 || textLines.Count == 0)
            return [];

        var matched = new List<string>();
        foreach (var kw in keywords)
        {
            if (string.IsNullOrWhiteSpace(kw)) continue;
            var hitLine = textLines.FirstOrDefault(l => l.Contains(kw, StringComparison.OrdinalIgnoreCase));
            if (hitLine is not null)
            {
                matched.Add(kw);
                log?.LogInformation("[PSE] Keyword '{Kw}' matched in line: '{Line}'", kw, hitLine);
            }
        }

        if (matched.Count > 0)
            log?.LogInformation("[PSE] Processing keywords matched: {Ops}", string.Join(", ", matched));
        return matched;
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

    /// <summary>
    /// Extracts all labeled header fields and computes the PdfSharp (top-left origin)
    /// bounding box for the two-column header cell area.
    /// Anchors on "PICKING SLIP" (cell top) and the "MSPC" column header (cell bottom).
    /// Returns (header, psTop, psHeight) where psTop/psHeight are null when anchors not found.
    /// </summary>
    private static (SlipHeader Header, double? PsTop, double? PsHeight) ExtractHeaderFields(
        List<TextLine> lines, double pigPageW, double pigPageH,
        string? customerName, ILogger? log)
    {
        string? attention      = null;
        string? contactPhone   = null;
        string? repName        = null;
        string? orderDate      = null;
        string? deliveryMethod = null;
        string? carrier        = null;
        string? poNumber       = null;

        TextLine? pickingSlipLine = null; // anchor: cell top is just below its bottom edge
        TextLine? mspcLine        = null; // anchor: cell bottom is just above its top edge

        static string? After(string text, string label)
        {
            var idx = text.IndexOf(label, StringComparison.OrdinalIgnoreCase);
            if (idx < 0) return null;
            var val = text[(idx + label.Length)..].TrimStart(' ', ':', '\t').Trim();
            return string.IsNullOrWhiteSpace(val) ? null : val;
        }

        foreach (var line in lines)
        {
            var text = line.Text;

            if (pickingSlipLine is null &&
                text.Contains("PICKING", StringComparison.OrdinalIgnoreCase) &&
                text.Contains("SLIP",    StringComparison.OrdinalIgnoreCase))
                pickingSlipLine = line;

            // "MSPC" is the leftmost column header; guard X < 30 to skip product-code
            // occurrences that contain "MSPC" as part of an item description.
            if (mspcLine is null &&
                text.StartsWith("MSPC", StringComparison.OrdinalIgnoreCase) &&
                line.X < 30.0)
                mspcLine = line;

            attention      ??= After(text, "Attention:");
            contactPhone   ??= After(text, "Contact Phone:");
            repName        ??= After(text, "Customer Rep:");
            orderDate      ??= After(text, "Order Date:");
            deliveryMethod ??= After(text, "Delivery Method:");
            carrier        ??= After(text, "Carrier:");
            poNumber       ??= After(text, "Customer Purchase Order #");
        }

        double? psTop    = null;
        double? psHeight = null;

        if (pickingSlipLine is not null && mspcLine is not null)
        {
            // PdfSharp Y of the bottom edge of "PICKING SLIP" text:
            //   psPickingSlipBottom = pigPageH - pickingSlipLine.Y
            // Cell top starts a few pt below that (between the text and the horizontal rule).
            const double topGap    = 5.0; // pt below "PICKING SLIP" bottom → cell top border
            const double bottomGap = 5.0; // pt above "MSPC" top → cell bottom border

            double psPickingSlipBottom = pigPageH - pickingSlipLine.Y;
            double psMspcTop           = pigPageH - (mspcLine.Y + mspcLine.Height);

            double top    = psPickingSlipBottom + topGap;
            double bottom = psMspcTop           - bottomGap;
            double height = bottom - top;

            if (height > 20)
            {
                psTop    = top;
                psHeight = height;
            }
        }

        log?.LogInformation(
            "[PSE] Header fields — cust='{C}' attn='{A}' phone='{P}' rep='{R}' date='{D}' del='{V}' carrier='{Ca}' po='{PO}'",
            customerName, attention, contactPhone, repName, orderDate, deliveryMethod, carrier, poNumber);

        if (psTop.HasValue)
            log?.LogInformation("[PSE] Header cell bounds — psTop={T:F1} psHeight={H:F1}", psTop, psHeight);
        else
            log?.LogWarning("[PSE] Header cell bounds not detected — PICKING SLIP={PS} MSPC={M}",
                pickingSlipLine?.Text ?? "(not found)", mspcLine?.Text ?? "(not found)");

        var h = new SlipHeader(customerName, attention?.Trim(), contactPhone?.Trim(),
            repName?.Trim(), orderDate?.Trim(), deliveryMethod?.Trim(),
            carrier?.Trim(), poNumber?.Trim());

        return (h, psTop, psHeight);
    }

    private static List<CommentBlock> ParseCommentBlocksFromDoc(PigDoc doc, ILogger? log)
        => ParseCommentBlocksFromDoc(doc, log, out _);

    private static List<CommentBlock> ParseCommentBlocksFromDoc(PigDoc doc, ILogger? log, out List<string> allTextLines)
    {
        var result = new List<CommentBlock>();
        allTextLines = [];
        foreach (var page in doc.GetPages())
        {
            int pageIdx = page.Number - 1;
            var rawLines = GroupIntoLines(page.GetWords().ToList());
            foreach (var l in rawLines)
                if (!string.IsNullOrWhiteSpace(l.Text)) allTextLines.Add(l.Text);
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

    /// <summary>
    /// White-outs the left header cell and redraws:
    ///   - Customer name (large bold, fitted)
    ///   - Attention: [name]    (12pt bold)
    ///   - Contact Phone: [num] (12pt bold)
    /// </summary>
    private static void StampLeftCellOnDoc(
        SharpDoc doc, SlipHeader h,
        double psTop, double psHeight, double pageW, ILogger? log)
    {
        var page = doc.Pages[0];
        using var gfx = XGraphics.FromPdfPage(page);
        const double margin = 8.0;
        double splitX = pageW / 2.0;

        // White-out existing cell content and redraw border
        var cellRect = new XRect(0, psTop, splitX, psHeight);
        gfx.DrawRectangle(XBrushes.White, cellRect);
        gfx.DrawRectangle(new XPen(XColors.Black, 0.5), cellRect);

        double innerX = margin;
        double innerW = splitX - margin * 2;
        double innerH = psHeight - margin * 2;

        // Reserve 16pt per contact row at the bottom
        const double rowH = 16.0;
        int extraRows = (h.Attention    != null ? 1 : 0) +
                        (h.ContactPhone != null ? 1 : 0);
        double nameH = Math.Max(innerH - extraRows * rowH, rowH);

        // Customer name — fitted bold, centred in name area
        if (!string.IsNullOrEmpty(h.CustomerName))
        {
            var (nameLines, nameFont) = FitTextInBox(gfx, h.CustomerName, innerW, nameH);
            double lh     = nameFont.GetHeight();
            double blockH = nameLines.Count * lh;
            double startY = psTop + margin + (nameH - blockH) / 2.0;
            foreach (var nl in nameLines)
            {
                gfx.DrawString(nl, nameFont, XBrushes.Black,
                    new XRect(innerX, startY, innerW, lh), XStringFormats.TopCenter);
                startY += lh;
            }
        }

        // Contact rows — 12pt bold, left-aligned below name area
        var smallFont = new XFont("Arial", 12, XFontStyleEx.Bold);
        double rowY = psTop + margin + nameH;
        if (!string.IsNullOrEmpty(h.Attention))
        {
            gfx.DrawString($"Attention: {h.Attention}", smallFont, XBrushes.Black,
                new XRect(innerX, rowY, innerW, rowH), XStringFormats.TopLeft);
            rowY += rowH;
        }
        if (!string.IsNullOrEmpty(h.ContactPhone))
        {
            gfx.DrawString($"Contact Phone: {h.ContactPhone}", smallFont, XBrushes.Black,
                new XRect(innerX, rowY, innerW, rowH), XStringFormats.TopLeft);
        }

        log?.LogInformation("[PSE] Left cell stamped — '{Cust}' / '{Attn}' / '{Phone}'",
            h.CustomerName, h.Attention, h.ContactPhone);
    }

    /// <summary>
    /// White-outs the right header cell and redraws:
    ///   Sales Rep / Order Date / Delivery Method / Carrier / PO# (omitted when null)
    /// Each row: bold label + regular value, vertically centred in the cell.
    /// Font size auto-shrinks to fit all rows.
    /// </summary>
    private static void StampRightCellOnDoc(
        SharpDoc doc, SlipHeader h,
        double psTop, double psHeight, double pageW, ILogger? log)
    {
        var page = doc.Pages[0];
        using var gfx = XGraphics.FromPdfPage(page);
        const double margin = 8.0;
        double splitX = pageW / 2.0;
        double cellW  = pageW - splitX;

        // White-out existing cell content and redraw border
        var cellRect = new XRect(splitX, psTop, cellW, psHeight);
        gfx.DrawRectangle(XBrushes.White, cellRect);
        gfx.DrawRectangle(new XPen(XColors.Black, 0.5), cellRect);

        var rows = new List<(string Label, string Value, bool ValueOnly, bool BoldValue)>();
        if (!string.IsNullOrEmpty(h.RepName))        rows.Add(("Sales Rep:",  h.RepName,        false, true));
        if (!string.IsNullOrEmpty(h.OrderDate))       rows.Add(("Order Date:", h.OrderDate,      false, false));
        if (!string.IsNullOrEmpty(h.DeliveryMethod))  rows.Add(("",            h.DeliveryMethod, true,  false));
        if (!string.IsNullOrEmpty(h.Carrier))         rows.Add(("Carrier:",    h.Carrier,        false, false));
        if (!string.IsNullOrEmpty(h.PoNumber))        rows.Add(("PO#:",        h.PoNumber,       false, false));
        if (rows.Count == 0) return;

        double innerW       = cellW  - margin * 2;
        double innerH       = psHeight - margin * 2;
        double rowH         = Math.Min(innerH / rows.Count, 18.0);
        double fontSize     = Math.Clamp(rowH * 0.65, 7.0, 11.0);
        var labelFont       = new XFont("Arial", fontSize,     XFontStyleEx.Bold);
        var valueFont       = new XFont("Arial", fontSize,     XFontStyleEx.Regular);
        var boldValueFont   = new XFont("Arial", fontSize + 2, XFontStyleEx.Bold);
        const double labelColW = 65.0;

        // Vertically centre the row block
        double blockH = rows.Count * rowH;
        double y = psTop + margin + (innerH - blockH) / 2.0;

        foreach (var (label, value, valueOnly, boldValue) in rows)
        {
            if (valueOnly)
            {
                // No label — align value with the value column start
                gfx.DrawString(value, labelFont, XBrushes.Black,
                    new XRect(splitX + margin + labelColW, y, innerW - labelColW, rowH), XStringFormats.TopLeft);
            }
            else
            {
                var vFont = boldValue ? boldValueFont : valueFont;
                gfx.DrawString(label, labelFont, XBrushes.Black,
                    new XRect(splitX + margin, y, labelColW, rowH), XStringFormats.TopLeft);
                gfx.DrawString(value, vFont, XBrushes.Black,
                    new XRect(splitX + margin + labelColW, y, innerW - labelColW, rowH), XStringFormats.TopLeft);
            }
            y += rowH;
        }

        log?.LogInformation("[PSE] Right cell stamped — rep='{R}' date='{D}' del='{V}' carrier='{C}' po='{P}'",
            h.RepName, h.OrderDate, h.DeliveryMethod, h.Carrier, h.PoNumber);
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

    private static void StampPageNumbersOnDoc(SharpDoc doc)
    {
        if (doc.Pages.Count <= 1) return;

        var textFont = new XFont("Arial", 16, XFontStyleEx.Bold);
        int total    = doc.Pages.Count;

        // Measure the widest possible label so all stamps are identical in size
        string widestLabel = $"Pág. {total} de {total}";
        XSize textSize;
        using (var mc = XGraphics.CreateMeasureContext(
                   new XSize(1000, 1000), XGraphicsUnit.Point, XPageDirection.Downwards))
            textSize = mc.MeasureString(widestLabel, textFont);

        const double hPad   = 8.0;   // horizontal padding each side
        const double vPad   = 6.0;   // vertical padding top and bottom
        const double gs     = 11.0;  // glyph size (triangle / square)
        const double gap    = 6.0;   // gap between glyph and text
        const double margin = 24.0;  // pt from page edges

        double boxW = hPad + gs + gap + textSize.Width + hPad;
        double boxH = Math.Max(gs, textSize.Height) + vPad * 2;
        var    border = new XPen(XColors.Black, 1.5);

        for (int i = 0; i < total; i++)
        {
            var page   = doc.Pages[i];
            double pageW = page.Width.Point;
            double pageH = page.Height.Point;
            using var gfx = XGraphics.FromPdfPage(page);

            double bx = pageW - boxW - margin;
            double by = pageH - boxH - margin;

            // Opaque white fill + black border — sized to content, no clipping
            gfx.DrawRectangle(border, XBrushes.White, new XRect(bx, by, boxW, boxH));

            bool   isLast = (i == total - 1);
            double glyphX = bx + hPad;
            double glyphY = by + (boxH - gs) / 2.0;

            if (isLast)
            {
                // Filled black square (last page)
                gfx.DrawRectangle(XBrushes.Black, new XRect(glyphX, glyphY, gs, gs));
            }
            else
            {
                // Filled right-pointing triangle (continue)
                var tri = new XPoint[]
                {
                    new XPoint(glyphX,      glyphY),
                    new XPoint(glyphX,      glyphY + gs),
                    new XPoint(glyphX + gs, glyphY + gs / 2),
                };
                gfx.DrawPolygon(XBrushes.Black, tri, XFillMode.Winding);
            }

            // Label: glyph to the left, text to the right, both vertically centred
            string label = $"Pág. {i + 1} de {total}";
            double textX = bx + hPad + gs + gap;
            double textY = by + (boxH - textSize.Height) / 2.0;
            gfx.DrawString(label, textFont, XBrushes.Black,
                new XRect(textX, textY, textSize.Width + 2, textSize.Height),
                XStringFormats.TopLeft);
        }
    }

    // ── Text layout helpers ───────────────────────────────────────────────────

    private static (List<string> Lines, XFont Font) FitTextInBox(
        XGraphics gfx, string text, double maxW, double maxH)
    {
        for (double size = 19; size >= 7; size -= 1)
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

    // ── Stamp-bounds detection (for WPF UI overlay) ───────────────────────────

    /// <summary>
    /// Finds the "Description (Special Instructions)" box on the first page that contains it
    /// Returns the dimensions of every page in the PDF, in points (1 pt = 1/72").
    /// Useful for diagnosing stamp coordinate mismatches.
    /// </summary>
    public static List<(double WidthPt, double HeightPt)> GetPageDimensions(byte[] pdfBytes)
    {
        using var pigDoc = PigDoc.Open(pdfBytes);
        var result = new List<(double, double)>();
        for (int p = 1; p <= pigDoc.NumberOfPages; p++)
        {
            var page = pigDoc.GetPage(p);
            result.Add((page.Width, page.Height));
        }
        return result;
    }

    /// (typically the last page of a multi-page picking slip).
    /// Returns (0-based page index, fractional bounds with top-left origin), or null when not found.
    /// </summary>
    public static (int PageIndex, double LeftFrac, double TopFrac, double WidthFrac, double HeightFrac)?
        ExtractDescriptionBoxBounds(byte[] pdfBytes)
    {
        using var pigDoc = PigDoc.Open(pdfBytes);

        for (int p = 1; p <= pigDoc.NumberOfPages; p++)
        {
            var page  = pigDoc.GetPage(p);
            double pigH = page.Height;
            var lines = GroupIntoLines(page.GetWords().ToList());

            foreach (var line in lines)
            {
                var t = line.Text.Trim();
                if (!t.StartsWith("Description", StringComparison.OrdinalIgnoreCase))
                    continue;

                // Box starts at the top edge of the label text and runs to the page bottom,
                // full width — so the label itself is covered and the entire area below is available.
                double psBoxTop = pigH - (line.Y + line.Height);
                double boxH     = pigH - psBoxTop; // = line.Y + line.Height

                return (p - 1, 0.0, psBoxTop / pigH, 1.0, boxH / pigH);
            }
        }

        return null;
    }

    // ── Product-box detection (for the WPF "+ Stock" decorators) ──────────────

    /// <summary>Anchor box for one product found on a PDF page (top-left origin fractions).</summary>
    public sealed record ProductBox(
        int PageIndex, double LeftFrac, double TopFrac, double WidthFrac, double HeightFrac,
        string ProductName, string? Mspc);

    /// <summary>A known line item from the ERP document, used to locate non-picking-slip products.</summary>
    public sealed record ProductHint(string? Code, string? Description);

    private static readonly HashSet<string> _productDocTypes =
        new(StringComparer.OrdinalIgnoreCase) { "PickingSlip", "Invoice", "SalesOrder", "Quotation" };

    /// <summary>
    /// Finds an anchor box per product on each page so the UI can render a "+ Stock" marker
    /// beside it. Picking slips are detected by the MSPC code at the start of each product row;
    /// invoices/sales-orders/quotes are matched against the document's known line items.
    /// Returns top-left-origin fractional bounds (same convention as ExtractDescriptionBoxBounds).
    /// </summary>
    public static List<ProductBox> ExtractProductBoxes(
        byte[] pdfBytes, string? docType, IReadOnlyList<ProductHint>? hints)
    {
        var result = new List<ProductBox>();
        if (docType is null || !_productDocTypes.Contains(docType))
            return result;

        bool isPickingSlip = docType.Equals("PickingSlip", StringComparison.OrdinalIgnoreCase);

        using var pigDoc = PigDoc.Open(pdfBytes);

        if (isPickingSlip)
        {
            for (int p = 1; p <= pigDoc.NumberOfPages; p++)
            {
                var page  = pigDoc.GetPage(p);
                double pigW = page.Width, pigH = page.Height;
                foreach (var line in GroupIntoLines(page.GetWords().ToList()))
                {
                    var text = line.Text.Trim();
                    if (!_mspcRegex.IsMatch(text)) continue;
                    var mspc = text.Split(' ', StringSplitOptions.RemoveEmptyEntries).FirstOrDefault();
                    result.Add(MakeBox(p - 1, line, pigW, pigH, text, mspc));
                }
            }
            return result;
        }

        // Invoice / SalesOrder / Quotation: match the document's known line items to PDF lines.
        if (hints is null || hints.Count == 0) return result;

        for (int p = 1; p <= pigDoc.NumberOfPages; p++)
        {
            var page  = pigDoc.GetPage(p);
            double pigW = page.Width, pigH = page.Height;
            var lines = GroupIntoLines(page.GetWords().ToList());
            var usedLineKeys = new HashSet<double>();

            foreach (var hint in hints)
            {
                var match = FindLineForHint(lines, hint, usedLineKeys);
                if (match is null) continue;
                usedLineKeys.Add(Math.Round(match.Y, 1));
                var name = !string.IsNullOrWhiteSpace(hint.Description) ? hint.Description!.Trim() : match.Text.Trim();
                result.Add(MakeBox(p - 1, match, pigW, pigH, name, null));
            }
        }
        return result;
    }

    private static ProductBox MakeBox(
        int pageIndex, TextLine line, double pigW, double pigH, string name, string? mspc)
    {
        double topFrac = (pigH - (line.Y + line.Height)) / pigH;
        return new ProductBox(
            pageIndex,
            line.X / pigW,
            topFrac,
            line.Width / pigW,
            line.Height / pigH,
            name,
            mspc);
    }

    private static TextLine? FindLineForHint(
        List<TextLine> lines, ProductHint hint, HashSet<double> usedLineKeys)
    {
        // Prefer an exact code hit (codes are distinctive enough to anchor on).
        if (!string.IsNullOrWhiteSpace(hint.Code) && hint.Code!.Trim().Length >= 3)
        {
            var code = hint.Code.Trim();
            var byCode = lines.FirstOrDefault(l =>
                !usedLineKeys.Contains(Math.Round(l.Y, 1)) &&
                l.Text.Contains(code, StringComparison.OrdinalIgnoreCase));
            if (byCode is not null) return byCode;
        }

        // Else fuzzy-match the description by shared significant tokens.
        var hintTokens = Tokenize(hint.Description);
        if (hintTokens.Count == 0) return null;

        TextLine? best = null;
        int bestShared = 0;
        foreach (var line in lines)
        {
            if (usedLineKeys.Contains(Math.Round(line.Y, 1))) continue;
            var lineTokens = Tokenize(line.Text);
            if (lineTokens.Count == 0) continue;
            int shared = hintTokens.Count(t => lineTokens.Contains(t));
            if (shared > bestShared) { bestShared = shared; best = line; }
        }
        // Require a couple of shared tokens so we don't anchor on a stray header word.
        return bestShared >= Math.Min(2, hintTokens.Count) ? best : null;
    }

    private static HashSet<string> Tokenize(string? text)
    {
        var set = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        if (string.IsNullOrWhiteSpace(text)) return set;
        foreach (var raw in text.Split((char[]?)null, StringSplitOptions.RemoveEmptyEntries))
        {
            var t = new string(raw.Where(char.IsLetterOrDigit).ToArray());
            if (t.Length >= 2) set.Add(t);
        }
        return set;
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

        log?.LogInformation("[PSE] StampRepName: rep='{Rep}' cover=({L:F1},{B:F1})-({R:F1},{T:F1})",
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
