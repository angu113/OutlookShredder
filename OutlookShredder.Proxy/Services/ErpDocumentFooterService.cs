using Microsoft.Extensions.Logging;
using PdfSharp.Drawing;
using PdfSharp.Pdf;
using PdfSharp.Pdf.IO;
using SharpDoc = PdfSharp.Pdf.PdfDocument;

namespace OutlookShredder.Proxy.Services;

/// <summary>
/// Stamps a configurable terms-and-conditions footer onto outbound ERP documents
/// (Sales Orders / Quotations). The footer is drawn as a solid white box at the bottom
/// of the page — the white fill blocks out whatever ERP boilerplate sits there — with the
/// disclaimer text wrapped inside as a small footnote.
///
/// Mirrors <see cref="PickingSlipEnricher"/>'s PdfSharp write pattern and reuses its
/// Arial font resolver. Stateless / pure: takes PDF bytes + options, returns new bytes.
/// </summary>
internal static class ErpDocumentFooterService
{
    /// <summary>Tunable footer geometry + text. All measurements are PDF points (72 = 1 inch).</summary>
    public sealed record FooterOptions
    {
        /// <summary>Footer body text. Wrapped to fit the box width.</summary>
        public required string Text { get; init; }

        /// <summary>Font size of the footnote text.</summary>
        public double FontSizePt { get; init; } = 7.0;

        /// <summary>Left/right page margin the white box leaves uncovered.</summary>
        public double SideMarginPt { get; init; } = 36.0;   // 0.5"

        /// <summary>Gap between the bottom of the white box and the page edge.</summary>
        public double BottomMarginPt { get; init; } = 14.0;

        /// <summary>Height of the white box (and so how much existing text it blocks out).</summary>
        public double BoxHeightPt { get; init; } = 30.0;

        /// <summary>Stamp every page. When false, only the last page is stamped.</summary>
        public bool EveryPage { get; init; } = true;

        /// <summary>When set, stamp ONLY this 0-based page index (overrides <see cref="EveryPage"/>).</summary>
        public int? OnlyPageIndex { get; init; }

        /// <summary>Draw a thin rule along the top edge of the box to separate it from content above.</summary>
        public bool TopRule { get; init; } = true;

        /// <summary>Centre the text horizontally; otherwise left-align.</summary>
        public bool Center { get; init; } = true;
    }

    /// <summary>
    /// Returns the PDF bytes with the footer stamped. Returns the input array unchanged
    /// when the text is blank or the document has no pages.
    /// </summary>
    public static byte[] StampFooter(byte[] pdfBytes, FooterOptions opts, ILogger? log = null)
    {
        if (pdfBytes is null || pdfBytes.Length == 0) return pdfBytes ?? [];
        if (string.IsNullOrWhiteSpace(opts.Text)) return pdfBytes;

        // Reuse the picking-slip enricher's Arial resolver so PdfSharp can embed a font in a
        // service process with no GDI. EnsureFontResolver only sets the global resolver once.
        PickingSlipEnricher.EnsureFontResolver();

        using var ms = new MemoryStream(pdfBytes);
        SharpDoc doc;
        try
        {
            doc = PdfReader.Open(ms, PdfDocumentOpenMode.Modify);
        }
        catch (Exception ex)
        {
            log?.LogWarning(ex, "[Footer] Could not open PDF for stamping — returning original");
            return pdfBytes;
        }

        int pageCount = doc.PageCount;
        if (pageCount == 0) return pdfBytes;

        // Resolve the page range to stamp.
        int from, to;
        if (opts.OnlyPageIndex is int idx)
        {
            from = to = Math.Clamp(idx, 0, pageCount - 1);
        }
        else if (opts.EveryPage)
        {
            from = 0;
            to = pageCount - 1;
        }
        else
        {
            from = to = pageCount - 1;   // last page only
        }

        for (int p = from; p <= to; p++)
        {
            var page = doc.Pages[p];

            double pageW = page.Width.Point;
            double pageH = page.Height.Point;

            double boxX = opts.SideMarginPt;
            double boxW = Math.Max(0, pageW - opts.SideMarginPt * 2);
            double boxY = pageH - opts.BottomMarginPt - opts.BoxHeightPt;
            if (boxW <= 0 || boxY < 0) continue;   // page too small for these margins

            var box = new XRect(boxX, boxY, boxW, opts.BoxHeightPt);

            // Scope the XGraphics per page so it is flushed + disposed before doc.Save.
            // Draws with gfx.DrawString primitives (matching PickingSlipEnricher) — XTextFormatter
            // left the page content stream in a state that broke the subsequent doc.Save.
            using (var gfx = XGraphics.FromPdfPage(page))
            {
                var font = new XFont("Arial", opts.FontSizePt, XFontStyleEx.Regular);

                // Solid white fill blocks any existing text underneath.
                gfx.DrawRectangle(XBrushes.White, box);

                if (opts.TopRule)
                    gfx.DrawLine(new XPen(XColors.Black, 0.5), boxX, boxY, boxX + boxW, boxY);

                const double pad = 4.0;
                double innerW = boxW - pad * 2;
                double lineH  = font.GetHeight();
                var lines = WrapText(gfx, opts.Text, font, innerW);

                // Vertically centre the wrapped block within the box.
                double blockH = lines.Count * lineH;
                double y = boxY + Math.Max(pad * 0.5, (opts.BoxHeightPt - blockH) / 2.0);
                var fmt = opts.Center ? XStringFormats.TopCenter : XStringFormats.TopLeft;

                foreach (var line in lines)
                {
                    gfx.DrawString(line, font, XBrushes.Black,
                        new XRect(boxX + pad, y, innerW, lineH), fmt);
                    y += lineH;
                }
            }
        }

        using var outMs = new MemoryStream();
        doc.Save(outMs);
        log?.LogInformation("[Footer] Stamped footer on page(s) {From}-{To} of {Total}",
            from, to, pageCount);
        return outMs.ToArray();
    }

    /// <summary>
    /// Greedy word-wrap of <paramref name="text"/> into lines that each fit within
    /// <paramref name="maxWidth"/> points, measured with the given graphics + font.
    /// A single word longer than the line is emitted on its own (not broken mid-word).
    /// </summary>
    private static List<string> WrapText(XGraphics gfx, string text, XFont font, double maxWidth)
    {
        var lines = new List<string>();
        var line = new System.Text.StringBuilder();

        foreach (var word in text.Split(' ', StringSplitOptions.RemoveEmptyEntries))
        {
            var candidate = line.Length == 0 ? word : line + " " + word;
            if (line.Length > 0 && gfx.MeasureString(candidate, font).Width > maxWidth)
            {
                lines.Add(line.ToString());
                line.Clear();
                line.Append(word);
            }
            else
            {
                if (line.Length > 0) line.Append(' ');
                line.Append(word);
            }
        }
        if (line.Length > 0) lines.Add(line.ToString());
        return lines.Count == 0 ? new List<string> { text } : lines;
    }
}
