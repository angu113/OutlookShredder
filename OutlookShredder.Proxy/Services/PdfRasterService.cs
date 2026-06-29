using PDFtoImage;
using SkiaSharp;

namespace OutlookShredder.Proxy.Services;

/// <summary>
/// Rasterizes a PDF's pages to JPEGs for outbound MMS — SignalWire MMS accepts images, NOT application/pdf, so a
/// PDF attachment is "printed" to one image per page and those images are sent. Uses PDFium (via PDFtoImage) +
/// SkiaSharp. Rendering DPI + JPEG quality are tuned to keep each page well under the MMS per-message size
/// budget while staying legible. CPU-bound + synchronous; callers run it off the request path as needed.
/// </summary>
public sealed class PdfRasterService
{
    private readonly ILogger<PdfRasterService> _log;
    public PdfRasterService(ILogger<PdfRasterService> log) => _log = log;

    /// <summary>One rendered page: its 1-based number and JPEG bytes.</summary>
    public sealed record Page(int Number, byte[] Jpeg);

    /// <summary>Renders up to <paramref name="maxPages"/> pages to JPEG bytes. Returns empty on a non-PDF / a
    /// render failure (the caller surfaces that to the operator). ~130 DPI on a Letter page is ~1100x1430 px,
    /// JPEG q75 ≈ 100–250 KB/page — comfortably inside the ~1.2 MB MMS message cap.</summary>
    public IReadOnlyList<Page> RenderToJpegs(byte[] pdfBytes, int maxPages = 20, int dpi = 130, int quality = 75)
    {
        var pages = new List<Page>();
        if (pdfBytes is null || pdfBytes.Length == 0) return pages;
        try
        {
            var n = 0;
            foreach (var bmp in Conversion.ToImages(pdfBytes, options: new RenderOptions(Dpi: dpi)))
            {
                using (bmp)
                {
                    if (n >= maxPages)
                    {
                        _log.LogWarning("[PdfRaster] PDF exceeds {Max} pages — sending the first {Max} only", maxPages, maxPages);
                        break;
                    }
                    using var data = bmp.Encode(SKEncodedImageFormat.Jpeg, quality);
                    pages.Add(new Page(n + 1, data.ToArray()));
                    n++;
                }
            }
        }
        catch (Exception ex)
        {
            _log.LogWarning(ex, "[PdfRaster] render failed");
            return [];
        }
        return pages;
    }
}
