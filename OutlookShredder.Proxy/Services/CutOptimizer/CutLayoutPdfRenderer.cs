using System.Globalization;
using OutlookShredder.Proxy.Models.CutOptimizer;
using OutlookShredder.Proxy.Services;
using OutlookShredder.Proxy.Services.Drawing;
using PdfSharp.Drawing;
using PdfSharp.Pdf;

namespace OutlookShredder.Proxy.Services.CutOptimizer;

/// <summary>
/// Renders the cut-optimization plan to a PDF report, reusing the Pixar PDF chrome (font resolver +
/// franchise logo). Flat sheets draw as scaled rectangles with placed parts + shaded drop and the
/// ordered guillotine cut lines (shear); long bars draw as horizontal strips segmented by cut with the
/// end drop shaded. Self-contained on purpose — it does NOT touch the active DrawingPdfRenderer chrome
/// (mid-calibration), only the two stable public statics. See <c>wip/feat-cut-optimizer.md</c> Phase 3.
/// </summary>
public static class CutLayoutPdfRenderer
{
    private static readonly XColor Ink = XColor.FromArgb(30, 30, 30);
    private static readonly XColor Faint = XColor.FromArgb(150, 150, 150);
    private static readonly XColor PartFill = XColor.FromArgb(220, 232, 246);   // light blue
    private static readonly XColor PartEdge = XColor.FromArgb(70, 110, 170);
    private static readonly XColor DropFill = XColor.FromArgb(238, 238, 238);   // light grey waste
    private static readonly XColor CutLine = XColor.FromArgb(200, 60, 60);      // guillotine cut (shear)
    private static readonly XColor Accent = XColor.FromArgb(0x7C, 0x4D, 0xFF);

    public static byte[] Render(OptimizeResult result, MaterialForm form, CutMethod method, bool precision)
    {
        PickingSlipEnricher.EnsureFontResolver();
        var doc = new PdfDocument();

        const double M = 36;                                   // page margin (pt)
        bool flat = form == MaterialForm.Flat;
        int cols = flat ? 2 : 1;
        int rows = flat ? 2 : 8;

        var titleFont = new XFont("Arial", 15, XFontStyleEx.Bold);
        var subFont = new XFont("Arial", 8.5, XFontStyleEx.Regular);
        var capFont = new XFont("Arial", 8.5, XFontStyleEx.Bold);

        string subtitle = $"{(flat ? "Flat sheet" : "Long stock")}  •  " +
                          (flat ? (method == CutMethod.Laser ? "Laser (free nest)" : "Shear (guillotine)")
                                : (precision ? "Precision (1/8\" kerf/cut)" : "No kerf"));

        var layouts = result.Layouts;
        int idx = 0, page = 0;
        int totalPages = 0;   // unknown up front; stamped as "Page n"

        while (idx < layouts.Count || page == 0)
        {
            page++;
            var pg = doc.AddPage();
            pg.Width = XUnit.FromInch(11);
            pg.Height = XUnit.FromInch(8.5);
            using var gfx = XGraphics.FromPdfPage(pg);
            double pw = pg.Width.Point, phh = pg.Height.Point;

            // ── header band ──
            gfx.DrawString("Cut Optimization Report", titleFont, new XSolidBrush(Ink),
                new XRect(M, M - 12, pw - 2 * M - 120, 20), XStringFormats.TopLeft);
            gfx.DrawString(subtitle, subFont, new XSolidBrush(Faint),
                new XRect(M, M + 8, pw - 2 * M - 120, 14), XStringFormats.TopLeft);
            DrawLogo(gfx, new XRect(pw - M - 110, M - 14, 110, 30));
            double bodyTop = M + 28;

            // Page 1: summary block above the diagrams.
            if (page == 1)
                bodyTop = DrawSummary(gfx, result, M, bodyTop, pw - 2 * M) + 8;

            double footH = 16;
            double bodyBottom = phh - M - footH;
            DrawFooter(gfx, new XRect(M, bodyBottom + 2, pw - 2 * M, footH), page);

            // ── diagram grid ──
            double gx = M, gy = bodyTop, gw = pw - 2 * M, gh = bodyBottom - bodyTop;
            double cellW = (gw - (cols - 1) * 14) / cols;
            double cellH = (gh - (rows - 1) * 10) / rows;

            for (int r = 0; r < rows && idx < layouts.Count; r++)
            for (int c = 0; c < cols && idx < layouts.Count; c++)
            {
                var cell = new XRect(gx + c * (cellW + 14), gy + r * (cellH + 10), cellW, cellH);
                var lay = layouts[idx];
                string cap = $"{(flat ? "Sheet" : "Bar")} {idx + 1}: {SizeLabel(lay)}  •  {lay.YieldPct:0.#}% yield" +
                             (lay.Purchased ? "  •  PURCHASE" : "");
                gfx.DrawString(cap, capFont, new XSolidBrush(lay.Purchased ? Accent : Ink),
                    new XRect(cell.X, cell.Y, cell.Width, 12), XStringFormats.TopLeft);
                var diag = new XRect(cell.X, cell.Y + 14, cell.Width, cell.Height - 14);
                if (flat) DrawSheet(gfx, diag, lay, method);
                else DrawBar(gfx, diag, lay);
                idx++;
            }

            if (idx >= layouts.Count) { totalPages = page; break; }
        }

        using var ms = new MemoryStream();
        doc.Save(ms);
        return ms.ToArray();
    }

    // ── summary block (stock used + yield + purchase + issues) ──
    private static double DrawSummary(XGraphics gfx, OptimizeResult result, double x, double y, double w)
    {
        var font = new XFont("Arial", 8.5, XFontStyleEx.Regular);
        var boldF = new XFont("Arial", 8.5, XFontStyleEx.Bold);
        var pen = new XPen(XColor.FromArgb(210, 210, 210), 0.6);
        double ry = y + 4;

        foreach (var u in result.Usage)
        {
            string line = $"{Group(u.Material, u.Gauge)}:  {u.Count} x {SizeText(u.Width, u.Length)}  ({u.YieldPct:0.#}% yield)";
            gfx.DrawString(line, font, new XSolidBrush(Ink), new XRect(x + 4, ry, w - 8, 12), XStringFormats.TopLeft);
            ry += 12;
        }
        foreach (var p in result.ToPurchase)
        {
            string line = $"To purchase:  {p.Count} x {Group(p.Material, p.Gauge)} {SizeText(p.Width, p.Length)}";
            gfx.DrawString(line, boldF, new XSolidBrush(Accent), new XRect(x + 4, ry, w - 8, 12), XStringFormats.TopLeft);
            ry += 12;
        }
        foreach (var i in result.Issues)
        {
            gfx.DrawString("! " + i.Message, font, new XSolidBrush(CutLine), new XRect(x + 4, ry, w - 8, 12), XStringFormats.TopLeft);
            ry += 12;
        }
        gfx.DrawLine(pen, x, ry + 2, x + w, ry + 2);
        return ry + 2;
    }

    // ── flat sheet diagram ──
    private static void DrawSheet(XGraphics gfx, XRect area, Layout lay, CutMethod method)
    {
        double sw = lay.StockWidth ?? 0, sl = lay.StockLength;
        if (sw <= 0 || sl <= 0) return;
        // Fit the sheet (width across, length down) into the area, preserving aspect.
        double scale = Math.Min(area.Width / sw, area.Height / sl);
        double dw = sw * scale, dh = sl * scale;
        double ox = area.X + (area.Width - dw) / 2, oy = area.Y;

        // Sheet background = drop (grey); parts overpaint as light blue.
        gfx.DrawRectangle(new XPen(Ink, 0.8), new XSolidBrush(DropFill), ox, oy, dw, dh);

        var lblFont = new XFont("Arial", 6.5, XFontStyleEx.Regular);
        foreach (var p in lay.Pieces)
        {
            double px = ox + p.X * scale, py = oy + p.Y * scale, pwd = p.W * scale, phd = p.L * scale;
            gfx.DrawRectangle(new XPen(PartEdge, 0.6), new XSolidBrush(PartFill), px, py, pwd, phd);
            string t = $"{Frac(p.W)}x{Frac(p.L)}" + (p.Rotated ? " R" : "");
            if (pwd > 22 && phd > 9)
                gfx.DrawString(t, lblFont, new XSolidBrush(Ink), new XRect(px, py + phd / 2 - 4, pwd, 8), XStringFormats.Center);
        }

        // Shear: overlay the ordered guillotine cut lines (dashed red, numbered).
        if (method == CutMethod.Shear)
            DrawGuillotineCuts(gfx, lay, ox, oy, scale, sw, sl);
    }

    // Extract guillotine cuts by recursively finding a full-span split that crosses no part, in
    // pre-order (cut the sheet, then each sub-piece). Draws them dashed + numbered.
    private static void DrawGuillotineCuts(XGraphics gfx, Layout lay, double ox, double oy, double scale, double sw, double sl)
    {
        var pen = new XPen(CutLine, 0.8) { DashStyle = XDashStyle.Dash };
        var numFont = new XFont("Arial", 6, XFontStyleEx.Bold);
        int n = 0;
        var rects = lay.Pieces.Select(p => (p.X, p.Y, X2: p.X + p.W, Y2: p.Y + p.L)).ToList();

        void Recurse(double x1, double y1, double x2, double y2, List<(double X, double Y, double X2, double Y2)> parts)
        {
            if (parts.Count <= 1) return;
            const double e = 1e-6;
            // Vertical cut candidates = part right edges strictly inside the region.
            foreach (var cx in parts.Select(p => p.X2).Where(v => v > x1 + e && v < x2 - e).Distinct().OrderBy(v => v))
                if (!parts.Any(p => p.X < cx - e && p.X2 > cx + e))   // nothing straddles the line
                {
                    var left = parts.Where(p => p.X2 <= cx + e).ToList();
                    var right = parts.Where(p => p.X >= cx - e).ToList();
                    if (left.Count > 0 && right.Count > 0)
                    {
                        gfx.DrawLine(pen, ox + cx * scale, oy + y1 * scale, ox + cx * scale, oy + y2 * scale);
                        Label(ox + cx * scale, oy + y1 * scale, ++n);
                        Recurse(x1, y1, cx, y2, left); Recurse(cx, y1, x2, y2, right);
                        return;
                    }
                }
            foreach (var cy in parts.Select(p => p.Y2).Where(v => v > y1 + e && v < y2 - e).Distinct().OrderBy(v => v))
                if (!parts.Any(p => p.Y < cy - e && p.Y2 > cy + e))
                {
                    var top = parts.Where(p => p.Y2 <= cy + e).ToList();
                    var bot = parts.Where(p => p.Y >= cy - e).ToList();
                    if (top.Count > 0 && bot.Count > 0)
                    {
                        gfx.DrawLine(pen, ox + x1 * scale, oy + cy * scale, ox + x2 * scale, oy + cy * scale);
                        Label(ox + x1 * scale, oy + cy * scale, ++n);
                        Recurse(x1, y1, x2, cy, top); Recurse(x1, cy, x2, y2, bot);
                        return;
                    }
                }
        }
        void Label(double lx, double ly, int k) =>
            gfx.DrawString(k.ToString(), numFont, new XSolidBrush(CutLine), new XRect(lx + 1, ly + 1, 12, 8), XStringFormats.TopLeft);

        Recurse(0, 0, sw, sl, rects);
    }

    // ── long bar diagram ──
    private static void DrawBar(XGraphics gfx, XRect area, Layout lay)
    {
        double sl = lay.StockLength;
        if (sl <= 0) return;
        double scale = area.Width / sl;
        double barH = Math.Min(area.Height - 14, 26);
        double oy = area.Y + 2, ox = area.X;

        gfx.DrawRectangle(new XPen(Ink, 0.8), new XSolidBrush(DropFill), ox, oy, sl * scale, barH);
        var lblFont = new XFont("Arial", 6.5, XFontStyleEx.Regular);

        double cursor = 0;
        foreach (var p in lay.Pieces)
        {
            double segW = p.Length * scale;
            gfx.DrawRectangle(new XPen(PartEdge, 0.6), new XSolidBrush(PartFill), ox + cursor * scale, oy, segW, barH);
            if (segW > 16)
                gfx.DrawString(Frac(p.Length), lblFont, new XSolidBrush(Ink),
                    new XRect(ox + cursor * scale, oy + barH / 2 - 4, segW, 8), XStringFormats.Center);
            cursor += p.Length;
        }
        if (lay.Drop > 1e-6)
        {
            double dropW = lay.Drop * scale;
            gfx.DrawString($"drop {Frac(lay.Drop)}", lblFont, new XSolidBrush(Faint),
                new XRect(ox + cursor * scale, oy + barH / 2 - 4, Math.Max(dropW, 40), 8), XStringFormats.Center);
        }
    }

    // ── chrome bits ──
    private static void DrawFooter(XGraphics gfx, XRect box, int page)
    {
        var font = new XFont("Arial", 7, XFontStyleEx.Regular);
        string left = $"Created by {Cap(Environment.UserName)} • {DateTime.Now:MMM d, yyyy} • NOT TO SCALE";
        gfx.DrawString(left, font, new XSolidBrush(Faint), new XRect(box.X, box.Y, box.Width - 60, box.Height), XStringFormats.CenterLeft);
        gfx.DrawString($"Page {page}", font, new XSolidBrush(Faint), new XRect(box.X, box.Y, box.Width, box.Height), XStringFormats.CenterRight);
    }

    private static void DrawLogo(XGraphics gfx, XRect box)
    {
        string? path = new[]
        {
            DrawingPdfRenderer.LogoPath,
            Path.Combine(AppContext.BaseDirectory, "Assets", "ms-primary-logo.png"),
        }
        .Where(c => !string.IsNullOrWhiteSpace(c))
        .Select(c => Environment.ExpandEnvironmentVariables(c!))
        .FirstOrDefault(File.Exists);
        if (path is null) return;
        try
        {
            var img = XImage.FromFile(path);
            double sc = Math.Min(box.Width / img.PixelWidth, box.Height / img.PixelHeight);
            double w = img.PixelWidth * sc, h = img.PixelHeight * sc;
            gfx.DrawImage(img, box.X + box.Width - w, box.Y + (box.Height - h) / 2, w, h);
        }
        catch { /* missing/unsupported logo — skip */ }
    }

    private static string SizeLabel(Layout l) => SizeText(l.StockWidth, l.StockLength);
    private static string SizeText(double? width, double length) =>
        width is double w && w > 0 ? $"{Frac(w)} x {Frac(length)}" : Frac(length);
    private static string Frac(double v) => DrawFormat.FracInch(v);
    private static string Group(string m, string g) => string.IsNullOrWhiteSpace(g) ? m : $"{m} {g}";
    private static string Cap(string s) => string.IsNullOrEmpty(s) ? s : char.ToUpper(s[0]) + s[1..];
}
