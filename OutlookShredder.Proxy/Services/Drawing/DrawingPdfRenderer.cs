using System.Globalization;
using OutlookShredder.Proxy.Models.Drawing;
using PdfSharp.Drawing;
using PdfSharp.Pdf;

namespace OutlookShredder.Proxy.Services.Drawing;

/// <summary>
/// Renders a <see cref="FlatPatternResult"/> to a printable PDF: plain-English title, flat-blank
/// line, three panels (dimensioned flat pattern, radiused end-section, thick-walled isometric)
/// and a footnote box with the secondary detail. Works for U / L / Z (and any future shape that
/// supplies a profile). PdfSharp; reuses the shared Arial resolver.
/// </summary>
public static class DrawingPdfRenderer
{
    private static readonly XColor CutColor  = XColors.Black;
    private static readonly XColor BendColor = XColors.RoyalBlue;
    private static readonly XColor DimColor  = XColor.FromArgb(155, 155, 155);
    private static readonly XBrush DimBrush  = new XSolidBrush(DimColor);
    private static readonly XBrush TextBrush = XBrushes.Black;

    private const double ExtGap = 2.0;
    private const double ExtOver = 3.0;

    public static byte[] Render(FlatPatternResult fp)
    {
        PickingSlipEnricher.EnsureFontResolver();

        var doc = new PdfDocument();
        var page = doc.AddPage();
        page.Width = XUnit.FromInch(11);
        page.Height = XUnit.FromInch(8.5);

        using (var gfx = XGraphics.FromPdfPage(page))
        {
            double pw = page.Width.Point, ph = page.Height.Point;
            const double M = 36;

            var titleFont = new XFont("Arial", 17, XFontStyleEx.Bold);
            var blankFont = new XFont("Arial", 11, XFontStyleEx.Bold);

            gfx.DrawString(fp.Title, titleFont, XBrushes.Black,
                new XRect(M, M - 10, pw - 2 * M, 24), XStringFormats.TopLeft);
            gfx.DrawString($"Flat blank:  {F(fp.FlatWidth)}\" x {F(fp.FlatHeight)}\"",
                blankFont, XBrushes.Black, new XRect(M, M + 16, pw - 2 * M, 14), XStringFormats.TopLeft);

            double top = M + 38;
            double usable = pw - 2 * M;
            const double gap = 16;
            double wFlat = (usable - 2 * gap) * 0.36;
            double wSect = (usable - 2 * gap) * 0.32;
            double wIso  = (usable - 2 * gap) * 0.32;

            double footH = 44;
            double footTop = ph - M - footH;
            double h = footTop - top - 8;

            if (fp.IsPlate)
            {
                DrawPlate(gfx, fp, new XRect(M, top, usable, h));
            }
            else
            {
                DrawFlatPattern(gfx, fp, new XRect(M, top, wFlat, h));
                DrawCrossSection(gfx, fp, new XRect(M + wFlat + gap, top, wSect, h));
                DrawIsometric(gfx, fp, new XRect(M + wFlat + wSect + 2 * gap, top, wIso, h));
            }
            DrawFootnote(gfx, fp, new XRect(M, footTop, usable, footH));

            // Copyright line under the details box (current year).
            var copyFont = new XFont("Arial", 7, XFontStyleEx.Regular);
            gfx.DrawString($"© Mithril Metals Corp, {System.DateTime.Now.Year}", copyFont, DimBrush,
                new XRect(M, footTop + footH + 2, usable, 10), XStringFormats.TopLeft);
        }

        using var ms = new MemoryStream();
        doc.Save(ms);
        return ms.ToArray();
    }

    // ── 1. Flat pattern (the cut blank) ──────────────────────────────────────
    private static void DrawFlatPattern(XGraphics gfx, FlatPatternResult fp, XRect box)
    {
        var labelFont = new XFont("Arial", 8, XFontStyleEx.Regular);
        var titleFont = new XFont("Arial", 9, XFontStyleEx.Bold);
        var cutPen  = new XPen(CutColor, 1.2);
        var bendPen = new XPen(BendColor, 0.8) { DashStyle = XDashStyle.Dash };

        gfx.DrawString("Flat pattern (cut)", titleFont, XBrushes.Black,
            new XRect(box.X, box.Y, box.Width, 12), XStringFormats.TopCenter);
        var area = new XRect(box.X, box.Y + 16, box.Width, box.Height - 16);

        double mw = fp.FlatWidth, mh = fp.FlatHeight;
        if (mw <= 0 || mh <= 0) return;

        const double dimBand = 40;
        double availW = area.Width - dimBand * 2, availH = area.Height - dimBand * 2;
        double scale = Math.Min(availW / mw, availH / mh);
        double drawW = mw * scale, drawH = mh * scale;
        double ox = area.X + dimBand + (availW - drawW) / 2;
        double oy = area.Y + dimBand + (availH - drawH) / 2;

        XPoint P(double mx, double my) => new(ox + mx * scale, oy + drawH - my * scale);

        gfx.DrawRectangle(cutPen, new XRect(ox, oy, drawW, drawH));
        foreach (var bx in fp.BendLinesX)
            gfx.DrawLine(bendPen, P(bx, 0), P(bx, mh));

        DimH(gfx, labelFont, P(0, 0).X, P(mw, 0).X, oy + drawH, oy + drawH + 22, F(mw), false);
        DimV(gfx, labelFont, ox, ox - 24, oy, oy + drawH, F(mh), false);

        double prev = 0;
        int idx = 0;
        foreach (var bx in fp.BendLinesX)
        {
            double dimY = oy - 16 - (idx % 2) * 14;   // stagger so labels don't collide on narrow blanks
            DimH(gfx, labelFont, P(prev, 0).X, P(bx, 0).X, oy, dimY, F(bx - prev), false);
            prev = bx;
            idx++;
        }
    }

    // ── 2. Dimensioned end-section (any shape, primary dims only) ─────────────
    private static void DrawCrossSection(XGraphics gfx, FlatPatternResult fp, XRect box)
    {
        var titleFont = new XFont("Arial", 9, XFontStyleEx.Bold);
        var dimFont   = new XFont("Arial", 8, XFontStyleEx.Bold);
        var matPen    = new XPen(CutColor, 1.1);

        gfx.DrawString("End section", titleFont, XBrushes.Black,
            new XRect(box.X, box.Y, box.Width, 12), XStringFormats.TopCenter);
        var area = new XRect(box.X, box.Y + 16, box.Width, box.Height - 16);

        var prof = fp.Profile;
        if (prof.Count < 3) return;
        double t = fp.Spec.Thickness, t2 = t / 2.0;
        double minX = prof.Min(p => p.x), maxX = prof.Max(p => p.x);
        double minY = prof.Min(p => p.y), maxY = prof.Max(p => p.y);
        double mw = maxX - minX, mh = maxY - minY;
        if (mw <= 0 || mh <= 0) return;

        double bandL = 56, bandB = 50, bandT = 14, bandR = 36;
        double availW = area.Width - bandL - bandR, availH = area.Height - bandB - bandT;
        double scale = Math.Min(availW / mw, availH / mh);
        double drawW = mw * scale, drawH = mh * scale;
        double ox = area.X + bandL + (availW - drawW) / 2;
        double oy = area.Y + bandT + (availH - drawH) / 2;

        XPoint P(double mx, double my) => new(ox + (mx - minX) * scale, oy + drawH - (my - minY) * scale);

        gfx.DrawPolygon(matPen, prof.Select(p => P(p.x, p.y)).ToArray());

        double sBottom = oy + drawH, sLeft = ox, sRight = ox + drawW;
        double webO = fp.WebOutside, flL = fp.FlangeLeftOutside, flR = fp.FlangeRightOutside;
        bool wIn = fp.Spec.Web.Basis == DimBasis.Inside;
        bool flIn = fp.Spec.FlangeLeft.Basis == DimBasis.Inside;
        bool frIn = fp.Spec.FlangeRight.Basis == DimBasis.Inside;

        void TLeader(double mx, double myA, double myB)
        {
            var a = P(mx, myA); var b = P(mx, myB);
            var lead = new XPen(DimColor, 0.6);
            gfx.DrawLine(lead, a, b);
            gfx.DrawLine(lead, b, new XPoint(b.X + 22, b.Y - 12));
            gfx.DrawString($"T {F(t)}", dimFont, TextBrush, new XRect(b.X + 24, b.Y - 18, 60, 11), XStringFormats.TopLeft);
        }

        switch (fp.Spec.Type)
        {
            case PartType.UChannel:
                // Web below (faces at y=0/outer or y=t/inner); left flange on the left.
                if (wIn) DimH(gfx, dimFont, P(t, 0).X, P(webO - t, 0).X, P(t, t).Y, sBottom + 22, $"{F(webO - 2 * t)} ID", true);
                else     DimH(gfx, dimFont, P(0, 0).X, P(webO, 0).X, P(0, 0).Y, sBottom + 22, $"{F(webO)} OD", true);
                if (flIn) DimV(gfx, dimFont, P(t, 0).X, sLeft - 24, P(0, t).Y, P(0, flL).Y, $"{F(flL - t)} ID", true);
                else      DimV(gfx, dimFont, P(0, 0).X, sLeft - 24, P(0, 0).Y, P(0, flL).Y, $"{F(flL)} OD", true);
                TLeader(webO * 0.5, 0, t);
                break;

            case PartType.LAngle:
                // legA (FlangeLeft) horizontal at the bottom; legB (FlangeRight) vertical at the left.
                // Walker frame: centreline legA along y=0 (faces ±t/2), legB along x=0.
                if (flIn) DimH(gfx, dimFont, P(t2, t2).X, P(flL, t2).X, P(0, t2).Y, sBottom + 22, $"{F(flL - t)} ID", true);
                else      DimH(gfx, dimFont, P(0, -t2).X, P(flL, -t2).X, P(0, -t2).Y, sBottom + 22, $"{F(flL)} OD", true);
                if (frIn) DimV(gfx, dimFont, P(t2, 0).X, sLeft - 24, P(t2, t2).Y, P(t2, flR).Y, $"{F(flR - t)} ID", true);
                else      DimV(gfx, dimFont, P(-t2, 0).X, sLeft - 24, P(-t2, 0).Y, P(-t2, flR).Y, $"{F(flR)} OD", true);
                TLeader(flL * 0.5, -t2, t2);
                break;

            case PartType.ZChannel:
                // Walker frame: web centreline y=0 (x 0..webO), left flange up x=0 (y 0..flL),
                // right flange down x=webO (y 0..-flR). Web dim sits just below the web (short witness).
                DimH(gfx, dimFont, P(0, -t2).X, P(webO, -t2).X, P(0, -t2).Y, P(0, -t2).Y + 24,
                     wIn ? $"{F(webO - 2 * t)} ID" : $"{F(webO)} OD", true);
                DimV(gfx, dimFont, P(-t2, 0).X, sLeft - 24, P(-t2, 0).Y, P(-t2, flL).Y,
                     flIn ? $"{F(flL - t)} ID" : $"{F(flL)} OD", true);
                DimV(gfx, dimFont, P(webO + t2, 0).X, sRight + 24, P(webO + t2, 0).Y, P(webO + t2, -flR).Y,
                     frIn ? $"{F(flR - t)} ID" : $"{F(flR)} OD", true);
                TLeader(webO * 0.5, t2, -t2);
                break;
        }
    }

    // ── 3. Thick-walled isometric (extrudes the profile) ─────────────────────
    private static void DrawIsometric(XGraphics gfx, FlatPatternResult fp, XRect box)
    {
        var titleFont = new XFont("Arial", 9, XFontStyleEx.Bold);
        var pen   = new XPen(CutColor, 0.9);
        var faint = new XPen(XColor.FromArgb(150, 150, 150), 0.6);
        var dimFont = new XFont("Arial", 8, XFontStyleEx.Bold);

        gfx.DrawString("Formed part", titleFont, XBrushes.Black,
            new XRect(box.X, box.Y, box.Width, 12), XStringFormats.TopCenter);
        var area = new XRect(box.X, box.Y + 16, box.Width, box.Height - 16);

        var prof = fp.Profile;
        double len = fp.FlatHeight;
        if (prof.Count < 3 || len <= 0) return;

        double c30 = Math.Cos(Math.PI / 6), s30 = Math.Sin(Math.PI / 6);
        (double u, double v) Iso(double x, double y, double z) => ((x - z) * c30, (x + z) * s30 + y);

        var front = prof.Select(p => Iso(p.x, p.y, 0)).ToArray();
        var back  = prof.Select(p => Iso(p.x, p.y, len)).ToArray();
        var all = front.Concat(back).ToList();

        double minU = all.Min(p => p.u), maxU = all.Max(p => p.u);
        double minV = all.Min(p => p.v), maxV = all.Max(p => p.v);
        double gw = maxU - minU, gh = maxV - minV;
        const double pad = 44;
        double scale = Math.Min((area.Width - pad * 2) / gw, (area.Height - pad * 2) / gh);
        double ox = area.X + (area.Width - gw * scale) / 2;
        double oy = area.Y + (area.Height - gh * scale) / 2;

        XPoint S((double u, double v) p) => new(ox + (p.u - minU) * scale, oy + (maxV - p.v) * scale);

        var frontS = front.Select(S).ToArray();
        var backS  = back.Select(S).ToArray();

        DrawLoop(gfx, faint, backS);
        int step = Math.Max(1, prof.Count / 10);          // sparse extrusion ridges (avoid clutter)
        for (int i = 0; i < prof.Count; i += step)
            gfx.DrawLine(faint, frontS[i], backS[i]);
        DrawLoop(gfx, pen, frontS);

        // Length dimension on the front-most outer edge, ~1/4" off the part.
        double bMinX = frontS.Concat(backS).Min(p => p.X), bMaxX = frontS.Concat(backS).Max(p => p.X);
        double bMinY = frontS.Concat(backS).Min(p => p.Y), bMaxY = frontS.Concat(backS).Max(p => p.Y);
        double cX = (bMinX + bMaxX) / 2, cY = (bMinY + bMaxY) / 2;

        // pick the rightmost profile point as the length-dim edge
        var rightMost = prof.OrderByDescending(p => p.x).ThenBy(p => p.y).First();
        XPoint A = S(Iso(rightMost.x, rightMost.y, 0)), B = S(Iso(rightMost.x, rightMost.y, len));
        double dx = B.X - A.X, dy = B.Y - A.Y, l = Math.Sqrt(dx * dx + dy * dy);
        if (l < 1e-6) return;
        double ux = dx / l, uy = dy / l, px = -uy, py = ux;
        double mX = (A.X + B.X) / 2, mY = (A.Y + B.Y) / 2;
        if (px * (mX - cX) + py * (mY - cY) < 0) { px = -px; py = -py; }
        const double off = 18;
        XPoint Ao = new(A.X + px * off, A.Y + py * off), Bo = new(B.X + px * off, B.Y + py * off);
        var dimPen = new XPen(DimColor, 0.6);
        Ext(gfx, dimPen, A, Ao); Ext(gfx, dimPen, B, Bo);
        gfx.DrawLine(dimPen, Ao, Bo);
        Arrow(gfx, Ao, -ux, -uy); Arrow(gfx, Bo, ux, uy);
        RotText(gfx, dimFont, $"L {F(len)}", new XPoint((Ao.X + Bo.X) / 2, (Ao.Y + Bo.Y) / 2), Math.Atan2(Bo.Y - Ao.Y, Bo.X - Ao.X));
    }

    // ── Flat plate: single dimensioned top view with bolt holes ──────────────
    private static void DrawPlate(XGraphics gfx, FlatPatternResult fp, XRect box)
    {
        var titleFont = new XFont("Arial", 9, XFontStyleEx.Bold);
        var dimFont = new XFont("Arial", 8, XFontStyleEx.Bold);
        var cutPen = new XPen(CutColor, 1.2);
        var thin = new XPen(DimColor, 0.5);

        gfx.DrawString("Plate — top view", titleFont, XBrushes.Black,
            new XRect(box.X, box.Y, box.Width, 12), XStringFormats.TopCenter);
        var area = new XRect(box.X, box.Y + 18, box.Width, box.Height - 18);

        double L = fp.FlatWidth, W = fp.FlatHeight;
        if (L <= 0 || W <= 0) return;
        const double band = 50;
        double availW = area.Width - band * 2, availH = area.Height - band * 2;
        double scale = Math.Min(availW / L, availH / W);
        double drawW = L * scale, drawH = W * scale;
        double ox = area.X + (area.Width - drawW) / 2;
        double oy = area.Y + (area.Height - drawH) / 2;
        XPoint P(double mx, double my) => new(ox + mx * scale, oy + drawH - my * scale);

        gfx.DrawRectangle(cutPen, new XRect(ox, oy, drawW, drawH));
        foreach (var (hx, hy, dia) in fp.Holes)
        {
            double r = dia / 2.0 * scale;
            var c = P(hx, hy);
            gfx.DrawEllipse(cutPen, c.X - r, c.Y - r, 2 * r, 2 * r);
            gfx.DrawLine(thin, c.X - r - 2, c.Y, c.X + r + 2, c.Y);   // centre cross
            gfx.DrawLine(thin, c.X, c.Y - r - 2, c.X, c.Y + r + 2);
        }

        DimH(gfx, dimFont, P(0, 0).X, P(L, 0).X, oy + drawH, oy + drawH + 22, F(L), true);
        DimV(gfx, dimFont, ox, ox - 26, oy, oy + drawH, F(W), true);

        if (fp.Holes.Count > 0)
        {
            var (hx, hy, dia) = fp.Holes[0];
            var c = P(hx, hy);
            double r = dia / 2.0 * scale;
            var tip = new XPoint(c.X + r * 0.7, c.Y - r * 0.7);
            var lp = new XPoint(c.X + 22, c.Y - 20);
            gfx.DrawLine(new XPen(DimColor, 0.6), tip, lp);
            gfx.DrawString($"{fp.Holes.Count} x {F(dia)} dia", dimFont, TextBrush,
                new XRect(lp.X + 2, lp.Y - 5, 90, 11), XStringFormats.TopLeft);
        }

        if (fp.Spec.Holes is { } hs)
        {
            if (hs.Pattern != HolePattern.Corner)
            {
                var xs = fp.Holes.Select(h => h.x).Distinct().OrderBy(x => x).ToList();
                if (xs.Count >= 2)
                    DimH(gfx, dimFont, P(xs[0], 0).X, P(xs[1], 0).X, oy, oy - 16, F(xs[1] - xs[0]), false);
            }
            else if (fp.Holes.Count > 0)
            {
                double hx = fp.Holes[0].x;
                DimH(gfx, dimFont, P(0, 0).X, P(hx, 0).X, oy, oy - 16, F(hx), false);
            }
        }
    }

    // ── Footnote box ─────────────────────────────────────────────────────────
    private static void DrawFootnote(XGraphics gfx, FlatPatternResult fp, XRect box)
    {
        gfx.DrawRectangle(new XPen(XColor.FromArgb(205, 205, 205), 0.8), box);
        var font = new XFont("Arial", 8, XFontStyleEx.Regular);
        double y = box.Y + 3;
        foreach (var line in fp.Summary.Split('\n'))
        {
            gfx.DrawString(line, font, XBrushes.Black, new XRect(box.X + 8, y, box.Width - 16, 10), XStringFormats.TopLeft);
            y += 10;
        }
        gfx.DrawString(
            $"solid = cut  |  dashed = bend up  |  bold = as specified  |  inside radius Ri {F(fp.Spec.InsideRadius)}\"  |  all dimensions in decimal inches",
            font, DimBrush, new XRect(box.X + 8, y, box.Width - 16, 10), XStringFormats.TopLeft);
    }

    // ── helpers ──────────────────────────────────────────────────────────────
    private static void DrawLoop(XGraphics gfx, XPen pen, XPoint[] loop)
    {
        for (int i = 0; i < loop.Length; i++)
            gfx.DrawLine(pen, loop[i], loop[(i + 1) % loop.Length]);
    }

    private static void Ext(XGraphics gfx, XPen pen, XPoint from, XPoint to)
    {
        double dx = to.X - from.X, dy = to.Y - from.Y, l = Math.Sqrt(dx * dx + dy * dy);
        if (l < 1e-6) return;
        double ux = dx / l, uy = dy / l;
        gfx.DrawLine(pen, new XPoint(from.X + ux * ExtGap, from.Y + uy * ExtGap),
                          new XPoint(to.X + ux * ExtOver, to.Y + uy * ExtOver));
    }

    private static void RotText(XGraphics gfx, XFont font, string s, XPoint mid, double angRad)
    {
        double deg = angRad * 180.0 / Math.PI;
        if (deg > 90) deg -= 180; else if (deg <= -90) deg += 180;
        var st = gfx.Save();
        gfx.TranslateTransform(mid.X, mid.Y);
        gfx.RotateTransform(deg);
        gfx.DrawString(s, font, TextBrush, new XPoint(0, -3), XStringFormats.BottomCenter);
        gfx.Restore(st);
    }

    private static void DimH(XGraphics gfx, XFont font, double x1, double x2,
        double faceY, double dimY, string label, bool bold)
    {
        var pen = new XPen(DimColor, 0.6);
        double dir = Math.Sign(dimY - faceY); if (dir == 0) dir = 1;
        gfx.DrawLine(pen, x1, faceY + ExtGap * dir, x1, dimY + ExtOver * dir);
        gfx.DrawLine(pen, x2, faceY + ExtGap * dir, x2, dimY + ExtOver * dir);
        gfx.DrawLine(pen, x1, dimY, x2, dimY);
        Arrow(gfx, new XPoint(x1, dimY), -1, 0);
        Arrow(gfx, new XPoint(x2, dimY), 1, 0);
        double ty = dir > 0 ? dimY + 2 : dimY - 12;
        gfx.DrawString(label, font, TextBrush, new XRect((x1 + x2) / 2 - 40, ty, 80, 11), XStringFormats.TopCenter);
    }

    private static void DimV(XGraphics gfx, XFont font, double faceX, double dimX,
        double y1, double y2, string label, bool bold)
    {
        var pen = new XPen(DimColor, 0.6);
        double dir = Math.Sign(dimX - faceX); if (dir == 0) dir = -1;
        gfx.DrawLine(pen, faceX + ExtGap * dir, y1, dimX + ExtOver * dir, y1);
        gfx.DrawLine(pen, faceX + ExtGap * dir, y2, dimX + ExtOver * dir, y2);
        gfx.DrawLine(pen, dimX, y1, dimX, y2);
        Arrow(gfx, new XPoint(dimX, y1), 0, -1);
        Arrow(gfx, new XPoint(dimX, y2), 0, 1);
        // place label on the far side of the dim line from the part
        bool right = dir > 0;
        double lx = right ? dimX + 6 : dimX - 54;
        gfx.DrawString(label, font, TextBrush, new XRect(lx, (y1 + y2) / 2 - 6, 48, 11),
            right ? XStringFormats.TopLeft : XStringFormats.TopRight);
    }

    private static void Arrow(XGraphics gfx, XPoint tip, double dx, double dy)
    {
        const double len = 5, half = 1.8;
        double px = -dy, py = dx;
        var b1 = new XPoint(tip.X - dx * len + px * half, tip.Y - dy * len + py * half);
        var b2 = new XPoint(tip.X - dx * len - px * half, tip.Y - dy * len - py * half);
        gfx.DrawPolygon(DimBrush, new[] { tip, b1, b2 }, XFillMode.Winding);
    }

    private static string F(double v) => v.ToString("0.000", CultureInfo.InvariantCulture);
}
