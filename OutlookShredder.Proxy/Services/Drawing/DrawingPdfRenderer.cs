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

    // Section-cut plane styling. Two distinct styles so multiple cut planes read apart:
    // A = green / dash (pan "side" + single-section views), B = orange / dash-dot (pan "end").
    private static readonly XColor SecColorA = XColor.FromArgb(0, 150, 70);
    private static readonly XColor SecColorB = XColor.FromArgb(202, 86, 0);

    private const double ExtGap = 2.0;
    private const double ExtOver = 3.0;

    // Section-view panel margins (shared so the pan's two sections render at one common scale).
    private const double SecBandL = 46, SecBandB = 42, SecBandT = 12, SecBandR = 18;

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
            // Plates carry the product-standard order (thickness x width x length); formed parts
            // keep the 2D blank extent (cut width x cut length).
            string blankLabel = fp.IsColumn
                ? $"Cut to:  pipe/tube {F(fp.ColumnTubeLength)}\"  +  base & bearing plate flat patterns (one DXF)"
                : fp.IsPlate
                ? $"Flat blank:  {F(fp.Spec.Thickness)}\" x {F(fp.FlatHeight)}\" x {F(fp.FlatWidth)}\""
                : fp.IsPaddle
                    ? $"Plate to cut:  {F(fp.Spec.Thickness)}\" x {F(fp.FlatHeight)}\" x {F(fp.FlatWidth)}\"  (rectangular blank, incl. handle, for quoting)"
                    : $"Flat blank:  {F(fp.FlatWidth)}\" x {F(fp.FlatHeight)}\"";
            gfx.DrawString(blankLabel,
                blankFont, XBrushes.Black, new XRect(M, M + 16, pw - 2 * M, 14), XStringFormats.TopLeft);

            double top = M + 38;
            if (fp.IsPlate && fp.Spec.Holes is { } h0 && fp.Holes.Count > 0)
            {
                gfx.DrawString($"{fp.Holes.Count} holes,  {F(h0.Diameter)}\" dia",
                    new XFont("Arial", 10, XFontStyleEx.Bold), XBrushes.Black,
                    new XRect(M, M + 31, pw - 2 * M, 12), XStringFormats.TopLeft);
                top = M + 48;
            }
            double usable = pw - 2 * M;
            const double gap = 16;
            // L / U / Z three-panel split: flat 1/5, section 2/5, iso 2/5 (the iso gets the most room).
            double wFlat = (usable - 2 * gap) * 0.2;
            double wSect = (usable - 2 * gap) * 0.4;
            double wIso  = (usable - 2 * gap) * 0.4;

            // Box grows with the summary (+1 line for the "solid = cut …" legend).
            int footLines = fp.Summary.Split('\n').Length + 1;
            double footH = footLines * 10 + 6;
            double footTop = ph - M - footH;
            double h = footTop - top - 8;

            if (fp.IsPan)
            {
                // Top row: flat pattern + formed iso.  Bottom row: side section + end section.
                double topH = h * 0.56, botH = h - topH - 10;
                double secY = top + topH + 10;
                // Iso gets 2/3 of the top row, flat pattern 1/3 (the iso is the busy view).
                double wFlatP = (usable - gap) / 3.0, wIsoP = (usable - gap) * 2.0 / 3.0;
                double isoX = M + wFlatP + gap;
                DrawPan(gfx, fp, new XRect(M, top, wFlatP, topH));
                DrawPanIso(gfx, fp, new XRect(isoX, top, wIsoP, topH - 46));
                DrawSectionKey(gfx, new XRect(isoX + (wIsoP - 176) / 2, top + topH - 36, 176, 30), isPan: true);

                // Both sections share ONE scale so an equal dimension reads the same size in each.
                double wSec = (usable - gap) / 2;
                var sideBox = new XRect(M, secY, wSec, botH);
                var endBox  = new XRect(M + wSec + gap, secY, wSec, botH);
                double secScale = Math.Min(SectionFitScale(sideBox, fp.PanSideProfile),
                                           SectionFitScale(endBox, fp.PanEndProfile));
                DrawSection(gfx, sideBox, "Side section", SecColorA, XDashStyle.Dash,
                    fp.PanSideProfile, fp.PanBaseY1 - fp.PanBaseY0, fp.PanDepth, fp.Spec.Thickness, secScale);
                DrawSection(gfx, endBox, "End section", SecColorB, XDashStyle.DashDot,
                    fp.PanEndProfile, fp.PanBaseX1 - fp.PanBaseX0, fp.PanDepth, fp.Spec.Thickness, secScale);
            }
            else if (fp.IsColumn)
            {
                DrawColumn(gfx, fp, new XRect(M, top, usable, h));
            }
            else if (fp.IsPlate)
            {
                DrawPlate(gfx, fp, new XRect(M, top, usable, h));
            }
            else if (fp.IsPaddle)
            {
                DrawPaddleBlind(gfx, fp, new XRect(M, top, usable, h));
            }
            else
            {
                DrawFlatPattern(gfx, fp, new XRect(M, top, wFlat, h));
                DrawCrossSection(gfx, fp, new XRect(M + wFlat + gap, top, wSect, h));
                double isoX = M + wFlat + wSect + 2 * gap;
                DrawIsometric(gfx, fp, new XRect(isoX, top, wIso, h - 46));
                DrawSectionKey(gfx, new XRect(isoX + (wIso - 120) / 2, top + h - 36, 120, 30), isPan: false);
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

        DimH(gfx, labelFont, P(0, 0).X, P(mw, 0).X, oy + drawH, oy + drawH + 22, Fq(mw), false);
        DimV(gfx, labelFont, ox, ox - 24, oy, oy + drawH, Fq(mh), false);

        double prev = 0;
        int idx = 0;
        foreach (var bx in fp.BendLinesX)
        {
            double dimY = oy - 16 - (idx % 2) * 14;   // stagger so labels don't collide on narrow blanks
            DimH(gfx, labelFont, P(prev, 0).X, P(bx, 0).X, oy, dimY, Fq(bx - prev), false);
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

        var tsz = gfx.MeasureString("End section", titleFont);
        double tStart = box.X + (box.Width - tsz.Width - 34) / 2;
        gfx.DrawString("End section", titleFont, XBrushes.Black, new XRect(tStart, box.Y, tsz.Width + 2, 12), XStringFormats.TopLeft);
        gfx.DrawLine(new XPen(SecColorA, 1.4) { DashStyle = XDashStyle.DashDot },
            new XPoint(tStart + tsz.Width + 8, box.Y + 7), new XPoint(tStart + tsz.Width + 34, box.Y + 7));
        var area = new XRect(box.X, box.Y + 16, box.Width, box.Height - 16);

        var prof = fp.Profile;
        if (prof.Count < 3) return;
        double t = fp.Spec.Thickness, t2 = t / 2.0;
        double minX = prof.Min(p => p.x), maxX = prof.Max(p => p.x);
        double minY = prof.Min(p => p.y), maxY = prof.Max(p => p.y);
        double mw = maxX - minX, mh = maxY - minY;
        if (mw <= 0 || mh <= 0) return;

        double bandL = 58, bandB = 50, bandT = 14, bandR = 66;   // room for a right-side dim (Z)
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
            var lp = new XPoint(b.X, Math.Max(area.Y + 9, b.Y - 16));   // up into the open section — never overflows
            gfx.DrawLine(lead, a, b);
            gfx.DrawLine(lead, b, lp);
            gfx.DrawString($"T {Fq(t)}", dimFont, TextBrush, new XRect(lp.X - 28, lp.Y - 9, 56, 11), XStringFormats.TopCenter);
        }

        switch (fp.Spec.Type)
        {
            case PartType.UChannel:
                // Web below (faces at y=0/outer or y=t/inner); left flange on the left.
                if (wIn) DimH(gfx, dimFont, P(t, 0).X, P(webO - t, 0).X, P(t, t).Y, sBottom + 22, $"{Fq(webO - 2 * t)} ID", true);
                else     DimH(gfx, dimFont, P(0, 0).X, P(webO, 0).X, P(0, 0).Y, sBottom + 22, $"{Fq(webO)} OD", true);
                if (flIn) DimV(gfx, dimFont, P(t, 0).X, sLeft - 24, P(0, t).Y, P(0, flL).Y, $"{Fq(flL - t)} ID", true);
                else      DimV(gfx, dimFont, P(0, 0).X, sLeft - 24, P(0, 0).Y, P(0, flL).Y, $"{Fq(flL)} OD", true);
                TLeader(webO * 0.5, 0, t);
                break;

            case PartType.LAngle:
                // legA (FlangeLeft) horizontal at the bottom; legB (FlangeRight) vertical at the left.
                // Walker frame: centreline legA along y=0 (faces ±t/2), legB along x=0.
                if (flIn) DimH(gfx, dimFont, P(t2, t2).X, P(flL, t2).X, P(0, t2).Y, sBottom + 22, $"{Fq(flL - t)} ID", true);
                else      DimH(gfx, dimFont, P(0, -t2).X, P(flL, -t2).X, P(0, -t2).Y, sBottom + 22, $"{Fq(flL)} OD", true);
                if (frIn) DimV(gfx, dimFont, P(t2, 0).X, sLeft - 24, P(t2, t2).Y, P(t2, flR).Y, $"{Fq(flR - t)} ID", true);
                else      DimV(gfx, dimFont, P(-t2, 0).X, sLeft - 24, P(-t2, 0).Y, P(-t2, flR).Y, $"{Fq(flR)} OD", true);
                TLeader(flL * 0.5, -t2, t2);
                break;

            case PartType.ZChannel:
                // Walker frame: web centreline y=0 (x 0..webO), left flange up x=0 (y 0..flL),
                // right flange down x=webO (y 0..-flR). Web dim below the whole part.
                DimH(gfx, dimFont, P(0, -t2).X, P(webO, -t2).X, P(0, -t2).Y, sBottom + 22,
                     wIn ? $"{Fq(webO - 2 * t)} ID" : $"{Fq(webO)} OD", true);
                DimV(gfx, dimFont, P(-t2, 0).X, sLeft - 24, P(-t2, 0).Y, P(-t2, flL).Y,
                     flIn ? $"{Fq(flL - t)} ID" : $"{Fq(flL)} OD", true);
                DimV(gfx, dimFont, P(webO + t2, 0).X, sRight + 24, P(webO + t2, 0).Y, P(webO + t2, -flR).Y,
                     frIn ? $"{Fq(flR - t)} ID" : $"{Fq(flR)} OD", true);
                TLeader(webO * 0.5, t2, -t2);
                break;
        }

        // ── "Finish" callout: boxed, highlighted label + leader to the finished face ──
        var finishFont = new XFont("Arial", 10, XFontStyleEx.Bold);
        void FinishCallout(double mx, double my, double sdx, double sdy)
        {
            var tip = P(mx, my);
            double l = Math.Sqrt(sdx * sdx + sdy * sdy); if (l < 1e-6) l = 1;
            double ux = sdx / l, uy = sdy / l;

            // Boxed label, offset clear of the part; clamp inside the panel so it never clips the border.
            var sz = gfx.MeasureString("Finish", finishFont);
            double bw = sz.Width + 9, bh = sz.Height + 5;
            var anchor = new XPoint(tip.X + ux * 34, tip.Y + uy * 34);
            double bx = Math.Max(area.X + 1, Math.Min(anchor.X - bw / 2, area.X + area.Width - bw - 1));
            double by = Math.Max(area.Y + 1, Math.Min(anchor.Y - bh / 2, area.Y + area.Height - bh - 1));
            var br = new XRect(bx, by, bw, bh);

            // Leader from the box toward the surface; arrowhead on the surface (box drawn over the tail).
            var bc = new XPoint(br.X + bw / 2, br.Y + bh / 2);
            gfx.DrawLine(new XPen(CutColor, 1.0), bc, tip);
            double ax = tip.X - bc.X, ay = tip.Y - bc.Y, al = Math.Sqrt(ax * ax + ay * ay); if (al < 1e-6) al = 1;
            double dx = ax / al, dy = ay / al, px = -dy, py = dx;
            var a1 = new XPoint(tip.X - dx * 5 + px * 1.9, tip.Y - dy * 5 + py * 1.9);
            var a2 = new XPoint(tip.X - dx * 5 - px * 1.9, tip.Y - dy * 5 - py * 1.9);
            gfx.DrawPolygon(XBrushes.Black, new[] { tip, a1, a2 }, XFillMode.Winding);

            gfx.DrawRectangle(XBrushes.White, br);
            gfx.DrawRectangle(new XPen(XColor.FromArgb(110, 110, 110), 0.9), br);
            gfx.DrawString("Finish", finishFont, XBrushes.Black, br, XStringFormats.Center);
        }

        switch (fp.Spec.Finish)
        {
            case FinishSide.Outside when fp.Spec.Type == PartType.UChannel:
                FinishCallout(webO, flR * 0.5, 1, 0); break;            // right flange outer face
            case FinishSide.Inside when fp.Spec.Type == PartType.UChannel:
                FinishCallout(webO - t, flR * 0.5, -1, 0); break;       // right flange inner (cavity)
            case FinishSide.Outside when fp.Spec.Type == PartType.LAngle:
                FinishCallout(flL, -t2, 0.7, 0.7); break;               // convex (lower-right) corner
            case FinishSide.Inside when fp.Spec.Type == PartType.LAngle:
                FinishCallout(flL * 0.55, t2, 0, -1); break;            // concave interior angle
            case FinishSide.Top:                                        // Z: first flange's top face
                FinishCallout(webO * 0.62, t2, 0, -1); break;
            case FinishSide.Bottom:                                     // Z: first flange's bottom face
                FinishCallout(webO * 0.38, -t2, 0, 1); break;
        }

        // ── Per-bend callouts: arc at the bend + label carried outside the part. Flange bends show
        //    the actual included angle (angle mode only); returns are always labelled (90° / HEM). ──
        if (fp.SectionBends.Count > 0)
        {
            var angFont = new XFont("Arial", 8, XFontStyleEx.Bold);
            // Part centroid (page space) — the label is pushed outward, away from this, so it clears the part.
            double cxs = 0, cys = 0;
            foreach (var p in prof) { var pp = P(p.x, p.y); cxs += pp.X; cys += pp.Y; }
            var centroid = new XPoint(cxs / prof.Count, cys / prof.Count);

            foreach (var b in fp.SectionBends)
            {
                // Flange bends are only called out in angle mode; returns always (they're a real feature).
                if (!b.IsReturn && !fp.Spec.AnglesAnnotated) continue;

                var apex = P(b.X, b.Y);

                // Face rays from the apex (page space — y points down, so negate model hy).
                double inx = b.InHx, iny = -b.InHy, outx = b.OutHx, outy = -b.OutHy;
                double il = Math.Sqrt(inx * inx + iny * iny); if (il < 1e-6) il = 1; inx /= il; iny /= il;
                double ol = Math.Sqrt(outx * outx + outy * outy); if (ol < 1e-6) ol = 1; outx /= ol; outy /= ol;
                double a1 = Math.Atan2(-iny, -inx), a2 = Math.Atan2(outy, outx);
                double da = a2 - a1; while (da <= -Math.PI) da += 2 * Math.PI; while (da > Math.PI) da -= 2 * Math.PI;

                // Callout text: the ACTUAL angle between the faces (what the arc spans). A 180° hem's
                // faces are parallel (≈0° between them) — label it "HEM" and skip the degenerate arc.
                double shownDeg = Math.Abs(da) * 180.0 / Math.PI;
                bool hem = b.IsReturn && b.AngleDeg >= 170;
                var txt = hem ? "HEM" : (b.IsReturn ? $"{shownDeg:0.#}° ret" : $"{shownDeg:0.#}°");

                // Small arc spanning the two faces at the bend (skip for a hem).
                double r = 12;
                var arcPen = new XPen(DimColor, 0.8);
                if (!hem)
                {
                    XPoint Prev = new(apex.X + r * Math.Cos(a1), apex.Y + r * Math.Sin(a1));
                    for (int i = 1; i <= 10; i++)
                    {
                        double a = a1 + da * i / 10.0;
                        var cur = new XPoint(apex.X + r * Math.Cos(a), apex.Y + r * Math.Sin(a));
                        gfx.DrawLine(arcPen, Prev, cur); Prev = cur;
                    }
                }

                // Label pushed outward (apex away from centroid), clamped inside the panel; short leader.
                double odx = apex.X - centroid.X, ody = apex.Y - centroid.Y;
                double ol2 = Math.Sqrt(odx * odx + ody * ody); if (ol2 < 1e-6) { odx = 0; ody = -1; ol2 = 1; }
                odx /= ol2; ody /= ol2;
                var sz = gfx.MeasureString(txt, angFont);
                double lx = apex.X + odx * (r + 14), ly = apex.Y + ody * (r + 14);
                double bx = Math.Max(area.X + 1, Math.Min(lx - sz.Width / 2, area.X + area.Width - sz.Width - 1));
                double by = Math.Max(area.Y + 1, Math.Min(ly - sz.Height / 2, area.Y + area.Height - sz.Height - 1));
                gfx.DrawLine(arcPen, new XPoint(apex.X + odx * r, apex.Y + ody * r), new XPoint(bx + sz.Width / 2, by + sz.Height / 2));
                gfx.DrawString(txt, angFont, TextBrush, new XRect(bx, by, sz.Width + 2, sz.Height + 1), XStringFormats.TopLeft);
            }
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

        // Section-cut plane (End section, dash-dot) on the formed part; keyed under the drawing.
        var secPen = new XPen(SecColorA, 1.3) { DashStyle = XDashStyle.DashDot };
        var midS = prof.Select(p => S(Iso(p.x, p.y, len / 2))).ToArray();
        DrawLoop(gfx, secPen, midS);

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
        const double off = 30;   // keep the length dim clear of the part (~1/4")
        XPoint Ao = new(A.X + px * off, A.Y + py * off), Bo = new(B.X + px * off, B.Y + py * off);
        var dimPen = new XPen(DimColor, 0.6);
        Ext(gfx, dimPen, A, Ao); Ext(gfx, dimPen, B, Bo);
        gfx.DrawLine(dimPen, Ao, Bo);
        Arrow(gfx, Ao, -ux, -uy); Arrow(gfx, Bo, ux, uy);
        RotText(gfx, dimFont, $"L {Fq(len)}", new XPoint((Ao.X + Bo.X) / 2, (Ao.Y + Bo.Y) / 2), Math.Atan2(Bo.Y - Ao.Y, Bo.X - Ao.X));
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

        DimH(gfx, dimFont, P(0, 0).X, P(L, 0).X, oy + drawH, oy + drawH + 22, Fq(L), true);
        DimV(gfx, dimFont, ox, ox - 46, oy, oy + drawH, Fq(W), true);

        // Flitch: edge-to-row distances (top edge -> top row, bottom edge -> bottom row).
        if (fp.Spec.Holes is { } hsr && hsr.Pattern != HolePattern.Corner && fp.Holes.Count > 0)
        {
            double topEdge = hsr.TopEdge > 0 ? hsr.TopEdge : W * 0.25;
            double botEdge = hsr.BottomEdge > 0 ? hsr.BottomEdge : W * 0.25;
            double rowTop = W - topEdge, rowBot = botEdge;
            DimV(gfx, dimFont, ox, ox - 22, P(0, W).Y, P(0, rowTop).Y, Fq(topEdge), false);
            DimV(gfx, dimFont, ox, ox - 22, P(0, 0).Y, P(0, rowBot).Y, Fq(botEdge), false);
        }

        if (fp.Holes.Count > 0)
        {
            var (hx, hy, dia) = fp.Holes[0];
            var c = P(hx, hy);
            double r = dia / 2.0 * scale;
            string label = $"{fp.Holes.Count} x {Fq(dia)} dia";
            var sz = gfx.MeasureString(label, dimFont);
            double bw = sz.Width + 9, bh = sz.Height + 5;
            // White boxed callout in the clear band above the spacing chain — clear of the part edge.
            double bx = Math.Max(area.X + 1, Math.Min(c.X - bw / 2, area.X + area.Width - bw - 1));
            double by = Math.Max(area.Y + 1, oy - 38 - bh);
            var br = new XRect(bx, by, bw, bh);
            double leadX = Math.Max(br.X + 4, Math.Min(c.X, br.X + bw - 4));
            gfx.DrawLine(new XPen(DimColor, 0.6), new XPoint(leadX, br.Y + bh), new XPoint(c.X, c.Y - r - 2));
            gfx.DrawRectangle(XBrushes.White, br);
            gfx.DrawRectangle(new XPen(XColor.FromArgb(110, 110, 110), 0.9), br);
            gfx.DrawString(label, dimFont, TextBrush, br, XStringFormats.Center);
        }

        if (fp.Spec.Holes is { } hs)
        {
            if (hs.Pattern != HolePattern.Corner)
            {
                // Dimension chain across the top: LHS -> first hole, each spacing, last hole -> RHS.
                var xs = fp.Holes.Select(h => h.x).Distinct().OrderBy(x => x).ToList();
                if (xs.Count > 0)
                {
                    var chain = new List<double> { 0 };
                    chain.AddRange(xs);
                    chain.Add(L);
                    for (int i = 0; i < chain.Count - 1; i++)
                        DimH(gfx, dimFont, P(chain[i], 0).X, P(chain[i + 1], 0).X, oy, oy - 16,
                             Fq(chain[i + 1] - chain[i]), false);
                }
            }
            else if (fp.Holes.Count > 0)
            {
                // Base plate (corner holes): show the offset from BOTH the vertical (left) edge and
                // the horizontal (bottom) edge to the first corner hole, so neither is left implicit.
                double hx = fp.Holes[0].x, hy = fp.Holes[0].y;
                DimH(gfx, dimFont, P(0, 0).X, P(hx, 0).X, oy, oy - 16, Fq(hx), false);
                DimV(gfx, dimFont, ox, ox - 22, P(0, 0).Y, P(0, hy).Y, Fq(hy), false);
            }
        }
    }

    // ── Structural column: two plate flat-patterns (left) + a dimensioned side elevation (right) ──
    private static void DrawColumn(XGraphics gfx, FlatPatternResult fp, XRect box)
    {
        const double gap = 16;
        double elevW = box.Width * 0.34;
        double platesW = box.Width - elevW - gap;
        double pw = (platesW - gap) / 2;

        DrawPlatePanel(gfx, new XRect(box.X, box.Y, pw, box.Height),
            "Base plate", fp.ColumnBaseL, fp.ColumnBaseW, fp.ColumnBaseHoles,
            fp.ColumnShape, fp.ColumnOuterWidth, fp.ColumnOuterDepth, fp.ColumnWall, fp.ColumnTubeCornerR);
        DrawPlatePanel(gfx, new XRect(box.X + pw + gap, box.Y, pw, box.Height),
            "Bearing plate", fp.ColumnBearingL, fp.ColumnBearingW, fp.ColumnBearingHoles,
            fp.ColumnShape, fp.ColumnOuterWidth, fp.ColumnOuterDepth, fp.ColumnWall, fp.ColumnTubeCornerR);
        DrawColumnElevation(gfx, fp, new XRect(box.X + platesW + gap, box.Y, elevW, box.Height));
    }

    // ── One plate top view (rectangle + bolt holes + L/W + corner-hole offsets + tube weld outline) ──
    private static void DrawPlatePanel(XGraphics gfx, XRect box, string title, double L, double W,
        List<(double x, double y, double dia)> holes,
        string tubeShape, double tubeW, double tubeD, double tubeWall, double tubeCornerR)
    {
        var titleFont = new XFont("Arial", 9, XFontStyleEx.Bold);
        var dimFont = new XFont("Arial", 8, XFontStyleEx.Bold);
        var cutPen = new XPen(CutColor, 1.2);
        var thin = new XPen(DimColor, 0.5);

        gfx.DrawString($"{title} — top view", titleFont, XBrushes.Black,
            new XRect(box.X, box.Y, box.Width, 12), XStringFormats.TopCenter);
        var area = new XRect(box.X, box.Y + 18, box.Width, box.Height - 18);
        if (L <= 0 || W <= 0) return;

        const double band = 46;
        double availW = area.Width - band * 2, availH = area.Height - band * 2;
        double scale = Math.Min(availW / L, availH / W);
        double drawW = L * scale, drawH = W * scale;
        double ox = area.X + (area.Width - drawW) / 2;
        double oy = area.Y + (area.Height - drawH) / 2;
        XPoint P(double mx, double my) => new(ox + mx * scale, oy + drawH - my * scale);

        gfx.DrawRectangle(cutPen, new XRect(ox, oy, drawW, drawH));
        foreach (var (hx, hy, dia) in holes)
        {
            double r = dia / 2.0 * scale;
            var c = P(hx, hy);
            gfx.DrawEllipse(cutPen, c.X - r, c.Y - r, 2 * r, 2 * r);
            gfx.DrawLine(thin, c.X - r - 2, c.Y, c.X + r + 2, c.Y);
            gfx.DrawLine(thin, c.X, c.Y - r - 2, c.X, c.Y + r + 2);
        }

        DimH(gfx, dimFont, P(0, 0).X, P(L, 0).X, oy + drawH, oy + drawH + 22, Fq(L), true);
        DimV(gfx, dimFont, ox, ox - 40, oy, oy + drawH, Fq(W), true);

        if (holes.Count > 0)
        {
            var first = holes.OrderBy(h => h.x).ThenBy(h => h.y).First();
            DimH(gfx, dimFont, P(0, 0).X, P(first.x, 0).X, oy, oy - 16, Fq(first.x), false);
            DimV(gfx, dimFont, ox, ox - 20, P(0, 0).Y, P(0, first.y).Y, Fq(first.y), false);
            gfx.DrawString($"{holes.Count} x {Fq(holes[0].dia)} dia", dimFont, TextBrush,
                new XRect(area.X, oy - 32, area.Width, 11), XStringFormats.TopCenter);
        }

        // Tube weld-locating outline, centred (matches the etched marking layer in the DXF).
        var weldPen = new XPen(SecColorA, 1.1);   // green = marking / weld locator
        double mcx = L / 2.0, mcy = W / 2.0;
        if (tubeShape == "round" && tubeW > 0)
        {
            void Circ(double r)
            {
                if (r <= 0) return;
                var c = P(mcx, mcy); double rs = r * scale;
                gfx.DrawEllipse(weldPen, c.X - rs, c.Y - rs, 2 * rs, 2 * rs);
            }
            Circ(tubeW / 2.0);
            Circ(tubeW / 2.0 - tubeWall);
        }
        else if (tubeW > 0 && tubeD > 0)
        {
            void RectO(double hw, double hh, double rr)
            {
                if (hw <= 0 || hh <= 0) return;
                var a = P(mcx - hw, mcy - hh); var b = P(mcx + hw, mcy + hh);
                double x = Math.Min(a.X, b.X), y = Math.Min(a.Y, b.Y), wpx = Math.Abs(b.X - a.X), hpx = Math.Abs(b.Y - a.Y);
                if (rr > 0.001)
                {
                    double e = Math.Min(2 * rr * scale, Math.Min(wpx, hpx));
                    gfx.DrawRoundedRectangle(weldPen, x, y, wpx, hpx, e, e);
                }
                else gfx.DrawRectangle(weldPen, x, y, wpx, hpx);
            }
            RectO(tubeW / 2.0, tubeD / 2.0, tubeCornerR);                          // outer — filleted (non-alum)
            RectO(tubeW / 2.0 - tubeWall, tubeD / 2.0 - tubeWall, tubeCornerR);     // inner bore — filleted too
        }
        var wc = P(mcx, mcy);
        gfx.DrawString("tube", new XFont("Arial", 7, XFontStyleEx.Regular), new XSolidBrush(SecColorA),
            new XRect(wc.X - 20, wc.Y - 4, 40, 9), XStringFormats.TopCenter);
    }

    // ── Column side elevation: base / tube / bearing stacked, welded centred, with callouts ──
    private static void DrawColumnElevation(XGraphics gfx, FlatPatternResult fp, XRect box)
    {
        var titleFont = new XFont("Arial", 9, XFontStyleEx.Bold);
        var dimFont = new XFont("Arial", 8, XFontStyleEx.Bold);
        var tagFont = new XFont("Arial", 7, XFontStyleEx.Regular);
        var cutPen = new XPen(CutColor, 1.1);

        gfx.DrawString("Column elevation", titleFont, XBrushes.Black,
            new XRect(box.X, box.Y, box.Width, 12), XStringFormats.TopCenter);
        var area = new XRect(box.X, box.Y + 18, box.Width, box.Height - 18);

        double H = fp.ColumnFullHeight, baseT = fp.ColumnBaseThickness, bearT = fp.ColumnBearingThickness;
        double tubeLen = fp.ColumnTubeLength;
        double baseW = fp.ColumnBaseW, bearW = fp.ColumnBearingW, tubeW = fp.ColumnOuterWidth;
        if (H <= 0 || tubeLen <= 0) return;
        double maxW = Math.Max(Math.Max(baseW, bearW), Math.Max(tubeW, 0.1));

        const double bandL = 64, bandR = 58, bandT = 14, bandB = 16;
        double availW = area.Width - bandL - bandR, availH = area.Height - bandT - bandB;
        double scale = Math.Min(availW / maxW, availH / H);
        double drawH = H * scale;
        double cx = area.X + bandL + availW / 2;                  // welded centreline
        double baseY = area.Y + bandT + (availH - drawH) / 2 + drawH;   // bottom of the stack
        double Yb(double hFromBottom) => baseY - hFromBottom * scale;

        void Slab(double w, double y0, double y1)
        {
            double half = w * scale / 2.0;
            gfx.DrawRectangle(cutPen, new XRect(cx - half, Yb(y1), w * scale, (y1 - y0) * scale));
        }
        Slab(baseW, 0, baseT);                       // base plate (bottom)
        Slab(tubeW, baseT, baseT + tubeLen);         // tube / pipe (middle)
        Slab(bearW, baseT + tubeLen, H);             // bearing plate (top)

        double leftFace = cx - maxW * scale / 2;
        double rightFace = cx + maxW * scale / 2;

        // Tube name centred in the tall middle segment (room there); its cut length dimensioned left.
        gfx.DrawString(fp.ColumnShape == "round" ? "Pipe" : "Tube", tagFont, TextBrush,
            new XRect(cx - 28, Yb(baseT + tubeLen / 2.0) - 5, 56, 10), XStringFormats.TopCenter);
        DimV(gfx, dimFont, leftFace, area.X + 56, Yb(baseT + tubeLen), Yb(baseT), Fq(tubeLen), true);

        // Overall full height on the right.
        DimV(gfx, dimFont, rightFace, area.X + area.Width - 56, Yb(H), Yb(0), Fq(H), true);

        // Base + bearing plates are too thin to dimension inline — leader-call them out into the
        // clear space just below / above the column so nothing overlaps.
        var leadPen = new XPen(DimColor, 0.7);
        void PlateCallout(double slabY, double labelY, string text)
        {
            gfx.DrawLine(leadPen, leftFace, slabY, area.X + 74, labelY + 5);
            gfx.DrawString(text, tagFont, TextBrush, new XRect(area.X + 2, labelY, 84, 10), XStringFormats.TopLeft);
        }
        PlateCallout(Yb(baseT / 2.0), baseY + 4, $"Base plate {Fq(baseT)}");
        PlateCallout(Yb(baseT + tubeLen + bearT / 2.0), Yb(H) - 13, $"Bearing plate {Fq(bearT)}");
    }

    // ── Paddle blind / spade: face view (solid disc + handle) with OD, handle, centre-to-end dims ──
    private static void DrawPaddleBlind(XGraphics gfx, FlatPatternResult fp, XRect box)
    {
        var titleFont = new XFont("Arial", 9, XFontStyleEx.Bold);
        var dimFont = new XFont("Arial", 8, XFontStyleEx.Bold);
        var cutPen = new XPen(CutColor, 1.2);
        var centre = new XPen(DimColor, 0.5) { DashStyle = XDashStyle.DashDot };

        gfx.DrawString("Spade — face view", titleFont, XBrushes.Black,
            new XRect(box.X, box.Y, box.Width, 12), XStringFormats.TopCenter);
        var area = new XRect(box.X, box.Y + 18, box.Width, box.Height - 18);

        double mw = fp.FlatWidth, mh = fp.FlatHeight;     // [0..C+R] x [0..2R], disc centre at (R,R)
        double R = fp.Spec.PaddleOd / 2.0, hw = fp.Spec.PaddleHandleWidth / 2.0, C = fp.Spec.PaddleCenterToEnd;
        if (mw <= 0 || mh <= 0 || R <= 0) return;

        // Generous side bands so the OD (left) and handle-width (right) labels never reach the page
        // edge on the small, very wide sizes (long handle, small disc — width-limited fit).
        const double bandL = 92, bandR = 78, bandT = 28, bandB = 52;
        double availW = area.Width - bandL - bandR, availH = area.Height - bandT - bandB;
        double scale = Math.Min(availW / mw, availH / mh);
        double drawW = mw * scale, drawH = mh * scale;
        double ox = area.X + bandL + (availW - drawW) / 2;
        double oy = area.Y + bandT + (availH - drawH) / 2;
        XPoint P(double mx, double my) => new(ox + mx * scale, oy + drawH - my * scale);

        // Cut outline (the disc-plus-handle polyline).
        foreach (var e in fp.Cut.Entities)
            if (e.Type == "polyline" && e.Vertices is { Count: > 1 })
                gfx.DrawPolygon(cutPen, e.Vertices.Select(v => P(v.X, v.Y)).ToArray());

        // Centre axis (through disc + handle) + a short cross tick at the disc centre.
        gfx.DrawLine(centre, P(0, R), P(R + C, R));
        gfx.DrawLine(centre, P(R, R - R * 0.22), P(R, R + R * 0.22));

        // OD — vertical dim down the left of the disc (disc spans y 0..2R).
        DimV(gfx, dimFont, P(0, 0).X, P(0, 0).X - 30, P(0, 0).Y, P(0, 2 * R).Y, $"{Fq(fp.Spec.PaddleOd)} dia", true);
        // Centre-to-end — horizontal dim below the part, disc centre → handle tip.
        DimH(gfx, dimFont, P(R, R).X, P(R + C, R).X, oy + drawH, oy + drawH + 24, $"{Fq(C)} to end", true);
        // Handle width — vertical dim at the full-width handle station, just past the tip.
        double xW = R + C - hw;
        DimV(gfx, dimFont, P(xW, R).X, P(R + C, R).X + 12, P(xW, R - hw).Y, P(xW, R + hw).Y, Fq(fp.Spec.PaddleHandleWidth), false);
    }

    // ── Pan: single flat-pattern top view (cut outline + bend lines + corner reliefs) ──
    private static void DrawPan(XGraphics gfx, FlatPatternResult fp, XRect box)
    {
        var titleFont = new XFont("Arial", 9, XFontStyleEx.Bold);
        var dimFont = new XFont("Arial", 8, XFontStyleEx.Bold);
        var cutPen = new XPen(CutColor, 1.2);
        var bendPen = new XPen(BendColor, 0.9) { DashStyle = XDashStyle.Dash };

        gfx.DrawString("Flat pattern  (solid = cut, dashed = bend up)", titleFont, XBrushes.Black,
            new XRect(box.X, box.Y, box.Width, 12), XStringFormats.TopCenter);
        var area = new XRect(box.X, box.Y + 18, box.Width, box.Height - 18);

        // Bounds from the actual cut geometry (returns push vertices beyond [0,FlatWidth]).
        double minX = double.MaxValue, maxX = double.MinValue, minY = double.MaxValue, maxY = double.MinValue;
        void Acc(double x, double y) { if (x < minX) minX = x; if (x > maxX) maxX = x; if (y < minY) minY = y; if (y > maxY) maxY = y; }
        foreach (var e in fp.Cut.Entities)
        {
            if (e.Vertices is { } vs) foreach (var v in vs) Acc(v.X, v.Y);
            if (e.Type == "line") { Acc(e.X1, e.Y1); Acc(e.X2, e.Y2); }
        }
        if (minX > maxX) { minX = 0; maxX = fp.FlatWidth; minY = 0; maxY = fp.FlatHeight; }
        double L = maxX - minX, W = maxY - minY;
        if (L <= 0 || W <= 0) return;
        const double band = 60;
        double availW = area.Width - band * 2, availH = area.Height - band * 2;
        double scale = Math.Min(availW / L, availH / W);
        double drawW = L * scale, drawH = W * scale;
        double ox = area.X + (area.Width - drawW) / 2;
        double oy = area.Y + (area.Height - drawH) / 2;
        XPoint P(double mx, double my) => new(ox + (mx - minX) * scale, oy + drawH - (my - minY) * scale);

        foreach (var e in fp.Cut.Entities)
        {
            var pen = e.Layer == FlatPattern.BendLayer ? bendPen : cutPen;
            switch (e.Type)
            {
                case "polyline" when e.Vertices is { Count: > 1 }:
                    var pts = e.Vertices.Select(v => P(v.X, v.Y)).ToArray();
                    if (e.Closed) gfx.DrawPolygon(pen, pts); else gfx.DrawLines(pen, pts);
                    break;
                case "line":
                    gfx.DrawLine(pen, P(e.X1, e.Y1), P(e.X2, e.Y2));
                    break;
                case "circle":
                    double r = Math.Max(2.2, e.R * scale); var c = P(e.Cx, e.Cy);   // keep reliefs visible
                    gfx.DrawEllipse(pen, c.X - r, c.Y - r, 2 * r, 2 * r);
                    break;
            }
        }

        // Dimensions: base length & width (between bend lines) + the fold inset (flange flat).
        var s = fp.Spec;
        double bx0 = fp.PanBaseX0, bx1 = fp.PanBaseX1, by0 = fp.PanBaseY0, by1 = fp.PanBaseY1;
        DimH(gfx, dimFont, P(bx0, by0).X, P(bx1, by0).X, P(bx0, by0).Y, oy + drawH + 24, $"{Fq(bx1 - bx0)} OD", true);
        DimV(gfx, dimFont, P(bx0, by0).X, ox - 46, P(bx0, by0).Y, P(bx0, by1).Y, $"{Fq(by1 - by0)} OD", true);
        if (s.PanBottom && by0 > 0)
            DimV(gfx, dimFont, P(bx1, 0).X, P(bx1, 0).X + 26, P(bx1, 0).Y, P(bx1, by0).Y, Fq(fp.PanWallDev), false);
        else if (s.PanLeft && bx0 > 0)
            DimH(gfx, dimFont, P(0, by0).X, P(bx0, by0).X, P(0, by0).Y, oy + drawH + 24, Fq(fp.PanWallDev), false);

        // Finish callout — boxed label on the base (pans show inside/outside on the base face).
        if (s.Finish is FinishSide.Outside or FinishSide.Inside)
        {
            var f = new XFont("Arial", 9, XFontStyleEx.Bold);
            string txt = s.Finish == FinishSide.Outside ? "Finish: outside" : "Finish: inside";
            var bc = P((bx0 + bx1) / 2, (by0 + by1) / 2);
            var sz = gfx.MeasureString(txt, f);
            var br = new XRect(bc.X - sz.Width / 2 - 4, bc.Y - sz.Height / 2 - 2, sz.Width + 8, sz.Height + 4);
            gfx.DrawRectangle(XBrushes.White, br);
            gfx.DrawRectangle(new XPen(XColor.FromArgb(110, 110, 110), 0.9), br);
            gfx.DrawString(txt, f, XBrushes.Black, br, XStringFormats.Center);
        }
    }

    // ── Pan: formed-part isometric (the folded tray) ─────────────────────────
    private static void DrawPanIso(XGraphics gfx, FlatPatternResult fp, XRect box)
    {
        var titleFont = new XFont("Arial", 9, XFontStyleEx.Bold);
        var pen = new XPen(CutColor, 0.9);
        var faint = new XPen(XColor.FromArgb(150, 150, 150), 0.6);
        var floorFill = new XSolidBrush(XColor.FromArgb(238, 241, 247));
        var wallFill = new XSolidBrush(XColor.FromArgb(246, 248, 252));

        gfx.DrawString("Formed part", titleFont, XBrushes.Black,
            new XRect(box.X, box.Y, box.Width, 12), XStringFormats.TopCenter);
        var area = new XRect(box.X, box.Y + 16, box.Width, box.Height - 16);

        var s = fp.Spec;
        double Lo = fp.PanBaseX1 - fp.PanBaseX0, Wo = fp.PanBaseY1 - fp.PanBaseY0, Do = fp.PanDepth;
        const double c30 = 0.8660254, s30 = 0.5;
        (double u, double v) Iso(double x, double y, double z) => ((x - y) * c30, (x + y) * s30 - z);

        var corners = new[]
        {
            (0.0, 0.0, 0.0), (Lo, 0.0, 0.0), (Lo, Wo, 0.0), (0.0, Wo, 0.0),
            (0.0, 0.0, Do),  (Lo, 0.0, Do),  (Lo, Wo, Do),  (0.0, Wo, Do),
        };
        var pr = corners.Select(p => Iso(p.Item1, p.Item2, p.Item3)).ToList();
        double minU = pr.Min(p => p.u), maxU = pr.Max(p => p.u), minV = pr.Min(p => p.v), maxV = pr.Max(p => p.v);
        double gw = Math.Max(maxU - minU, 1e-6), gh = Math.Max(maxV - minV, 1e-6);
        const double pad = 38;
        double scale = Math.Min((area.Width - pad * 2) / gw, (area.Height - pad * 2) / gh);
        double ox = area.X + (area.Width - gw * scale) / 2, oy = area.Y + (area.Height - gh * scale) / 2;
        XPoint S(double x, double y, double z) { var (u, v) = Iso(x, y, z); return new XPoint(ox + (u - minU) * scale, oy + (v - minV) * scale); }

        // Floor.
        var floor = new[] { S(0, 0, 0), S(Lo, 0, 0), S(Lo, Wo, 0), S(0, Wo, 0) };
        gfx.DrawPolygon(floorFill, floor, XFillMode.Winding);
        gfx.DrawPolygon(faint, floor);

        // Walls (back walls first so the front ones overlap them). Every wall edge — including the
        // two vertical corner edges — is drawn so the folded corners read correctly. A return lip is
        // drawn at the top: 90° folds inward over the opening, 180° hems back down onto the wall.
        var ret = s.PanReturn;
        void Wall(bool present, double ax, double ay, double bx, double by, bool front)
        {
            if (!present) return;
            var p0 = S(ax, ay, 0); var p1 = S(bx, by, 0); var p2 = S(bx, by, Do); var p3 = S(ax, ay, Do);
            gfx.DrawPolygon(wallFill, new[] { p0, p1, p2, p3 }, XFillMode.Winding);
            var wp = front ? pen : faint;
            gfx.DrawLine(wp, p0, p1); gfx.DrawLine(wp, p1, p2); gfx.DrawLine(wp, p2, p3); gfx.DrawLine(wp, p3, p0);

            if (ret is not null)
            {
                double Lr = ret.Length;
                // Inward normal (toward the pan centre) in the XY plane.
                double ex = bx - ax, ey = by - ay; double el = Math.Sqrt(ex * ex + ey * ey); if (el < 1e-6) el = 1; ex /= el; ey /= el;
                double nx = -ey, ny = ex;
                double mx = (ax + bx) / 2, my = (ay + by) / 2, cx = Lo / 2, cy = Wo / 2;
                double dPlus = (mx + nx - cx) * (mx + nx - cx) + (my + ny - cy) * (my + ny - cy);
                double dMinus = (mx - nx - cx) * (mx - nx - cx) + (my - ny - cy) * (my - ny - cy);
                if (dPlus > dMinus) { nx = -nx; ny = -ny; }
                XPoint q0, q1, q2, q3;
                if (ret.AngleDeg >= 170)   // hem: folds back down along the wall face
                {
                    double th = s.Thickness;
                    q0 = S(ax, ay, Do); q1 = S(bx, by, Do);
                    q2 = S(bx + nx * th, by + ny * th, Do - Lr); q3 = S(ax + nx * th, ay + ny * th, Do - Lr);
                }
                else                        // 90° return: folds inward over the opening
                {
                    q0 = S(ax, ay, Do); q1 = S(bx, by, Do);
                    q2 = S(bx + nx * Lr, by + ny * Lr, Do); q3 = S(ax + nx * Lr, ay + ny * Lr, Do);
                }
                gfx.DrawPolygon(wallFill, new[] { q0, q1, q2, q3 }, XFillMode.Winding);
                gfx.DrawLine(wp, q1, q2); gfx.DrawLine(wp, q2, q3); gfx.DrawLine(wp, q3, q0);
            }
        }
        // Painter's order: far walls (back corner, top of image) first, near walls (front corner,
        // bottom of image) last — so the front walls sit ON TOP and aren't clipped by the back ones.
        Wall(s.PanBottom, 0,  0,  Lo, 0,  front: false);   // far  (y = 0)
        Wall(s.PanLeft,   0,  Wo, 0,  0,  front: false);   // far  (x = 0)
        Wall(s.PanTop,    Lo, Wo, 0,  Wo, front: true);    // near (y = Wo)
        Wall(s.PanRight,  Lo, 0,  Lo, Wo, front: true);    // near (x = Lo)

        // Section-cut planes: Side = green dash, End = orange dash-dot (distinct styles; keyed under
        // the drawing — no labels on the part).
        var sidePen = new XPen(SecColorA, 1.4) { DashStyle = XDashStyle.Dash };
        var endPen = new XPen(SecColorB, 1.4) { DashStyle = XDashStyle.DashDot };
        double xm = Lo / 2, ym = Wo / 2;
        gfx.DrawLine(sidePen, S(xm, 0, 0), S(xm, Wo, 0));
        if (s.PanBottom) gfx.DrawLine(sidePen, S(xm, 0, 0), S(xm, 0, Do));
        if (s.PanTop) gfx.DrawLine(sidePen, S(xm, Wo, 0), S(xm, Wo, Do));
        gfx.DrawLine(endPen, S(0, ym, 0), S(Lo, ym, 0));
        if (s.PanLeft) gfx.DrawLine(endPen, S(0, ym, 0), S(0, ym, Do));
        if (s.PanRight) gfx.DrawLine(endPen, S(Lo, ym, 0), S(Lo, ym, Do));

        // (No dimensions on the formed-part iso — the dimensioned views are the flat pattern + sections.)
    }

    // ── Generic radiused U cross-section (web + two walls), thickness shown like the channels ──
    private static void DrawSection(XGraphics gfx, XRect box, string title, XColor secColor, XDashStyle dash,
        List<(double x, double y)> prof, double webOD, double wallOD, double thickness, double? fixedScale = null)
    {
        var titleFont = new XFont("Arial", 9, XFontStyleEx.Bold);
        var dimFont = new XFont("Arial", 8, XFontStyleEx.Bold);
        var matPen = new XPen(CutColor, 1.1);

        // Title + a sample of this section's cut-plane style (colour + dash matched to the iso plane).
        var tsz = gfx.MeasureString(title, titleFont);
        double sampleW = 26, startX = box.X + (box.Width - tsz.Width - sampleW - 8) / 2;
        gfx.DrawString(title, titleFont, XBrushes.Black, new XRect(startX, box.Y, tsz.Width + 2, 12), XStringFormats.TopLeft);
        gfx.DrawLine(new XPen(secColor, 1.4) { DashStyle = dash },
            new XPoint(startX + tsz.Width + 8, box.Y + 7), new XPoint(startX + tsz.Width + 8 + sampleW, box.Y + 7));
        var area = new XRect(box.X, box.Y + 14, box.Width, box.Height - 14);
        if (prof.Count < 3) return;

        double minX = prof.Min(p => p.x), maxX = prof.Max(p => p.x);
        double minY = prof.Min(p => p.y), maxY = prof.Max(p => p.y);
        double mw = maxX - minX, mh = maxY - minY;
        if (mw <= 0 || mh <= 0) return;

        double availW = area.Width - SecBandL - SecBandR, availH = area.Height - SecBandB - SecBandT;
        double scale = fixedScale ?? Math.Min(availW / mw, availH / mh);   // shared scale → equal dims read equal
        double drawW = mw * scale, drawH = mh * scale;
        double ox = area.X + SecBandL + (availW - drawW) / 2;
        double oy = area.Y + SecBandT + (availH - drawH) / 2;
        XPoint P(double mx, double my) => new(ox + (mx - minX) * scale, oy + drawH - (my - minY) * scale);

        gfx.DrawPolygon(matPen, prof.Select(p => P(p.x, p.y)).ToArray());

        double sBottom = oy + drawH, sLeft = ox;
        DimH(gfx, dimFont, P(0, 0).X, P(webOD, 0).X, P(0, 0).Y, sBottom + 20, $"{Fq(webOD)} OD", true);
        DimV(gfx, dimFont, P(0, 0).X, sLeft - 22, P(0, 0).Y, P(0, wallOD).Y, $"{Fq(wallOD)} OD", true);

        // Thickness leader off the web, label placed up inside the open U (always clear of dims).
        var a = P(webOD * 0.5, 0); var b = P(webOD * 0.5, thickness);
        var lead = new XPen(DimColor, 0.6);
        var lp = P(webOD * 0.5, Math.Min(wallOD * 0.55, mh * 0.55));
        gfx.DrawLine(lead, a, b);
        gfx.DrawLine(lead, b, lp);
        gfx.DrawString($"T {Fq(thickness)}", dimFont, TextBrush, new XRect(lp.X - 28, lp.Y - 5, 56, 11), XStringFormats.TopCenter);
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
        // Paddle blinds and columns have no bends — drop the bend/Ri legend bits that don't apply.
        string legend = fp.IsPaddle || fp.IsColumn
            ? "solid = cut  |  bold = as specified  |  all dimensions in decimal inches"
            : $"solid = cut  |  dashed = bend up  |  bold = as specified  |  inside radius Ri {F(fp.Spec.InsideRadius)}\"  |  all dimensions in decimal inches";
        gfx.DrawString(legend, font, DimBrush, new XRect(box.X + 8, y, box.Width - 16, 10), XStringFormats.TopLeft);
    }

    // ── Key for the formed-part section-cut planes (Side = dashed, End = dash-dot) ──────────────
    private static void DrawSectionKey(XGraphics gfx, XRect box, bool isPan)
    {
        var kf = new XFont("Arial", 7, XFontStyleEx.Regular);
        var brushA = new XSolidBrush(SecColorA);
        var brushB = new XSolidBrush(SecColorB);
        gfx.DrawRectangle(new XPen(XColor.FromArgb(205, 205, 205), 0.8), box);
        gfx.DrawString("Section cuts", new XFont("Arial", 7, XFontStyleEx.Bold), XBrushes.Black,
            new XRect(box.X + 4, box.Y + 2, box.Width - 8, 9), XStringFormats.TopLeft);
        double ly = box.Y + 19, x = box.X + 6;
        // Match the iso planes: pan = Side (green dash) + End (orange dash-dot); single view = End (green dash-dot).
        if (isPan)
        {
            gfx.DrawLine(new XPen(SecColorA, 1.4) { DashStyle = XDashStyle.Dash }, new XPoint(x, ly), new XPoint(x + 22, ly));
            gfx.DrawString("Side", kf, brushA, new XRect(x + 25, ly - 5, 30, 9), XStringFormats.TopLeft);
            x += 78;
            gfx.DrawLine(new XPen(SecColorB, 1.4) { DashStyle = XDashStyle.DashDot }, new XPoint(x, ly), new XPoint(x + 22, ly));
            gfx.DrawString("End", kf, brushB, new XRect(x + 25, ly - 5, 30, 9), XStringFormats.TopLeft);
        }
        else
        {
            gfx.DrawLine(new XPen(SecColorA, 1.4) { DashStyle = XDashStyle.DashDot }, new XPoint(x, ly), new XPoint(x + 22, ly));
            gfx.DrawString("End", kf, brushA, new XRect(x + 25, ly - 5, 30, 9), XStringFormats.TopLeft);
        }
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

    /// <summary>Dimension label: value with the inch sign, e.g. 4.000".</summary>
    private static string Fq(double v) => F(v) + "\"";

    /// <summary>Best-fit scale (model→points) for a section profile inside <paramref name="box"/>.</summary>
    private static double SectionFitScale(XRect box, List<(double x, double y)> prof)
    {
        if (prof.Count < 3) return double.MaxValue;
        double availW = box.Width - SecBandL - SecBandR;
        double availH = (box.Height - 14) - SecBandB - SecBandT;
        double mw = prof.Max(p => p.x) - prof.Min(p => p.x);
        double mh = prof.Max(p => p.y) - prof.Min(p => p.y);
        if (mw <= 0 || mh <= 0) return double.MaxValue;
        return Math.Min(availW / mw, availH / mh);
    }
}
