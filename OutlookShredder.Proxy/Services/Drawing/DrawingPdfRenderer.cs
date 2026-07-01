using System.Globalization;
using System.Reflection;
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

    // Accent for the as-specified (highlighted) dimension — bold + this colour + a leader. Distinct
    // from the RoyalBlue bend lines and the muted gray derived dims.
    private static readonly XColor AccentColor = XColor.FromArgb(176, 0, 102);   // deep magenta
    private static readonly XBrush AccentBrush = new XSolidBrush(AccentColor);

    private static readonly string ForgeVersion =
        typeof(DrawingPdfRenderer).Assembly
            .GetCustomAttribute<System.Reflection.AssemblyInformationalVersionAttribute>()
            ?.InformationalVersion ?? "unknown";

    // Section-cut plane styling. Two distinct styles so multiple cut planes read apart:
    // A = green / dash (pan "side" + single-section views), B = orange / dash-dot (pan "end").
    private static readonly XColor SecColorA = XColor.FromArgb(0, 150, 70);
    private static readonly XColor SecColorB = XColor.FromArgb(202, 86, 0);

    private const double ExtGap = 2.0;
    private const double ExtOver = 3.0;

    // Section-view panel margins (shared so the pan's two sections render at one common scale).
    private const double SecBandL = 46, SecBandB = 42, SecBandT = 12, SecBandR = 18;

    // polishBilingual: the "Dirección de pulido" callout is bilingual ("Polish direction / …") on Pixar
    // PDFs but Spanish-only on the auto picking-slip drawings (the shop floor reads those). Default = the
    // Pixar (bilingual) behaviour; PickingSlipFabAppender passes false.
    // customerName: non-null triggers fab/picking-slip context for columns — draws the BOM header instead
    // of the regular Pixar title/spec table, and suppresses the Technics footnote.
    public static byte[] Render(FlatPatternResult fp, bool calibrate = false, bool polishBilingual = true,
        string? customerName = null, string? itemTag = null, string? createdBy = null)
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
            double usable = pw - 2 * M;
            const double gap = 16;

            bool fabColumn = fp.IsColumn && customerName != null;
            double top, footTop, footH;

            // FAB-note item letter (A, B, C…) drawn in the title block so it matches the lettered
            // note list in the ERP viewer and survives the 90° append rotation (it rides the title).
            double tagW = itemTag is not null ? 24 : 0;

            if (fabColumn)
            {
                // BOM header replaces the Pixar title + spec table; letter sits in the top-left margin above it.
                if (itemTag is not null) DrawItemTag(gfx, itemTag, M, M - 12);
                double bomBottom = DrawColumnBomHeader(gfx, fp, customerName!, M, M, usable);
                top = bomBottom + 8;
                footH = 92;   // info box only (technics suppressed)
                footTop = ph - M - footH;
            }
            else
            {
                if (itemTag is not null) DrawItemTag(gfx, itemTag, M, M - 10);
                var titleFont = new XFont("Arial", 16, XFontStyleEx.Bold);
                double maxTitleW = usable - tagW;
                double titleW = gfx.MeasureString(fp.Title, titleFont).Width;
                if (titleW > maxTitleW)
                    titleFont = new XFont("Arial", Math.Max(11.0, 16.0 * maxTitleW / titleW), XFontStyleEx.Bold);
                gfx.DrawString(fp.Title, titleFont, XBrushes.Black,
                    new XRect(M + tagW, M - 10, maxTitleW, 22), XStringFormats.TopLeft);
                double tableBottom = DrawSpecTable(gfx, fp, M, M + 16, maxTitleW);
                top = tableBottom + 10;
                int footLines = fp.Summary.Split('\n').Length;
                footH = Math.Max(footLines * 10 + 28, 92);
                footTop = ph - M - footH;
            }

            // L / U / Z three-panel split: flat 1/5, section 2/5, iso 2/5 (the iso gets the most room).
            double wFlat = (usable - 2 * gap) * 0.24;
            double wSect = (usable - 2 * gap) * 0.40;
            double wIso  = (usable - 2 * gap) * 0.36;
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
                DrawPolishCallout(gfx, fp, new XRect(M, top, wFlatP, topH), polishBilingual);
                DrawPanIso(gfx, fp, new XRect(isoX, top, wIsoP, topH - 46));
                DrawSectionKey(gfx, new XRect(isoX + (wIsoP - 250) / 2, top + topH - 36, 250, 30), isPan: true);

                // Both sections share ONE scale so an equal dimension reads the same size in each.
                double wSec = (usable - gap) / 2;
                var sideBox = new XRect(M, secY, wSec, botH);
                var endBox  = new XRect(M + wSec + gap, secY, wSec, botH);
                double secScale = Math.Min(SectionFitScale(sideBox, fp.PanSideProfile),
                                           SectionFitScale(endBox, fp.PanEndProfile));
                DrawSection(gfx, sideBox, Bi.T("sideSection"), SecColorA, XDashStyle.Dash,
                    fp.PanSideProfile, fp.PanBaseY1 - fp.PanBaseY0, fp.PanDepth,
                    fp.Spec.Width, fp.Spec.WidthBasis, fp.Spec.Depth, fp.Spec.DepthBasis, secScale);
                DrawSection(gfx, endBox, Bi.T("endSection"), SecColorB, XDashStyle.DashDot,
                    fp.PanEndProfile, fp.PanBaseX1 - fp.PanBaseX0, fp.PanDepth,
                    fp.Spec.Length, fp.Spec.LengthBasis, fp.Spec.Depth, fp.Spec.DepthBasis, secScale);
            }
            else if (fp.IsColumn)
            {
                DrawColumn(gfx, fp, new XRect(M, top, usable, h));
                if (!fabColumn)
                    DrawPolishCallout(gfx, fp, new XRect(M, top, usable, h), polishBilingual);
            }
            else if (fp.IsCircle)
            {
                DrawCircle(gfx, fp, new XRect(M, top, usable, h));
                DrawPolishCallout(gfx, fp, new XRect(M, top, usable, h), polishBilingual);
            }
            else if (fp.IsGusset)
            {
                DrawGusset(gfx, fp, new XRect(M, top, usable, h));
                DrawPolishCallout(gfx, fp, new XRect(M, top, usable, h), polishBilingual);
            }
            else if (fp.IsPlate)
            {
                DrawPlate(gfx, fp, new XRect(M, top, usable, h));
                DrawPolishCallout(gfx, fp, new XRect(M, top, usable, h), polishBilingual);
            }
            else if (fp.IsPaddle)
            {
                DrawPaddleBlind(gfx, fp, new XRect(M, top, usable, h));
                DrawPolishCallout(gfx, fp, new XRect(M, top, usable, h), polishBilingual);
            }
            else if (fp.IsMiterCut)
            {
                DrawMiterCut(gfx, fp, new XRect(M, top, usable, h));
            }
            else
            {
                DrawFlatPattern(gfx, fp, new XRect(M, top, wFlat, h));
                DrawPolishCallout(gfx, fp, new XRect(M, top, wFlat, h), polishBilingual);
                DrawCrossSection(gfx, fp, new XRect(M + wFlat + gap, top, wSect, h), calibrate);
                double isoX = M + wFlat + wSect + 2 * gap;
                // Single-section views: the "End section" header already carries the dash key, so we
                // drop the floating Section-cuts box and let the iso use the full panel height.
                DrawIsometric(gfx, fp, new XRect(isoX, top, wIso, h));
            }
            // Footer band: in fab-column mode only the info box is shown (no Technics).
            double infoGap = 6;
            if (fabColumn)
            {
                DrawInfoBox(gfx, new XRect(M, footTop, usable, footH), createdBy);
            }
            else
            {
                double infoW = (usable - infoGap) / 3.0;
                double techW = (usable - infoGap) * 2.0 / 3.0;
                DrawFootnote(gfx, fp, new XRect(M, footTop, techW, footH));
                DrawInfoBox(gfx, new XRect(M + techW + infoGap, footTop, infoW, footH), createdBy);
            }

            // Copyright line under the footer band.
            var copyFont = new XFont("Arial", 7, XFontStyleEx.Regular);
            gfx.DrawString($"Copyright {System.DateTime.Now.Year} Silmaril Corp. Forge Version: {ForgeVersion}.",
                copyFont, DimBrush, new XRect(M, footTop + footH + 2, usable, 10), XStringFormats.TopLeft);
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

        gfx.DrawString(Bi.T("flatPattern.cut"), titleFont, XBrushes.Black,
            new XRect(box.X, box.Y, box.Width, 12), XStringFormats.TopCenter);
        var area = new XRect(box.X, box.Y + 16, box.Width, box.Height - 16);

        double mw = fp.FlatWidth, mh = fp.FlatHeight;
        if (mw <= 0 || mh <= 0) return;

        const double dimBand = 40;   // top band must clear the staggered bend-dim chain from the title
        double availW = area.Width - dimBand * 2, availH = area.Height - dimBand * 2;
        double scale = Math.Min(availW / mw, availH / mh);
        double drawW = mw * scale, drawH = mh * scale;
        double ox = area.X + dimBand + (availW - drawW) / 2;
        double oy = area.Y + dimBand + (availH - drawH) / 2;

        XPoint P(double mx, double my) => new(ox + mx * scale, oy + drawH - my * scale);

        gfx.DrawRectangle(cutPen, new XRect(ox, oy, drawW, drawH));
        foreach (var bx in fp.BendLinesX)
            gfx.DrawLine(bendPen, P(bx, 0), P(bx, mh));

        // Flat-blank dimensions are cut dimensions — always decimal inches (developed lengths rarely
        // land on a clean fraction, and the shop cuts to the precise value).
        DimH(gfx, labelFont, P(0, 0).X, P(mw, 0).X, oy + drawH, oy + drawH + 22, DrawFormat.DecInch(mw), false);
        DimV(gfx, labelFont, ox, ox - 24, oy, oy + drawH, DrawFormat.DecInch(mh), false);

        double prev = 0;
        int idx = 0;
        foreach (var bx in fp.BendLinesX)
        {
            double dimY = oy - 16 - (idx % 2) * 14;   // stagger so labels don't collide on narrow blanks
            DimH(gfx, labelFont, P(prev, 0).X, P(bx, 0).X, oy, dimY, DrawFormat.DecInch(bx - prev), false);
            prev = bx;
            idx++;
        }
    }

    // ── Polish / grain-direction callout (screen + PDF only — NOT the DXF) ─────
    // A double-headed arrow along the chosen axis (centred on the panel/part); the "Dirección de pulido"
    // label sits in the panel margin OUTSIDE the part, oriented PARALLEL to the arrow (rotated 90° for a
    // vertical axis) and tied back to the arrow with a short leader. No-op when unset.
    // NB: this is the PDF label ONLY. The DXF carries its own PolishLabel geometry on the no-cut L1
    // layer, which the PDF render explicitly skips (see the cut-geometry loops) so it never leaks here.
    // Placed INDEPENDENTLY of the finish callout by design — they annotate different drawing components.
    private static void DrawPolishCallout(XGraphics gfx, FlatPatternResult fp, XRect box, bool bilingual)
    {
        if (fp.Spec.PolishDirection == PolishDirection.None) return;
        bool vertical = fp.Spec.PolishDirection == PolishDirection.Vertical;

        var pen  = new XPen(XColor.FromArgb(120, 70, 180), 1.4);   // violet — distinct from cut/bend/dim
        var thin = new XPen(XColor.FromArgb(120, 70, 180), 0.6);   // leader from the label to the arrow
        var font = new XFont("Arial", 7, XFontStyleEx.Bold);
        // Bilingual on Pixar PDFs; Spanish-only on the auto picking-slip drawings the shop reads.
        string label = bilingual ? Bi.T("polish.direction") : Bi.Es("polish.direction");
        var sz = gfx.MeasureString(label, font);
        const double head = 5, pad = 3;

        double cx = box.X + box.Width * 0.5, cy = box.Y + box.Height * 0.5;

        if (vertical)
        {
            double half = box.Height * 0.09;   // arrow ~18% of the view tall — sits well inside the part
            var tp = new XPoint(cx, cy - half); var bt = new XPoint(cx, cy + half);
            gfx.DrawLine(pen, tp, bt);
            gfx.DrawLine(pen, tp, new XPoint(cx - head, tp.Y + head)); gfx.DrawLine(pen, tp, new XPoint(cx + head, tp.Y + head));
            gfx.DrawLine(pen, bt, new XPoint(cx - head, bt.Y - head)); gfx.DrawLine(pen, bt, new XPoint(cx + head, bt.Y - head));

            // Vertical (rotated) label hugging the right margin, parallel to the arrow; leader back to it.
            var lc = new XPoint(box.X + box.Width - pad - sz.Height / 2.0, cy);
            gfx.DrawLine(thin, new XPoint(cx + 1, cy), new XPoint(lc.X - sz.Height / 2.0, cy));
            var st = gfx.Save();
            gfx.TranslateTransform(lc.X, lc.Y);
            gfx.RotateTransform(-90);
            gfx.DrawString(label, font, XBrushes.Black,
                new XRect(-sz.Width / 2, -sz.Height / 2, sz.Width, sz.Height), XStringFormats.Center);
            gfx.Restore(st);
        }
        else
        {
            double half = box.Width * 0.09;   // arrow ~18% of the view wide — sits well inside the part
            var lf = new XPoint(cx - half, cy); var rg = new XPoint(cx + half, cy);
            gfx.DrawLine(pen, lf, rg);
            gfx.DrawLine(pen, lf, new XPoint(lf.X + head, cy - head)); gfx.DrawLine(pen, lf, new XPoint(lf.X + head, cy + head));
            gfx.DrawLine(pen, rg, new XPoint(rg.X - head, cy - head)); gfx.DrawLine(pen, rg, new XPoint(rg.X - head, cy + head));

            // Horizontal label in the bottom margin, parallel to the arrow, centred + clamped; leader up.
            double ly = box.Y + box.Height - pad - sz.Height;
            double lx = Math.Max(box.X + pad, Math.Min(cx - sz.Width / 2, box.X + box.Width - sz.Width - pad));
            gfx.DrawLine(thin, new XPoint(cx, cy + 1), new XPoint(cx, ly - 1));
            gfx.DrawString(label, font, XBrushes.Black,
                new XRect(lx, ly, sz.Width, sz.Height), XStringFormats.TopLeft);
        }
    }

    // ── 2. Dimensioned end-section (any shape, primary dims only) ─────────────
    private static void DrawCrossSection(XGraphics gfx, FlatPatternResult fp, XRect box, bool calibrate = false)
    {
        var titleFont = new XFont("Arial", 9, XFontStyleEx.Bold);
        var dimFont   = new XFont("Arial", 8, XFontStyleEx.Bold);
        var matPen    = new XPen(CutColor, 1.1);

        string secTitle = Bi.T("endSection");
        var tsz = gfx.MeasureString(secTitle, titleFont);
        double tStart = box.X + (box.Width - tsz.Width - 34) / 2;
        gfx.DrawString(secTitle, titleFont, XBrushes.Black, new XRect(tStart, box.Y, tsz.Width + 2, 12), XStringFormats.TopLeft);
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

        double webO = fp.WebOutside, flL = fp.FlangeLeftOutside, flR = fp.FlangeRightOutside;
        // Value + the Spanish basis word (Adentro / Afuera). No thickness callout on the geometry.
        string BL(double v, bool inside) => $"{Fq(v)} {Bi.Basis(inside ? DimBasis.Inside : DimBasis.Outside)}";

        // Page-space part bbox + centroid drive outward dim/label placement and collision avoidance.
        var ptsPage = prof.Select(p => P(p.x, p.y)).ToArray();
        var centroid = new XPoint(ptsPage.Average(p => p.X), ptsPage.Average(p => p.Y));
        var placed = new List<XRect> { BBox(ptsPage) };

        // Aligned dimensions anchored to the TRUE outer/inner sharp corners (intersection of the offset
        // faces), so the witness lines land exactly on the material edges even on thick stock. Lips are
        // pushed a little further out so they don't crash into the flange dimension.
        var csDims = ComputeCrossSectionDims(fp);
        for (int di = 0; di < csDims.Count; di++)
        {
            var d = csDims[di];
            double off = d.Kind == DimKind.Lip ? 30 : 24;
            // The lip dim runs along the lip's own face (set in ComputeCrossSectionDims). A 90° return's
            // label sits on its outer face (the default away-from-centroid offset). A 180° hem runs back
            // alongside the flange, where away-from-centroid would crash the flange dim — push it to the
            // opposite (interior) side instead.
            (double x, double y)? forceDir = null;
            if (d.Kind == DimKind.Lip && d.Hem)
            {
                var p1 = P(d.X1, d.Y1); var p2 = P(d.X2, d.Y2);
                var mid = new XPoint((p1.X + p2.X) / 2, (p1.Y + p2.Y) / 2);
                forceDir = (centroid.X - mid.X, centroid.Y - mid.Y);   // toward centroid (mirror of flange dim)
            }
            // Calibration: tag each dimension with a letter (A, B, C…) so the user can map it to geometry.
            string? tag = calibrate ? ((char)('A' + di)).ToString() : null;
            DimAligned(gfx, dimFont, P(d.X1, d.Y1), P(d.X2, d.Y2), off, centroid, BL(d.Value, d.Inside), true, placed, area, forceDir, tag);
        }

        // ── "Finish" callout: boxed, highlighted label + leader to the finished face ──
        var finishFont = new XFont("Arial", 10, XFontStyleEx.Bold);
        void FinishCallout(double mx, double my, double sdx, double sdy)
        {
            var tip = P(mx, my);
            double l = Math.Sqrt(sdx * sdx + sdy * sdy); if (l < 1e-6) l = 1;
            double ux = sdx / l, uy = sdy / l;

            // Boxed label, offset clear of the part; clamp inside the panel so it never clips the border.
            var sz = gfx.MeasureString(Bi.T("finish"), finishFont);
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
            gfx.DrawString(Bi.T("finish"), finishFont, XBrushes.Black, br, XStringFormats.Center);
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
            case FinishSide.Outside when fp.Spec.Type == PartType.ZChannel:   // Z carries Outside/Inside
                FinishCallout(webO * 0.62, t2, 0, -1); break;           // first flange top face
            case FinishSide.Inside when fp.Spec.Type == PartType.ZChannel:
                FinishCallout(webO * 0.38, -t2, 0, 1); break;           // first flange bottom face
            case FinishSide.Top:                                        // Z: first flange's top face
                FinishCallout(webO * 0.62, t2, 0, -1); break;
            case FinishSide.Bottom:                                     // Z: first flange's bottom face
                FinishCallout(webO * 0.38, -t2, 0, 1); break;
        }

        // ── Per-bend callouts: a right-angle mark for 90° bends, a degree arc + leader (arrow on the
        //    arc) for other angles. Returns/hems carry no callout — the drawn lip makes them plain. ──
        if (fp.SectionBends.Count > 0)
        {
            var angFont = new XFont("Arial", 8, XFontStyleEx.Bold);
            var arcPen = new XPen(DimColor, 0.8);
            const double r = 12;

            foreach (var b in fp.SectionBends)
            {
                if (b.IsReturn) continue;

                var apex = P(b.X, b.Y);
                // Face rays from the apex (page space — y points down, so negate model hy).
                double inx = b.InHx, iny = -b.InHy, outx = b.OutHx, outy = -b.OutHy;
                double il = Math.Sqrt(inx * inx + iny * iny); if (il < 1e-6) il = 1; inx /= il; iny /= il;
                double ol = Math.Sqrt(outx * outx + outy * outy); if (ol < 1e-6) ol = 1; outx /= ol; outy /= ol;
                double a1 = Math.Atan2(-iny, -inx), a2 = Math.Atan2(outy, outx);
                double da = a2 - a1; while (da <= -Math.PI) da += 2 * Math.PI; while (da > Math.PI) da -= 2 * Math.PI;
                double shownDeg = Math.Abs(da) * 180.0 / Math.PI;

                if (Math.Abs(shownDeg - 90) < 1.5)
                {
                    // Right-angle mark in the corner (no degree text). Faces from the apex: -in and +out.
                    RightAngleMark(gfx, arcPen, apex, new XPoint(-inx, -iny), new XPoint(outx, outy), 7);
                    continue;
                }

                // Angled bend: arc spanning the two faces + a degree leader whose arrow touches the arc.
                XPoint Prev = new(apex.X + r * Math.Cos(a1), apex.Y + r * Math.Sin(a1));
                for (int i = 1; i <= 10; i++)
                {
                    double a = a1 + da * i / 10.0;
                    var cur = new XPoint(apex.X + r * Math.Cos(a), apex.Y + r * Math.Sin(a));
                    gfx.DrawLine(arcPen, Prev, cur); Prev = cur;
                }
                double amid = a1 + da / 2.0;
                var arcMid = new XPoint(apex.X + r * Math.Cos(amid), apex.Y + r * Math.Sin(amid));
                LeaderLabel(gfx, angFont, arcMid, Unit(arcMid.X - centroid.X, arcMid.Y - centroid.Y), $"{shownDeg:0.#}°", false, placed, area);
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

        gfx.DrawString(Bi.T("formedPart"), titleFont, XBrushes.Black,
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
        RotText(gfx, dimFont, $"Largo {Fq(len)}", new XPoint((Ao.X + Bo.X) / 2, (Ao.Y + Bo.Y) / 2), Math.Atan2(Bo.Y - Ao.Y, Bo.X - Ao.X));
    }

    // ── Flat plate: single dimensioned top view with bolt holes ──────────────
    private static void DrawPlate(XGraphics gfx, FlatPatternResult fp, XRect box)
    {
        var titleFont = new XFont("Arial", 9, XFontStyleEx.Bold);
        var dimFont = new XFont("Arial", 8, XFontStyleEx.Bold);
        var cutPen = new XPen(CutColor, 1.2);
        var thin = new XPen(DimColor, 0.5);

        gfx.DrawString(Bi.T("plate.topView"), titleFont, XBrushes.Black,
            new XRect(box.X, box.Y, box.Width, 12), XStringFormats.TopCenter);
        var area = new XRect(box.X, box.Y + 18, box.Width, box.Height - 18);

        double L = fp.FlatWidth, W = fp.FlatHeight;
        if (L <= 0 || W <= 0) return;
        const double band = 52;   // left band must hold the height dim label without clipping the page
        double availW = area.Width - band * 2, availH = area.Height - band * 2;
        double scale = Math.Min(availW / L, availH / W);
        double drawW = L * scale, drawH = W * scale;
        double ox = area.X + (area.Width - drawW) / 2;
        double oy = area.Y + (area.Height - drawH) / 2;
        XPoint P(double mx, double my) => new(ox + mx * scale, oy + drawH - my * scale);

        if (fp.Spec.CornerRadius > 0)
        {
            // Rounded corners (base plate): ellipse axis = 2 x radius, clamped to half the shorter side.
            double er = Math.Min(fp.Spec.CornerRadius * scale, Math.Min(drawW, drawH) / 2.0) * 2.0;
            gfx.DrawRoundedRectangle(cutPen, ox, oy, drawW, drawH, er, er);
        }
        else
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
            if (!hsr.SingleRow)
                DimV(gfx, dimFont, ox, ox - 22, P(0, 0).Y, P(0, rowBot).Y, Fq(botEdge), false);
        }

        if (fp.Holes.Count > 0)
        {
            var (hx, hy, dia) = fp.Holes[0];
            var c = P(hx, hy);
            double r = dia / 2.0 * scale;
            string label = $"{fp.Holes.Count} × {Fq(dia)} {Bi.T("dia")}";
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

    // ── Structural column BOM header (fab/picking-slip context) — replaces the Pixar title+spec band. ──
    // Returns the bottom Y of the header so the caller can position the CAD panels below it.
    private static double DrawColumnBomHeader(XGraphics gfx, FlatPatternResult fp, string customerName,
        double x, double y, double w)
    {
        const double rowH = 15.0;
        var boldFont  = new XFont("Arial", 13, XFontStyleEx.Bold);
        var descFont  = new XFont("Arial", 10, XFontStyleEx.Bold);
        var hdrFont   = new XFont("Arial", 8, XFontStyleEx.Bold);
        var cellFont  = new XFont("Arial", 8, XFontStyleEx.Regular);
        var boldCell  = new XFont("Arial", 8, XFontStyleEx.Bold);
        var hdrBrush  = new XSolidBrush(XColor.FromArgb(30, 91, 170));
        var linePen   = new XPen(XColor.FromArgb(200, 200, 200), 0.5);
        var borderPen = new XPen(XColor.FromArgb(160, 160, 160), 0.8);

        double cy = y;

        // "PREPARED FOR {customer}"
        gfx.DrawString($"PREPARED FOR {customerName.ToUpper()}", boldFont, XBrushes.Black,
            new XRect(x, cy, w, 18), XStringFormats.TopLeft);
        cy += 20;

        // Job description
        gfx.DrawString(ColumnBomJobDescription(fp), descFont, XBrushes.Black,
            new XRect(x, cy, w, 14), XStringFormats.TopLeft);
        cy += 16;

        // BOM table — 4 columns: METAL (46%), SIZE (30%), QTY (10%), LABEL (14%)
        double cM = w * 0.46, cS = w * 0.30, cQ = w * 0.10, cL = w * 0.14;
        double[] colX = { x, x + cM, x + cM + cS, x + cM + cS + cQ };
        double[] colW = { cM, cS, cQ, cL };

        // Header row
        double tableTop = cy;
        gfx.DrawRectangle(hdrBrush, new XRect(x, cy, w, rowH));
        string[] hdrs = { "METAL", "SIZE", "QTY", "" };
        for (int i = 0; i < 4; i++)
            gfx.DrawString(hdrs[i], hdrFont, XBrushes.White,
                new XRect(colX[i] + 2, cy + 1, colW[i] - 4, rowH - 2), XStringFormats.TopLeft);
        cy += rowH;

        void Row(string metal, string size, string qtyStr, string lbl, bool bold = false)
        {
            gfx.DrawLine(linePen, x, cy, x + w, cy);
            var f = bold ? boldCell : cellFont;
            gfx.DrawString(metal,  f, TextBrush, new XRect(colX[0] + 2, cy + 1, colW[0] - 4, rowH - 2), XStringFormats.TopLeft);
            gfx.DrawString(size,   f, TextBrush, new XRect(colX[1] + 2, cy + 1, colW[1] - 4, rowH - 2), XStringFormats.TopLeft);
            gfx.DrawString(qtyStr, f, TextBrush, new XRect(colX[2] + 2, cy + 1, colW[2] - 4, rowH - 2), XStringFormats.TopLeft);
            gfx.DrawString(lbl,    f, TextBrush, new XRect(colX[3] + 2, cy + 1, colW[3] - 4, rowH - 2), XStringFormats.TopLeft);
            cy += rowH;
        }

        string qtyS = fp.ColumnQty.ToString();
        string tubeMetal = !string.IsNullOrWhiteSpace(fp.ColumnProductName)
            ? fp.ColumnProductName.ToUpper() : $"{fp.ColumnShape.ToUpper()} TUBE";
        Row(tubeMetal, $"{Fq(fp.ColumnTubeLength)}\"", qtyS, "");

        string pm = (fp.ColumnPlateMetal ?? "HOT ROLL").ToUpper();
        if (fp.ColumnBaseIncluded)
            Row($"{pm} PLATE", $"{Fq(fp.ColumnBaseThickness)}\" {Fq(fp.ColumnBaseL)}\" x {Fq(fp.ColumnBaseW)}\"", qtyS, "BASE");
        if (fp.ColumnBearingIncluded)
            Row($"{pm} PLATE", $"{Fq(fp.ColumnBearingThickness)}\" {Fq(fp.ColumnBearingL)}\" x {Fq(fp.ColumnBearingW)}\"", qtyS, "BEARING");

        bool wB = fp.ColumnBaseIncluded && fp.ColumnBaseWelded;
        bool wR = fp.ColumnBearingIncluded && fp.ColumnBearingWelded;
        if (wB || wR)
        {
            string weldText = (wB && wR) ? "WELD TOP AND BOTTOM" : wB ? "WELD BOTTOM" : "WELD TOP";
            gfx.DrawLine(linePen, x, cy, x + w, cy);
            gfx.DrawString(weldText, boldCell, TextBrush, new XRect(colX[0] + 2, cy + 1, cM + cS - 4, rowH - 2), XStringFormats.TopLeft);
            cy += rowH;
        }

        Row("TOTAL COLUMN HEIGHT", $"{Fq(fp.ColumnFullHeight)}\"", "", "", bold: true);

        gfx.DrawRectangle(borderPen, new XRect(x, tableTop, w, cy - tableTop));
        return cy + 6;
    }

    private static string ColumnBomJobDescription(FlatPatternResult fp)
    {
        bool bI = fp.ColumnBaseIncluded, rI = fp.ColumnBearingIncluded;
        bool wB = bI && fp.ColumnBaseWelded, wR = rI && fp.ColumnBearingWelded;
        if (!bI && !rI) return "CUT COLUMNS";
        if (bI && rI) return (wB && wR) ? "COLUMNS WITH WELDED BASE & BEARING PLATES"
            : wB ? "COLUMNS WITH WELDED BASE PLATE"
            : wR ? "COLUMNS WITH WELDED BEARING PLATE"
            : "COLUMNS WITH BASE & BEARING PLATES";
        if (bI) return wB ? "COLUMNS WITH WELDED BASE PLATE" : "COLUMNS WITH BASE PLATE";
        return wR ? "COLUMNS WITH WELDED BEARING PLATE" : "COLUMNS WITH BEARING PLATE";
    }

    // ── Structural column: included plate top-views (left) + a dimensioned side elevation (right) ──
    private static void DrawColumn(XGraphics gfx, FlatPatternResult fp, XRect box)
    {
        const double gap = 16;
        bool baseIncl = fp.ColumnBaseIncluded;
        bool bearIncl = fp.ColumnBearingIncluded;
        int platePanels = (baseIncl ? 1 : 0) + (bearIncl ? 1 : 0);

        if (platePanels == 0)
        {
            DrawColumnElevation(gfx, fp, box);
            return;
        }

        double elevW = box.Width * 0.34;
        double platesW = box.Width - elevW - gap;
        double panelW = (platesW - (platePanels - 1) * gap) / platePanels;
        double x = box.X;

        if (baseIncl)
        {
            DrawPlatePanel(gfx, new XRect(x, box.Y, panelW, box.Height),
                Bi.T("basePlate.topView"), fp.ColumnBaseL, fp.ColumnBaseW, fp.ColumnBaseHoles,
                fp.ColumnShape, fp.ColumnOuterWidth, fp.ColumnOuterDepth, fp.ColumnWall, fp.ColumnTubeCornerR);
            x += panelW + gap;
        }
        if (bearIncl)
        {
            DrawPlatePanel(gfx, new XRect(x, box.Y, panelW, box.Height),
                Bi.T("bearingPlate.topView"), fp.ColumnBearingL, fp.ColumnBearingW, fp.ColumnBearingHoles,
                fp.ColumnShape, fp.ColumnOuterWidth, fp.ColumnOuterDepth, fp.ColumnWall, fp.ColumnTubeCornerR);
        }
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

        gfx.DrawString(title, titleFont, XBrushes.Black,
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
            gfx.DrawString($"{holes.Count} × {Fq(holes[0].dia)} {Bi.T("dia")}", dimFont, TextBrush,
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
        gfx.DrawString(Bi.T("tube"), new XFont("Arial", 7, XFontStyleEx.Regular), new XSolidBrush(SecColorA),
            new XRect(wc.X - 32, wc.Y - 4, 64, 9), XStringFormats.TopCenter);
    }

    // ── Column side elevation: base / tube / bearing stacked, welded centred, with callouts ──
    private static void DrawColumnElevation(XGraphics gfx, FlatPatternResult fp, XRect box)
    {
        var titleFont = new XFont("Arial", 9, XFontStyleEx.Bold);
        var dimFont = new XFont("Arial", 8, XFontStyleEx.Bold);
        var tagFont = new XFont("Arial", 7, XFontStyleEx.Regular);
        var cutPen = new XPen(CutColor, 1.1);

        gfx.DrawString(Bi.T("columnElevation"), titleFont, XBrushes.Black,
            new XRect(box.X, box.Y, box.Width, 12), XStringFormats.TopCenter);
        var area = new XRect(box.X, box.Y + 18, box.Width, box.Height - 18);

        double H = fp.ColumnFullHeight;
        double baseT = fp.ColumnBaseThickness, bearT = fp.ColumnBearingThickness;
        double tubeLen = fp.ColumnTubeLength;
        double baseW = fp.ColumnBaseW, bearW = fp.ColumnBearingW, tubeW = fp.ColumnOuterWidth;
        if (H <= 0 || tubeLen <= 0) return;

        // Only include present plates in the max-width calculation.
        double maxW = tubeW;
        if (fp.ColumnBaseIncluded)    maxW = Math.Max(maxW, baseW);
        if (fp.ColumnBearingIncluded) maxW = Math.Max(maxW, bearW);
        maxW = Math.Max(maxW, 0.1);

        const double bandL = 48, bandR = 44, bandT = 14, bandB = 16;
        double availW = area.Width - bandL - bandR, availH = area.Height - bandT - bandB;
        double scale = Math.Min(availW / maxW, availH / H);
        double drawH = H * scale;
        double cx = area.X + bandL + availW / 2;                       // welded centreline
        double baseY = area.Y + bandT + (availH - drawH) / 2 + drawH; // bottom of the stack
        double Yb(double hFromBottom) => baseY - hFromBottom * scale;

        // Effective base height for stacking (0 when base plate not included).
        double effBaseT = fp.ColumnBaseIncluded ? baseT : 0;

        void Slab(double w, double y0, double y1)
        {
            double half = w * scale / 2.0;
            gfx.DrawRectangle(cutPen, new XRect(cx - half, Yb(y1), w * scale, (y1 - y0) * scale));
        }
        if (fp.ColumnBaseIncluded)    Slab(baseW, 0, baseT);
        Slab(tubeW, effBaseT, effBaseT + tubeLen);
        if (fp.ColumnBearingIncluded) Slab(bearW, effBaseT + tubeLen, H);

        double leftFace  = cx - maxW * scale / 2;
        double rightFace = cx + maxW * scale / 2;

        // Tube name centred in the middle segment; its cut length dimensioned left.
        gfx.DrawString(fp.ColumnShape == "round" ? Bi.T("pipe.cap") : Bi.T("tube.cap"), tagFont, TextBrush,
            new XRect(cx - 40, Yb(effBaseT + tubeLen / 2.0) - 5, 80, 10), XStringFormats.TopCenter);
        DimV(gfx, dimFont, leftFace, area.X + 56, Yb(effBaseT + tubeLen), Yb(effBaseT), Fq(tubeLen), true);

        // Overall full height on the right.
        DimV(gfx, dimFont, rightFace, area.X + area.Width - 56, Yb(H), Yb(0), Fq(H), true);

        // Leader callouts for thin plates (only when included).
        var leadPen = new XPen(DimColor, 0.7);
        void PlateCallout(double slabY, double labelY, string text)
        {
            gfx.DrawLine(leadPen, leftFace, slabY, area.X + 74, labelY + 5);
            gfx.DrawString(text, tagFont, TextBrush, new XRect(area.X + 2, labelY, 84, 10), XStringFormats.TopLeft);
        }
        if (fp.ColumnBaseIncluded)
            PlateCallout(Yb(baseT / 2.0), baseY + 4, $"Base plate {Fq(baseT)}");
        if (fp.ColumnBearingIncluded)
            PlateCallout(Yb(effBaseT + tubeLen + bearT / 2.0), Yb(H) - 13, $"Bearing plate {Fq(bearT)}");
    }

    // ── Paddle blind / spade: face view (solid disc + handle) with OD, handle, centre-to-end dims ──
    private static void DrawPaddleBlind(XGraphics gfx, FlatPatternResult fp, XRect box)
    {
        var titleFont = new XFont("Arial", 9, XFontStyleEx.Bold);
        var dimFont = new XFont("Arial", 8, XFontStyleEx.Bold);
        var cutPen = new XPen(CutColor, 1.2);
        var centre = new XPen(DimColor, 0.5) { DashStyle = XDashStyle.DashDot };

        gfx.DrawString(Bi.T("spade.faceView"), titleFont, XBrushes.Black,
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

        // Cut outline (the disc-plus-handle polyline). Skip the no-cut L1 label layer — that geometry
        // (PartLabel / PolishLabel) is for the DXF only; the PDF carries its own labels.
        foreach (var e in fp.Cut.Entities)
            if (e.Type == "polyline" && e.Vertices is { Count: > 1 } && e.Layer != PartLabel.LayerName)
                gfx.DrawPolygon(cutPen, e.Vertices.Select(v => P(v.X, v.Y)).ToArray());

        // Centre axis (through disc + handle) + a short cross tick at the disc centre.
        gfx.DrawLine(centre, P(0, R), P(R + C, R));
        gfx.DrawLine(centre, P(R, R - R * 0.22), P(R, R + R * 0.22));

        // OD — vertical dim down the left of the disc (disc spans y 0..2R).
        DimV(gfx, dimFont, P(0, 0).X, P(0, 0).X - 30, P(0, 0).Y, P(0, 2 * R).Y, $"{Fq(fp.Spec.PaddleOd)} {Bi.T("dia")}", true);
        // Centre-to-end — horizontal dim below the part, disc centre → handle tip.
        DimH(gfx, dimFont, P(R, R).X, P(R + C, R).X, oy + drawH, oy + drawH + 24, $"{Fq(C)} {Bi.T("toEnd")}", true);
        // Handle width — vertical dim at the full-width handle station, just past the tip.
        double xW = R + C - hw;
        DimV(gfx, dimFont, P(xW, R).X, P(R + C, R).X + 12, P(xW, R - hw).Y, P(xW, R + hw).Y, Fq(fp.Spec.PaddleHandleWidth), false);
    }

    // ── Circle / disc (and donut): flat face view with Ø dimension(s) ──
    private static void DrawCircle(XGraphics gfx, FlatPatternResult fp, XRect box)
    {
        var titleFont = new XFont("Arial", 9, XFontStyleEx.Bold);
        var dimFont = new XFont("Arial", 8, XFontStyleEx.Bold);
        var cutPen = new XPen(CutColor, 1.2);
        var centre = new XPen(DimColor, 0.5) { DashStyle = XDashStyle.DashDot };

        gfx.DrawString(Bi.T("flatPattern.cut"), titleFont, XBrushes.Black,
            new XRect(box.X, box.Y, box.Width, 12), XStringFormats.TopCenter);
        var area = new XRect(box.X, box.Y + 18, box.Width, box.Height - 18);

        double od = fp.Spec.Diameter, R = od / 2.0, innerD = fp.Spec.InnerDiameter;
        if (od <= 0) return;

        // Bands hold the OD label (left) + the cross-axis ticks; keep the disc clear of the page edge.
        const double bandL = 64, bandR = 32, bandT = 24, bandB = 40;
        double availW = area.Width - bandL - bandR, availH = area.Height - bandT - bandB;
        double scale = Math.Min(availW / od, availH / od);
        double drawSz = od * scale;
        double ox = area.X + bandL + (availW - drawSz) / 2;
        double oy = area.Y + bandT + (availH - drawSz) / 2;
        XPoint P(double mx, double my) => new(ox + mx * scale, oy + drawSz - my * scale);

        // Disc outline (+ donut bore). Origin lower-left; disc centre at (R, R).
        var c = P(R, R);
        gfx.DrawEllipse(cutPen, c.X - R * scale, c.Y - R * scale, drawSz, drawSz);
        if (innerD > 0)
        {
            double ir = innerD / 2.0 * scale;
            gfx.DrawEllipse(cutPen, c.X - ir, c.Y - ir, 2 * ir, 2 * ir);
        }

        // Centre cross.
        gfx.DrawLine(centre, P(0, R), P(od, R));
        gfx.DrawLine(centre, P(R, 0), P(R, od));

        // OD — vertical dim down the left of the disc (spans y 0..OD).
        DimV(gfx, dimFont, P(0, 0).X, P(0, 0).X - 30, P(0, 0).Y, P(0, od).Y, $"{Fq(od)} {Bi.T("dia")}", true);
        // Donut bore — horizontal dim across the inner circle, below the centre line.
        if (innerD > 0)
        {
            double iy = R - innerD / 2.0;
            DimH(gfx, dimFont, P(R - innerD / 2.0, iy).X, P(R + innerD / 2.0, iy).X, oy + drawSz, oy + drawSz + 22,
                 $"{Fq(innerD)} {Bi.T("dia")}", true);
        }
    }

    // ── Gusset: single flat right-triangle face view (base + height + hypotenuse) ──
    private static void DrawGusset(XGraphics gfx, FlatPatternResult fp, XRect box)
    {
        var titleFont = new XFont("Arial", 9, XFontStyleEx.Bold);
        var dimFont = new XFont("Arial", 8, XFontStyleEx.Bold);
        var cutPen = new XPen(CutColor, 1.2);
        var thin = new XPen(DimColor, 0.6);

        gfx.DrawString(Bi.T("flatPattern.cut"), titleFont, XBrushes.Black,
            new XRect(box.X, box.Y, box.Width, 12), XStringFormats.TopCenter);
        var area = new XRect(box.X, box.Y + 18, box.Width, box.Height - 18);

        double W = fp.FlatWidth, H = fp.FlatHeight;
        if (W <= 0 || H <= 0) return;
        const double band = 52;   // left/bottom bands hold the dim labels without clipping the page
        double availW = area.Width - band * 2, availH = area.Height - band * 2;
        double scale = Math.Min(availW / W, availH / H);
        double drawW = W * scale, drawH = H * scale;
        double ox = area.X + (area.Width - drawW) / 2;
        double oy = area.Y + (area.Height - drawH) / 2;
        XPoint P(double mx, double my) => new(ox + mx * scale, oy + drawH - my * scale);

        // Right angle at (0,0): base along the bottom, height up the left, hypotenuse joins their ends.
        var a = P(0, 0); var b = P(W, 0); var c = P(0, H);
        gfx.DrawPolygon(cutPen, new[] { a, b, c });

        // Right-angle square marker at the (0,0) corner.
        double sq = Math.Min(Math.Min(drawW, drawH) * 0.12, 14);
        gfx.DrawLine(thin, a.X + sq, a.Y, a.X + sq, a.Y - sq);
        gfx.DrawLine(thin, a.X, a.Y - sq, a.X + sq, a.Y - sq);

        // Base (W) along the bottom; height (H) up the left.
        DimH(gfx, dimFont, P(0, 0).X, P(W, 0).X, oy + drawH, oy + drawH + 22, Fq(W), true);
        DimV(gfx, dimFont, ox, ox - 46, oy, oy + drawH, Fq(H), true);

        // Hypotenuse length — boxed callout near the mid-point of the slope.
        double hyp = Math.Sqrt(W * W + H * H);
        var mid = new XPoint((b.X + c.X) / 2, (b.Y + c.Y) / 2);
        string hlabel = Fq(hyp);
        var sz = gfx.MeasureString(hlabel, dimFont);
        double bw = sz.Width + 8, bh = sz.Height + 4;
        var hr = new XRect(mid.X + 4, mid.Y - bh - 2, bw, bh);
        if (hr.X + bw > area.X + area.Width) hr = new XRect(mid.X - bw - 4, mid.Y - bh - 2, bw, bh);
        gfx.DrawRectangle(XBrushes.White, hr);
        gfx.DrawRectangle(new XPen(XColor.FromArgb(110, 110, 110), 0.9), hr);
        gfx.DrawString(hlabel, dimFont, TextBrush, hr, XStringFormats.Center);
    }

    // ── Miter Cut: cross-section (mitered face highlighted) + elevation (both end angles + outside length) ──
    private static void DrawMiterCut(XGraphics gfx, FlatPatternResult fp, XRect box)
    {
        const double rowGap = 10, colGap = 16;
        // The isometric is the PRIMARY view — it's the one a shop worker actually reads for
        // orientation (which end, which way the bevel leans, which face it runs through); the
        // section + elevation below are the secondary, precisely-dimensioned engineering views.
        double topH = box.Height * 0.56, botH = box.Height - topH - rowGap;
        DrawMiterIsometric(gfx, fp, new XRect(box.X, box.Y, box.Width, topH));

        double secY = box.Y + topH + rowGap;
        double wSect = box.Width * 0.30;
        double wElev = box.Width - wSect - colGap;
        DrawMiterSection(gfx, fp, new XRect(box.X, secY, wSect, botH));
        DrawMiterElevation(gfx, fp, new XRect(box.X + wSect + colGap, secY, wElev, botH));
    }

    /// <summary>True when the miter's single bevel-plane taper runs along the cross-section's
    /// "depth" axis (angle's short-leg direction / tube's D) rather than its "width" axis (angle's
    /// long-leg direction / tube's W / flat bar's W) — i.e. which axis the SELECTED miter face is
    /// referenced against. The reference face itself lies at a CONSTANT position on that axis (e.g.
    /// angle's long leg sits at depth≈0 for its whole width), so the visible taper is necessarily
    /// carried by the OTHER, perpendicular axis. Shared by the elevation (which axis gets a real,
    /// dimensioned height) and the isometric (which cross-section vertices taper) so every Miter
    /// view agrees on the same physical cut.</summary>
    private static bool MiterTaperOnDepthAxis(PartSpec s) => s.MiterShape switch
    {
        "angle" => s.MiterFace != 1,                            // long leg (face 0) ref -> tapers with D; short leg (face 1) -> tapers with W
        "tube_square" or "tube_rect" => s.MiterFace is 0 or 2,  // bottom/top ref -> tapers with D; left/right -> tapers with W
        "flatbar" => false,                                     // always the long diagonal across the flat face (tapers with W)
        _ => true,                                              // tube_round: symmetric, axis choice doesn't change the result
    };

    /// <summary>Shape cross-section — angle L, tube outline (square/rect/round), or flat bar — with the
    /// selected miter face highlighted in accent colour + a "MITER FACE" leader. Round tube/pipe and flat
    /// bar have no face concept (symmetric / single face), so no highlight is drawn for them.</summary>
    private static void DrawMiterSection(XGraphics gfx, FlatPatternResult fp, XRect box)
    {
        var s = fp.Spec;
        var titleFont = new XFont("Arial", 9, XFontStyleEx.Bold);
        var labelFont = new XFont("Arial", 7.5, XFontStyleEx.Bold);
        var cutPen = new XPen(CutColor, 1.2);
        var facePen = new XPen(AccentColor, 2.5);

        gfx.DrawString(Bi.T("miter.crossSection"), titleFont, XBrushes.Black,
            new XRect(box.X, box.Y, box.Width, 12), XStringFormats.TopCenter);
        var area = new XRect(box.X, box.Y + 18, box.Width, box.Height - 34);   // leave room for a face callout below

        double W = s.MiterOuterWidth, D = s.MiterShape == "flatbar" ? s.MiterWall : (s.MiterOuterDepth > 0 ? s.MiterOuterDepth : W);
        double wall = s.MiterWall;
        if (W <= 0) return;
        const double band = 40;
        double availW = area.Width - band * 2, availH = area.Height - band * 2;
        double scale = Math.Min(availW / W, availH / D);
        double ox = area.X + (area.Width - W * scale) / 2;
        double oy = area.Y + (area.Height - D * scale) / 2;
        XPoint P(double mx, double my) => new(ox + mx * scale, oy + D * scale - my * scale);

        string? faceCallout = null;
        switch (s.MiterShape)
        {
            case "angle":
            {
                // L-shape: long leg along X (width W), short leg up Y (depth D), outer corner at origin.
                var pts = new[]
                {
                    P(0, 0), P(W, 0), P(W, wall), P(wall, wall), P(wall, D), P(0, D),
                };
                gfx.DrawPolygon(cutPen, pts);
                if (s.MiterFace == 1)   // short leg
                { gfx.DrawLine(facePen, P(0, 0), P(0, D)); faceCallout = Bi.T("miter.face"); }
                else                    // long leg (default)
                { gfx.DrawLine(facePen, P(0, 0), P(W, 0)); faceCallout = Bi.T("miter.face"); }
                break;
            }
            case "tube_square":
            case "tube_rect":
            {
                gfx.DrawRectangle(cutPen, P(0, D).X, P(0, D).Y, W * scale, D * scale);
                double iw = Math.Max(0, W - 2 * wall), id = Math.Max(0, D - 2 * wall);
                if (iw > 0 && id > 0)
                    gfx.DrawRectangle(cutPen, P(wall, D - wall).X, P(wall, D - wall).Y, iw * scale, id * scale);
                // 4 faces, 0-based: 0=bottom, 1=right, 2=top, 3=left.
                (XPoint a, XPoint b) edge = s.MiterFace switch
                {
                    1 => (P(W, 0), P(W, D)),
                    2 => (P(0, D), P(W, D)),
                    3 => (P(0, 0), P(0, D)),
                    _ => (P(0, 0), P(W, 0)),
                };
                gfx.DrawLine(facePen, edge.a, edge.b);
                faceCallout = Bi.T("miter.face");
                break;
            }
            case "tube_round":
            {
                double r = W / 2 * scale;
                var c = P(W / 2, D / 2);
                gfx.DrawEllipse(cutPen, c.X - r, c.Y - r, 2 * r, 2 * r);
                double ir = Math.Max(0, W / 2 - wall) * scale;
                if (ir > 0) gfx.DrawEllipse(cutPen, c.X - ir, c.Y - ir, 2 * ir, 2 * ir);
                break;   // symmetric — no face to highlight
            }
            default:   // flatbar
                gfx.DrawRectangle(cutPen, P(0, D).X, P(0, D).Y, W * scale, D * scale);
                break;   // one meaningful face (through the width) — not worth a callout
        }

        // Dimensions along the bottom + left.
        var dimFont = new XFont("Arial", 8, XFontStyleEx.Bold);
        DimH(gfx, dimFont, P(0, 0).X, P(W, 0).X, P(0, 0).Y, oy + D * scale + 20, Fq(W), false);
        if (s.MiterShape != "tube_round") DimV(gfx, dimFont, P(0, 0).X, ox - 32, P(0, D).Y, P(0, 0).Y, Fq(D), false);

        if (faceCallout is not null)
        {
            var f = new XFont("Arial", 8, XFontStyleEx.Bold);
            gfx.DrawString(faceCallout, f, AccentBrush,
                new XRect(box.X, box.Y + box.Height - 14, box.Width, 12), XStringFormats.TopCenter);
        }
    }

    /// <summary>Elevation/profile view — a schematic side silhouette of the part with both ends cut at
    /// their specified angle (measured from the length axis; 90° would be a square cut, smaller angles cut
    /// more aggressively), dimensioned with the OUTSIDE (long-point) length and both end angles.</summary>
    private static void DrawMiterElevation(XGraphics gfx, FlatPatternResult fp, XRect box)
    {
        var s = fp.Spec;
        var titleFont = new XFont("Arial", 9, XFontStyleEx.Bold);
        var dimFont = new XFont("Arial", 8, XFontStyleEx.Bold);
        var cutPen = new XPen(CutColor, 1.2);

        gfx.DrawString(Bi.T("miter.elevation"), titleFont, XBrushes.Black,
            new XRect(box.X, box.Y, box.Width, 12), XStringFormats.TopCenter);
        var area = new XRect(box.X, box.Y + 18, box.Width, box.Height - 18);

        double len = s.Length;
        // Real cross-section dimension, not a schematic stand-in — whichever axis the bevel actually
        // tapers along (see MiterTaperOnDepthAxis), so this panel is self-sufficient: the shop reads
        // length, both end angles, AND the flange/leg/OD size without cross-referencing the section view.
        bool taperOnDepth = MiterTaperOnDepthAxis(s);
        double hgt = taperOnDepth ? s.MiterOuterDepth : s.MiterOuterWidth;
        if (len <= 0 || hgt <= 0) return;

        const double band = 50;
        double availW = area.Width - band * 2, availH = area.Height - band * 2 - 20;
        double scale = Math.Min(availW / len, availH / hgt);
        double drawH = hgt * scale;
        double ox = area.X + (area.Width - len * scale) / 2;
        double oy = area.Y + (area.Height - drawH) / 2;
        XPoint P(double mx, double my) => new(ox + mx * scale, oy + drawH - my * scale);

        double insetL = hgt / Math.Tan(s.MiterAngleLeft * Math.PI / 180.0);
        double insetR = hgt / Math.Tan(s.MiterAngleRight * Math.PI / 180.0);
        insetL = Math.Clamp(insetL, 0, len / 2 - 0.01);
        insetR = Math.Clamp(insetR, 0, len / 2 - 0.01);

        var pts = new[] { P(0, 0), P(len, 0), P(len - insetR, hgt), P(insetL, hgt) };
        gfx.DrawPolygon(cutPen, pts);

        // Outside (long-point) length — the as-specified accent dimension.
        DimH(gfx, dimFont, P(0, 0).X, P(len, 0).X, P(0, 0).Y, oy + drawH + 22, Fq(len), true);

        // Real flange/leg/OD dimension (the axis the bevel tapers along) — same value + styling as
        // the matching dimension on the cross-section view. The "A" prefix matches the isometric's
        // own flange/leg/OD label (same value, same axis) so the shop can trace it between views.
        DimV(gfx, dimFont, P(0, 0).X, ox - 30, P(0, hgt).Y, P(0, 0).Y, $"A {Fq(hgt)}", false);

        // End-angle callouts near each slanted edge, pushed clear of the material along that edge's own
        // outward normal (away from the shape's centroid), with a short witness tick — never sitting on
        // top of the drawn trapezoid. Degrees read as a short decimal (0.# — "45°", "45.5°"), not F()'s
        // fixed 3-decimal money-style formatting ("45.000°"), which reads oddly for an angle. "B"/"C"
        // prefixes match the isometric's END A / END B angle callouts.
        string Deg(double v) => v.ToString("0.#", CultureInfo.InvariantCulture);
        var angFont = new XFont("Arial", 8, XFontStyleEx.Bold);
        var elevCentroid = new XPoint(pts.Average(p => p.X), pts.Average(p => p.Y));
        var witnessPen = new XPen(DimColor, 0.6);
        void SlantLabel(XPoint a, XPoint b, string text)
        {
            double ex = b.X - a.X, ey = b.Y - a.Y, el = Math.Sqrt(ex * ex + ey * ey); if (el < 1e-6) el = 1;
            double nx = -ey / el, ny = ex / el;
            double mx = (a.X + b.X) / 2, my = (a.Y + b.Y) / 2;
            if (nx * (mx - elevCentroid.X) + ny * (my - elevCentroid.Y) < 0) { nx = -nx; ny = -ny; }
            var w1 = new XPoint(mx, my);
            var w2 = new XPoint(mx + nx * 22, my + ny * 22);
            gfx.DrawLine(witnessPen, w1, w2);
            gfx.DrawString(text, angFont, TextBrush, new XRect(w2.X - 24, w2.Y - 6, 48, 12), XStringFormats.Center);
        }
        SlantLabel(P(0, 0), P(insetL, hgt), $"B {Deg(s.MiterAngleLeft)}°");
        SlantLabel(P(len, 0), P(len - insetR, hgt), $"C {Deg(s.MiterAngleRight)}°");
    }

    /// <summary>Two independent isometric "bevel detail" sketches, one per cut end, so the shop sees
    /// the TRUE 3D shape of the part — which leg/face the bevel runs through and which way each end
    /// leans — instead of inferring it from the flat cross-section + elevation. ONE continuous stick
    /// (both ends of the SAME solid, not two separate sketches — two disconnected pictures read as
    /// "which one is really the part?" and invite exactly the confusion this view exists to prevent).
    /// Drawn at a REPRESENTATIVE length, not true length (96" of part vs 2-4" of cross-section would
    /// make either end unreadable at one true scale — the true length lives on the elevation, and
    /// every drawing already carries a blanket "NOT TO SCALE" footer). Same single-plane bevel model
    /// as the cross-section's face highlight and the elevation's dimensions/angles
    /// (<see cref="MiterTaperOnDepthAxis"/>), and the SAME lettered tags (A/B/C) as the elevation, so
    /// every view agrees on which physical edge is which. Same true-isometric convention as
    /// <see cref="DrawPanIso"/> (x/y recede at 30°, z is vertical).</summary>
    private static void DrawMiterIsometric(XGraphics gfx, FlatPatternResult fp, XRect box)
    {
        var s = fp.Spec;
        var titleFont = new XFont("Arial", 9, XFontStyleEx.Bold);
        var angFont = new XFont("Arial", 9, XFontStyleEx.Bold);
        var dimFontIso = new XFont("Arial", 7.5, XFontStyleEx.Bold);
        var cutPen = new XPen(CutColor, 1.1);
        var facePen = new XPen(AccentColor, 2.2);
        var cutFaceFill = new XSolidBrush(XColor.FromArgb(235, 235, 240));
        var faceFill = new XSolidBrush(XColor.FromArgb(60, AccentColor.R, AccentColor.G, AccentColor.B));

        gfx.DrawString(Bi.T("miter.isometric"), titleFont, XBrushes.Black,
            new XRect(box.X, box.Y, box.Width, 12), XStringFormats.TopCenter);
        var area = new XRect(box.X, box.Y + 16, box.Width, box.Height - 30);   // room for END A/B + angle labels below
        if (s.MiterOuterWidth <= 0) return;

        var poly = MiterCrossSectionPoints(s);   // local (v, w): v = width/long-leg axis, w = depth/short-leg axis
        if (poly.Count < 3) return;
        var refEdge = MiterFaceEdge(s);
        bool taperOnDepth = MiterTaperOnDepthAxis(s);

        double vExt = poly.Max(p => p.v), wExt = poly.Max(p => p.w);
        double maxExt = Math.Max(vExt, wExt);
        if (maxExt <= 0) return;

        // A representative TOTAL length for the whole stick — proportioned to the cross-section, NOT
        // the true part length. Each end's bevel is capped at 42% of it, always leaving a visible
        // straight middle span (the "rest of the material") even at extreme angles.
        double Lrep = maxExt * 5.5;
        double capEach = Lrep * 0.42;
        double kA = 1.0 / Math.Tan(Math.Clamp(s.MiterAngleLeft, 3, 87) * Math.PI / 180.0);
        double kB = 1.0 / Math.Tan(Math.Clamp(s.MiterAngleRight, 3, 87) * Math.PI / 180.0);
        double UA((double v, double w) p) => Math.Min((taperOnDepth ? p.w : p.v) * kA, capEach);
        double UB((double v, double w) p) => Lrep - Math.Min((taperOnDepth ? p.w : p.v) * kB, capEach);

        // True isometric (same convention as the Pan formed-part view): x/y both recede at 30°, z is
        // vertical. u = length (left = End A, right = End B), v = cross-section width axis, w =
        // cross-section depth axis.
        const double c30 = 0.8660254, s30 = 0.5;
        (double px, double py) Iso(double u, double v, double w) => ((u - v) * c30, (u + v) * s30 - w);

        XPoint RingA(int i) { var (v, w) = poly[i]; var (px, py) = Iso(UA(poly[i]), v, w); return new XPoint(px, py); }
        XPoint RingB(int i) { var (v, w) = poly[i]; var (px, py) = Iso(UB(poly[i]), v, w); return new XPoint(px, py); }

        double minX = double.MaxValue, maxX = double.MinValue, minY = double.MaxValue, maxY = double.MinValue;
        void Acc(XPoint p) { if (p.X < minX) minX = p.X; if (p.X > maxX) maxX = p.X; if (p.Y < minY) minY = p.Y; if (p.Y > maxY) maxY = p.Y; }
        for (int i = 0; i < poly.Count; i++) { Acc(RingA(i)); Acc(RingB(i)); }
        double gw = Math.Max(maxX - minX, 1e-6), gh = Math.Max(maxY - minY, 1e-6);

        const double pad = 60;   // room for the END-label leaders + dimension witness lines pushed outside the shape
        double scale = Math.Min((area.Width - pad * 2) / gw, (area.Height - pad * 2) / gh);
        double ox = area.X + (area.Width - gw * scale) / 2 - minX * scale;
        double oy = area.Y + (area.Height - gh * scale) / 2 - minY * scale;
        XPoint S(XPoint p) => new(ox + p.X * scale, oy + p.Y * scale);

        var aPts = Enumerable.Range(0, poly.Count).Select(i => S(RingA(i))).ToArray();
        var bPts = Enumerable.Range(0, poly.Count).Select(i => S(RingB(i))).ToArray();

        // Longitudinal edges connecting the two cut faces — the constant-cross-section material
        // between them (drawn first so the two end faces layer cleanly on top).
        for (int i = 0; i < poly.Count; i++)
            gfx.DrawLine(cutPen, aPts[i], bPts[i]);

        // The two cut faces themselves — the actual exposed miter planes — filled so each reads as a
        // solid surface, not just an outline.
        gfx.DrawPolygon(cutFaceFill, aPts, XFillMode.Winding);
        gfx.DrawPolygon(cutPen, aPts);
        gfx.DrawPolygon(cutFaceFill, bPts, XFillMode.Winding);
        gfx.DrawPolygon(cutPen, bPts);

        // Highlight the SAME reference face the cross-section view calls "MITER FACE", on BOTH ends
        // (one physical face, shared the whole length) — identical accent colour, so a shop worker
        // matches this 3D picture to that 2D callout on sight.
        if (refEdge.HasValue)
        {
            var (r0, r1) = refEdge.Value;
            gfx.DrawPolygon(faceFill, new[] { aPts[r0], aPts[r1], bPts[r1], bPts[r0] }, XFillMode.Winding);
            gfx.DrawLine(facePen, aPts[r0], aPts[r1]);
            gfx.DrawLine(facePen, bPts[r0], bPts[r1]);
            gfx.DrawLine(facePen, aPts[r0], bPts[r0]);
            gfx.DrawLine(facePen, aPts[r1], bPts[r1]);
        }

        // END A / END B + angle labels — anchored to each end's own OUTERMOST corner (the true tip,
        // not a generic average position), then pushed further out along that same direction with a
        // thin witness leader, so the callout sits clearly outside the drawn material and still points
        // straight at the corner it's describing. "B"/"C" prefixes match the elevation's own End A /
        // End B angle callouts.
        string Deg(double v) => v.ToString("0.#", CultureInfo.InvariantCulture);
        double acx = aPts.Average(p => p.X), acy = aPts.Average(p => p.Y);
        double bcx = bPts.Average(p => p.X), bcy = bPts.Average(p => p.Y);
        double outAx = acx - bcx, outAy = acy - bcy;   // "away from End B" = outward for End A
        double outAl = Math.Sqrt(outAx * outAx + outAy * outAy); if (outAl < 1e-6) outAl = 1;
        outAx /= outAl; outAy /= outAl;
        var dimPen = new XPen(DimColor, 0.6);

        void EndLabel(XPoint[] ring, double ox2, double oy2, string endName, string angleText)
        {
            int tipI = 0; double best = double.MinValue;
            for (int i = 0; i < ring.Length; i++)
            {
                double proj = ring[i].X * ox2 + ring[i].Y * oy2;
                if (proj > best) { best = proj; tipI = i; }
            }
            var tip = ring[tipI];
            var anchor = new XPoint(tip.X + ox2 * 34, tip.Y + oy2 * 34);
            gfx.DrawLine(dimPen, tip, anchor);
            var box2 = new XRect(anchor.X - 40, anchor.Y - 12, 80, 24);
            gfx.DrawString(endName, angFont, TextBrush, new XRect(box2.X, box2.Y, box2.Width, 11), XStringFormats.TopCenter);
            gfx.DrawString(angleText, angFont, AccentBrush, new XRect(box2.X, box2.Y + 12, box2.Width, 12), XStringFormats.TopCenter);
        }
        EndLabel(aPts, outAx, outAy, "END A", $"B {Deg(s.MiterAngleLeft)}°");
        EndLabel(bPts, -outAx, -outAy, "END B", $"C {Deg(s.MiterAngleRight)}°");

        // Compact cross-section dimension labels — a quick visual size check (the cross-section view
        // carries the precise, fully-dimensioned callouts). The "A" prefix matches the elevation's
        // flange/leg/OD dimension exactly (same axis, same value); the other extent has no elevation
        // counterpart to cross-reference, so it's shown plain, un-prefixed. Offset along the EDGE's own
        // outward normal (not a distance-from-centroid heuristic, which under-clears a long thin part),
        // with a short witness tick from the material edge to the value.
        double cx = aPts.Average(p => p.X), cy = aPts.Average(p => p.Y);
        void EdgeLabel(XPoint a, XPoint b, string text)
        {
            double ex = b.X - a.X, ey = b.Y - a.Y, el = Math.Sqrt(ex * ex + ey * ey); if (el < 1e-6) el = 1;
            double nx = -ey / el, ny = ex / el;
            double mx = (a.X + b.X) / 2, my = (a.Y + b.Y) / 2;
            if (nx * (mx - cx) + ny * (my - cy) < 0) { nx = -nx; ny = -ny; }
            var w1 = new XPoint(mx, my);
            var w2 = new XPoint(mx + nx * 26, my + ny * 26);
            gfx.DrawLine(dimPen, w1, w2);
            gfx.DrawString(text, dimFontIso, DimBrush, new XRect(w2.X - 24, w2.Y - 5, 48, 10), XStringFormats.Center);
        }
        if (s.MiterShape == "tube_round")
        {
            // Round has no flat edge to offset from — lead off the ring's topmost point instead (same
            // "extreme point, then push further out" technique as the END A/B labels).
            int topI = 0; double bestY = double.MaxValue;
            for (int i = 0; i < aPts.Length; i++) if (aPts[i].Y < bestY) { bestY = aPts[i].Y; topI = i; }
            var top = aPts[topI];
            double dxr = top.X - cx, dyr = top.Y - cy;
            double dlr = Math.Sqrt(dxr * dxr + dyr * dyr); if (dlr < 1e-6) { dxr = 0; dyr = -1; dlr = 1; }
            var w2r = new XPoint(top.X + dxr / dlr * 26, top.Y + dyr / dlr * 26);
            gfx.DrawLine(dimPen, top, w2r);
            gfx.DrawString($"A {Fq(vExt)} OD", dimFontIso, DimBrush, new XRect(w2r.X - 30, w2r.Y - 5, 60, 10), XStringFormats.Center);
        }
        else
        {
            // Elevation's "A" = hgt = taperOnDepth ? wExt(D) : vExt(W) — match that EXACT axis here, not the
            // edge that happens to share the same w/v-constant condition (that flip is the actual bug this fixes).
            EdgeLabel(aPts[0], aPts[1], taperOnDepth ? Fq(vExt) : $"A {Fq(vExt)}");            // w=0 edge — spans vExt
            EdgeLabel(aPts[poly.Count - 1], aPts[0], taperOnDepth ? $"A {Fq(wExt)}" : Fq(wExt));   // v=0 edge — spans wExt
        }
    }

    /// <summary>Cross-section boundary in LOCAL (v, w) coordinates — v = width/long-leg axis, w =
    /// depth/short-leg axis — same shape conventions as <see cref="DrawMiterSection"/>'s 2D polygons,
    /// reused here for the isometric's 3D extrusion. Round tube/pipe is approximated as a 28-gon.</summary>
    private static List<(double v, double w)> MiterCrossSectionPoints(PartSpec s)
    {
        double W = s.MiterOuterWidth, D = s.MiterOuterDepth > 0 ? s.MiterOuterDepth : W, wall = s.MiterWall;
        switch (s.MiterShape)
        {
            case "angle":
                return new() { (0, 0), (W, 0), (W, wall), (wall, wall), (wall, D), (0, D) };
            case "tube_round":
            {
                var pts = new List<(double, double)>();
                double r = W / 2;
                const int n = 28;
                for (int i = 0; i < n; i++)
                {
                    double th = 2 * Math.PI * i / n;
                    pts.Add((r + r * Math.Cos(th), r + r * Math.Sin(th)));
                }
                return pts;
            }
            case "flatbar":
                return new() { (0, 0), (W, 0), (W, wall), (0, wall) };
            default:   // tube_square, tube_rect
                return new() { (0, 0), (W, 0), (W, D), (0, D) };
        }
    }

    /// <summary>Which cross-section polygon edge (vertex-index pair, matching
    /// <see cref="MiterCrossSectionPoints"/>) is the selected miter face — same edges
    /// <see cref="DrawMiterSection"/> highlights in 2D. Null for round tube/pipe and flat bar (no
    /// face concept — symmetric / single meaningful face).</summary>
    private static (int, int)? MiterFaceEdge(PartSpec s) => s.MiterShape switch
    {
        "angle" => s.MiterFace == 1 ? (5, 0) : (0, 1),
        "tube_square" or "tube_rect" => s.MiterFace switch { 1 => (1, 2), 2 => (2, 3), 3 => (3, 0), _ => (0, 1) },
        _ => null,
    };

    // ── Pan: single flat-pattern top view (cut outline + bend lines + corner reliefs) ──
    private static void DrawPan(XGraphics gfx, FlatPatternResult fp, XRect box)
    {
        var titleFont = new XFont("Arial", 9, XFontStyleEx.Bold);
        var dimFont = new XFont("Arial", 8, XFontStyleEx.Bold);
        var cutPen = new XPen(CutColor, 1.2);
        var bendPen = new XPen(BendColor, 0.9) { DashStyle = XDashStyle.Dash };

        gfx.DrawString(Bi.T("flatPattern.cut"), titleFont, XBrushes.Black,
            new XRect(box.X, box.Y, box.Width, 12), XStringFormats.TopCenter);
        var area = new XRect(box.X, box.Y + 18, box.Width, box.Height - 18);

        // Bounds from the actual cut geometry (returns push vertices beyond [0,FlatWidth]).
        double minX = double.MaxValue, maxX = double.MinValue, minY = double.MaxValue, maxY = double.MinValue;
        void Acc(double x, double y) { if (x < minX) minX = x; if (x > maxX) maxX = x; if (y < minY) minY = y; if (y > maxY) maxY = y; }
        foreach (var e in fp.Cut.Entities)
        {
            if (e.Layer == PartLabel.LayerName) continue;   // no-cut DXF label layer — never sized/drawn on the PDF
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
            if (e.Layer == PartLabel.LayerName) continue;   // no-cut DXF label layer — never drawn on the PDF
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
        // Value only on the flat pattern (the basis word + Adentro/Afuera is carried on the sections),
        // so the leftmost label can't run off the page edge.
        DimH(gfx, dimFont, P(bx0, by0).X, P(bx1, by0).X, P(bx0, by0).Y, oy + drawH + 24, Fq(bx1 - bx0), true);
        DimV(gfx, dimFont, P(bx0, by0).X, ox - 46, P(bx0, by0).Y, P(bx0, by1).Y, Fq(by1 - by0), true);
        if (s.PanBottom && by0 > 0)
            DimV(gfx, dimFont, P(bx1, 0).X, P(bx1, 0).X + 26, P(bx1, 0).Y, P(bx1, by0).Y, Fq(fp.PanWallDev), false);
        else if (s.PanLeft && bx0 > 0)
            DimH(gfx, dimFont, P(0, by0).X, P(bx0, by0).X, P(0, by0).Y, oy + drawH + 24, Fq(fp.PanWallDev), false);

        // Finish callout — boxed label on the base (pans show inside/outside on the base face).
        if (s.Finish is FinishSide.Outside or FinishSide.Inside)
        {
            var f = new XFont("Arial", 9, XFontStyleEx.Bold);
            string txt = s.Finish == FinishSide.Outside ? Bi.T("finish.outside") : Bi.T("finish.inside");
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

        gfx.DrawString(Bi.T("formedPart"), titleFont, XBrushes.Black,
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
        List<(double x, double y)> prof, double webOD, double wallOD,
        double webSpec, DimBasis webBasis, double wallSpec, DimBasis wallBasis, double? fixedScale = null)
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
        DimH(gfx, dimFont, P(0, 0).X, P(webOD, 0).X, P(0, 0).Y, sBottom + 20, $"{Fq(webSpec)} {Bi.Basis(webBasis)}", true);
        DimV(gfx, dimFont, P(0, 0).X, sLeft - 22, P(0, 0).Y, P(0, wallOD).Y, $"{Fq(wallSpec)} {Bi.Basis(wallBasis)}", true);
        // (thickness is shown in the header details, not called out on the geometry)
    }

    // ── Footnote box ─────────────────────────────────────────────────────────
    // Bilingual spec table under the header: "English / Spanish" attribute label + value column.
    // Rows vary by part type; Material + Flat-blank rows are always present. Returns the table's
    // bottom Y so the drawing panels start below it.
    private static double DrawSpecTable(XGraphics gfx, FlatPatternResult fp, double x, double y, double maxWidth)
    {
        var s = fp.Spec;
        string Frac(double v) => DrawFormat.FracInch(v);
        string thk = DrawFormat.ThicknessLabel(s.Material, s.Thickness);

        // Just the three at-a-glance attributes — length, thickness, flat blank. Everything else
        // (web/flanges/legs/section/holes/material) is already labelled on the drawing itself.
        var rows = new List<(string Label, string Value)>();

        // Quantity — always first, every drawing type.
        rows.Add((Bi.T("spec.quantity"), fp.Qty.ToString()));

        // Length (or a column's overall height); paddle blinds / circles / sheets have no single "length".
        if (fp.IsColumn)                                          rows.Add((Bi.T("spec.height"), Frac(s.ColumnFullHeight)));
        else if (!fp.IsPaddle && !fp.IsCircle && s.Length > 0)    rows.Add((Bi.T("spec.length"), Frac(s.Length)));

        // Thickness — single-sheet parts; a column carries several (its plates/wall are on the drawing).
        if (!fp.IsColumn) rows.Add((Bi.T("thickness"), thk));

        // Flat blank / cut-to — cut dimensions in decimal inches. Miter Cut has no flat blank (no DXF) —
        // the mitered face + both end angles are called out on the drawing itself instead.
        if (!fp.IsMiterCut)
        {
            string Dec(double v) => DrawFormat.DecInch(v);
            string blankLabel = fp.IsColumn ? Bi.T("cutTo") : fp.IsPaddle ? Bi.T("plateToCut") : Bi.T("flatBlank");
            string blankValue = fp.IsColumn ? $"tube {Dec(fp.ColumnTubeLength)} + plates"
                : fp.IsCircle ? (s.InnerDiameter > 0 ? $"{thk} x Ø{Dec(s.Diameter)} / Ø{Dec(s.InnerDiameter)}" : $"{thk} x Ø{Dec(s.Diameter)}")
                : (fp.IsPlate || fp.IsPaddle) ? $"{thk} x {Dec(fp.FlatHeight)} x {Dec(fp.FlatWidth)}"
                : $"{Dec(fp.FlatWidth)} x {Dec(fp.FlatHeight)}";
            rows.Add((blankLabel, blankValue));
        }

        // ── layout ──
        var labelFont = new XFont("Arial", 8.5, XFontStyleEx.Regular);
        var valueFont = new XFont("Arial", 8.5, XFontStyleEx.Bold);
        var labelBrush = new XSolidBrush(XColor.FromArgb(90, 90, 90));
        var rulePen = new XPen(XColor.FromArgb(210, 210, 210), 0.6);
        const double labelColW = 150, pad = 5, rowH = 13;
        double valueColW = Math.Min(240, Math.Max(120, maxWidth - labelColW));
        double tableW = labelColW + valueColW;

        double ry = y;
        gfx.DrawLine(rulePen, x, ry, x + tableW, ry);                       // top rule
        foreach (var (label, value) in rows)
        {
            string[] vlines = { value };
            if (gfx.MeasureString(value, valueFont).Width > valueColW - pad && value.Contains(" / "))
            {
                int i = value.IndexOf(" / ", StringComparison.Ordinal);
                vlines = new[] { value.Substring(0, i + 2), value.Substring(i + 3) };
            }
            gfx.DrawString(label, labelFont, labelBrush,
                new XRect(x + pad, ry + 2, labelColW - pad, rowH), XStringFormats.TopLeft);
            for (int li = 0; li < vlines.Length; li++)
                gfx.DrawString(vlines[li], valueFont, XBrushes.Black,
                    new XRect(x + labelColW + pad, ry + 2 + li * rowH, valueColW - pad, rowH), XStringFormats.TopLeft);
            ry += rowH * vlines.Length;
            gfx.DrawLine(rulePen, x, ry, x + tableW, ry);                   // row rule
        }
        gfx.DrawLine(rulePen, x, y, x, ry);                                 // left border
        gfx.DrawLine(rulePen, x + labelColW, y, x + labelColW, ry);         // column separator
        gfx.DrawLine(rulePen, x + tableW, y, x + tableW, ry);              // right border
        return ry;
    }

    /// <summary>Path to the franchise logo (env vars expanded) shown in the drawing info box; blank = none.</summary>
    // The logo lives in the synced SharePoint library (same %USERPROFILE%\Mithril Metals Corp\… root the
    // publish + secrets use). %OneDriveCommercial% points at the *personal* OneDrive, which doesn't hold it.
    public static string? LogoPath { get; set; } = @"%USERPROFILE%\Mithril Metals Corp\Metal Supermarkets Hackensack - Documents\MS_Primary_Logo_C100_M44_Flattened_Preview.png";

    // Right 1/3 of the footnote band: Created By / Created Date / NOT TO SCALE, then the franchise logo.
    private static void DrawInfoBox(XGraphics gfx, XRect box, string? createdBy = null)
    {
        gfx.DrawRectangle(new XPen(XColor.FromArgb(205, 205, 205), 0.8), box);
        var clip = gfx.Save();
        gfx.IntersectClip(box);
        var lblFont = new XFont("Arial", 7.5, XFontStyleEx.Regular);
        var valFont = new XFont("Arial", 7.5, XFontStyleEx.Bold);
        var lblBrush = new XSolidBrush(XColor.FromArgb(90, 90, 90));
        double x = box.X + 6, y = box.Y + 4, labelW = 56, innerW = box.Width - 12;

        void Row(string label, string val)
        {
            gfx.DrawString(label, lblFont, lblBrush, new XRect(x, y, labelW, 11), XStringFormats.TopLeft);
            gfx.DrawString(val, valFont, XBrushes.Black, new XRect(x + labelW, y, innerW - labelW, 11), XStringFormats.TopLeft);
            y += 11;
        }
        Row("Created By:", Capitalize(createdBy is { Length: > 0 } cb ? cb : Environment.UserName));
        Row("Created Date:", DateTime.Now.ToString("MMM d, yyyy", CultureInfo.InvariantCulture));
        gfx.DrawString("NOT TO SCALE", valFont, XBrushes.Black, new XRect(x, y, innerW, 11), XStringFormats.TopLeft);
        y += 13;

        // Franchise logo anchored bottom-right of the box. Resolve across known locations (the configured
        // path can point at a OneDrive root that doesn't hold the file); skipped silently if none decode.
        string? path = new[]
        {
            LogoPath,
            @"%USERPROFILE%\Mithril Metals Corp\Metal Supermarkets Hackensack - Documents\MS_Primary_Logo_C100_M44_Flattened_Preview.png",
            @"%OneDriveCommercial%\MS_Primary_Logo_C100_M44_Flattened_Preview.png",
            // Reliable bundled fallback — deployed next to the proxy exe; works without OneDrive.
            Path.Combine(AppContext.BaseDirectory, "Assets", "ms-primary-logo.png"),
        }
        .Where(c => !string.IsNullOrWhiteSpace(c))
        .Select(c => Environment.ExpandEnvironmentVariables(c!))
        .FirstOrDefault(File.Exists);

        if (path is not null)
        {
            try
            {
                var img = XImage.FromFile(path);
                double availW = innerW, availH = box.Y + box.Height - y - 4;
                if (availW > 6 && availH > 6)
                {
                    double sc = Math.Min(availW / img.PixelWidth, availH / img.PixelHeight);
                    double w = img.PixelWidth * sc, h = img.PixelHeight * sc;
                    gfx.DrawImage(img, box.X + box.Width - w - 5, box.Y + box.Height - h - 4, w, h);
                }
            }
            catch { /* missing / unsupported logo image — keep the info text */ }
        }
        gfx.Restore(clip);
    }

    private static string Capitalize(string s) => string.IsNullOrEmpty(s) ? s : char.ToUpper(s[0]) + s[1..];

    private static void DrawFootnote(XGraphics gfx, FlatPatternResult fp, XRect box)
    {
        gfx.DrawRectangle(new XPen(XColor.FromArgb(205, 205, 205), 0.8), box);
        var clip = gfx.Save();
        gfx.IntersectClip(box);                       // keep all text inside the 2/3 box — no spill-over
        var font = new XFont("Arial", 8, XFontStyleEx.Regular);
        double y = box.Y + 3, innerW = box.Width - 16;
        foreach (var line in fp.Summary.Split('\n'))
        {
            gfx.DrawString(line, font, XBrushes.Black, new XRect(box.X + 8, y, innerW, 10), XStringFormats.TopLeft);
            y += 10;
        }
        // Paddle blinds, columns, and miter cuts have no bends — drop the bend/Ri legend bits that don't apply.
        string legend = fp.IsPaddle || fp.IsColumn || fp.IsMiterCut
            ? $"{Bi.T("legend.solidCut")}  |  {Bi.T("legend.boldSpec")}  |  {Bi.T("legend.fracInches")}"
            : $"{Bi.T("legend.solidCut")}  |  {Bi.T("legend.dashedBend")}  |  {Bi.T("legend.boldSpec")}  |  {Bi.T("legend.insideRadius")} {F(fp.Spec.InsideRadius)}\"  |  {Bi.T("legend.fracInches")}";
        var legFont = new XFont("Arial", 6, XFontStyleEx.Regular);
        foreach (var wline in WrapBySegments(gfx, legend, "  |  ", legFont, innerW))
        {
            gfx.DrawString(wline, legFont, DimBrush, new XRect(box.X + 8, y, innerW, 9), XStringFormats.TopLeft);
            y += 8;
        }
        gfx.Restore(clip);
    }

    // Greedily pack segment-separated text into lines no wider than maxW (keeps the bilingual legend
    // inside its box instead of overflowing).
    private static List<string> WrapBySegments(XGraphics gfx, string text, string sep, XFont font, double maxW)
    {
        var lines = new List<string>();
        string cur = "";
        foreach (var seg in text.Split(new[] { sep }, StringSplitOptions.None))
        {
            string trial = cur.Length == 0 ? seg : cur + sep + seg;
            if (cur.Length > 0 && gfx.MeasureString(trial, font).Width > maxW) { lines.Add(cur); cur = seg; }
            else cur = trial;
        }
        if (cur.Length > 0) lines.Add(cur);
        return lines;
    }

    // ── Key for the formed-part section-cut planes (Side = dashed, End = dash-dot) ──────────────
    private static void DrawSectionKey(XGraphics gfx, XRect box, bool isPan)
    {
        var kf = new XFont("Arial", 7, XFontStyleEx.Regular);
        var brushA = new XSolidBrush(SecColorA);
        var brushB = new XSolidBrush(SecColorB);
        gfx.DrawRectangle(new XPen(XColor.FromArgb(205, 205, 205), 0.8), box);
        gfx.DrawString(Bi.T("sectionCuts"), new XFont("Arial", 7, XFontStyleEx.Bold), XBrushes.Black,
            new XRect(box.X + 4, box.Y + 2, box.Width - 8, 9), XStringFormats.TopLeft);
        double ly = box.Y + 19, x = box.X + 6;
        // Pan = Side (green dash) + End (orange dash-dot). (Single-section views no longer draw a key.)
        if (isPan)
        {
            gfx.DrawLine(new XPen(SecColorA, 1.4) { DashStyle = XDashStyle.Dash }, new XPoint(x, ly), new XPoint(x + 22, ly));
            gfx.DrawString(Bi.T("side"), kf, brushA, new XRect(x + 25, ly - 5, 90, 9), XStringFormats.TopLeft);
            x += 120;
            gfx.DrawLine(new XPen(SecColorB, 1.4) { DashStyle = XDashStyle.DashDot }, new XPoint(x, ly), new XPoint(x + 22, ly));
            gfx.DrawString(Bi.T("end"), kf, brushB, new XRect(x + 25, ly - 5, 90, 9), XStringFormats.TopLeft);
        }
        else
        {
            gfx.DrawLine(new XPen(SecColorA, 1.4) { DashStyle = XDashStyle.DashDot }, new XPoint(x, ly), new XPoint(x + 22, ly));
            gfx.DrawString(Bi.T("end"), kf, brushA, new XRect(x + 25, ly - 5, 90, 9), XStringFormats.TopLeft);
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

    // accent = the as-specified dimension: bold label + accent-coloured leader/arrows. Otherwise the
    // muted gray derived-dimension style. (Witness lines stay perpendicular to the face, dim line
    // parallel to it.)
    private static void DimH(XGraphics gfx, XFont font, double x1, double x2,
        double faceY, double dimY, string label, bool accent)
    {
        var col = accent ? AccentColor : DimColor;
        var brush = accent ? AccentBrush : DimBrush;
        var lblFont = accent ? new XFont("Arial", font.Size, XFontStyleEx.Bold) : font;
        var lblBrush = accent ? AccentBrush : TextBrush;
        var pen = new XPen(col, accent ? 0.9 : 0.6);
        double dir = Math.Sign(dimY - faceY); if (dir == 0) dir = 1;
        gfx.DrawLine(pen, x1, faceY + ExtGap * dir, x1, dimY + ExtOver * dir);
        gfx.DrawLine(pen, x2, faceY + ExtGap * dir, x2, dimY + ExtOver * dir);
        gfx.DrawLine(pen, x1, dimY, x2, dimY);
        Arrow(gfx, new XPoint(x1, dimY), -1, 0, brush);
        Arrow(gfx, new XPoint(x2, dimY), 1, 0, brush);
        double ty = dir > 0 ? dimY + 2 : dimY - 12;
        gfx.DrawString(label, lblFont, lblBrush, new XRect((x1 + x2) / 2 - 40, ty, 80, 11), XStringFormats.TopCenter);
    }

    private static void DimV(XGraphics gfx, XFont font, double faceX, double dimX,
        double y1, double y2, string label, bool accent)
    {
        var col = accent ? AccentColor : DimColor;
        var brush = accent ? AccentBrush : DimBrush;
        var lblFont = accent ? new XFont("Arial", font.Size, XFontStyleEx.Bold) : font;
        var lblBrush = accent ? AccentBrush : TextBrush;
        var pen = new XPen(col, accent ? 0.9 : 0.6);
        double dir = Math.Sign(dimX - faceX); if (dir == 0) dir = -1;
        gfx.DrawLine(pen, faceX + ExtGap * dir, y1, dimX + ExtOver * dir, y1);
        gfx.DrawLine(pen, faceX + ExtGap * dir, y2, dimX + ExtOver * dir, y2);
        gfx.DrawLine(pen, dimX, y1, dimX, y2);
        Arrow(gfx, new XPoint(dimX, y1), 0, -1, brush);
        Arrow(gfx, new XPoint(dimX, y2), 0, 1, brush);
        // place label on the far side of the dim line from the part
        bool right = dir > 0;
        double lx = right ? dimX + 6 : dimX - 54;
        gfx.DrawString(label, lblFont, lblBrush, new XRect(lx, (y1 + y2) / 2 - 6, 48, 11),
            right ? XStringFormats.TopLeft : XStringFormats.TopRight);
    }

    private static void Arrow(XGraphics gfx, XPoint tip, double dx, double dy, XBrush? brush = null)
    {
        const double len = 5, half = 1.8;
        double px = -dy, py = dx;
        var b1 = new XPoint(tip.X - dx * len + px * half, tip.Y - dy * len + py * half);
        var b2 = new XPoint(tip.X - dx * len - px * half, tip.Y - dy * len - py * half);
        gfx.DrawPolygon(brush ?? DimBrush, new[] { tip, b1, b2 }, XFillMode.Winding);
    }

    // ── Dimension / label primitives (aligned dims, leaders, a greedy non-overlap placer) ──

    public enum DimKind { Web, Flange, Lip }

    /// <summary>A cross-section dimension anchored to the two true material corners (model coords).
    /// <paramref name="Hem"/> distinguishes a 180° hem lip (dim mirrored from the flange dim) from a
    /// 90° return lip (dim placed vertically above the material).</summary>
    public readonly record struct CsDim(double X1, double Y1, double X2, double Y2, double Value, bool Inside, DimKind Kind, bool Hem = false);

    /// <summary>
    /// Cross-section dimensions for U/L/Z, anchored to the TRUE outer/inner sharp corners (each the
    /// intersection of the two adjacent faces offset by ±T/2). A dim therefore spans the exact outside
    /// (or inside) measurement regardless of bend radius or thickness — the witness lines land on the
    /// real edges. Exposed for numeric verification.
    /// </summary>
    public static List<CsDim> ComputeCrossSectionDims(FlatPatternResult fp)
    {
        var dims = new List<CsDim>();
        var prof = fp.Profile;
        if (prof.Count < 3) return dims;
        var s = fp.Spec;
        double t = s.Thickness;
        double cx = prof.Average(p => p.x), cy = prof.Average(p => p.y);
        double webO = fp.WebOutside, flL = fp.FlangeLeftOutside, flR = fp.FlangeRightOutside;
        var flanges = fp.SectionBends.Where(b => !b.IsReturn).ToList();

        (double x, double y) Um(double dx, double dy) { double l = Math.Sqrt(dx * dx + dy * dy); if (l < 1e-9) l = 1; return (dx / l, dy / l); }
        (double x, double y) Nout(double atx, double aty, double dx, double dy)   // outward normal (away from centroid)
        {
            var (ux, uy) = Um(dx, dy);
            double nx = -uy, ny = ux;
            if (nx * (atx - cx) + ny * (aty - cy) < 0) { nx = -nx; ny = -ny; }
            return (nx, ny);
        }
        // Sharp corner where the two faces meeting at a bend intersect, offset to outer(+1)/inner(-1).
        (double x, double y) Corner(SectionBend b, double sign)
        {
            var (a1x, a1y) = Um(-b.InHx, -b.InHy);
            var (a2x, a2y) = Um(b.OutHx, b.OutHy);
            var (n1x, n1y) = Nout(b.X, b.Y, a1x, a1y);
            var (n2x, n2y) = Nout(b.X, b.Y, a2x, a2y);
            return LineX(b.X + n1x * sign * t / 2, b.Y + n1y * sign * t / 2, a1x, a1y,
                         b.X + n2x * sign * t / 2, b.Y + n2y * sign * t / 2, a2x, a2y);
        }
        (double x, double y) AddFlange(SectionBend b, double dx, double dy, double outsideLen, bool inside)
        {
            var (ux, uy) = Um(dx, dy);
            var c = Corner(b, inside ? -1 : 1);
            double len = inside ? outsideLen - t : outsideLen;
            var top = (x: c.x + ux * len, y: c.y + uy * len);
            dims.Add(new CsDim(c.x, c.y, top.x, top.y, len, inside, DimKind.Flange));
            return top;   // flange free-edge corner — a hem on this flange dimensions from here
        }

        (double x, double y)? leftFlangeTop = null, rightFlangeTop = null;
        if (s.Type is PartType.UChannel or PartType.ZChannel && flanges.Count >= 2)
        {
            var b0 = flanges[0]; var b1 = flanges[1];
            bool wIn = s.Web.Basis == DimBasis.Inside;
            var c0 = Corner(b0, wIn ? -1 : 1); var c1 = Corner(b1, wIn ? -1 : 1);
            dims.Add(new CsDim(c0.x, c0.y, c1.x, c1.y, wIn ? webO - 2 * t : webO, wIn, DimKind.Web));
            leftFlangeTop  = AddFlange(b0, -b0.InHx, -b0.InHy, flL, s.FlangeLeft.Basis == DimBasis.Inside);
            rightFlangeTop = AddFlange(b1, b1.OutHx, b1.OutHy, flR, s.FlangeRight.Basis == DimBasis.Inside);
        }
        else if (s.Type == PartType.LAngle && flanges.Count >= 1)
        {
            // The single L bend joins legB (FlangeRight) on its incoming/-InH side and legA (FlangeLeft)
            // on its outgoing/OutH side — so the -InH leg carries the FlangeRight length, the OutH leg the
            // FlangeLeft length (the channel pairs each side with its own bend, the L shares one).
            var b = flanges[0];
            leftFlangeTop  = AddFlange(b, -b.InHx, -b.InHy, flR, s.FlangeRight.Basis == DimBasis.Inside);
            rightFlangeTop = AddFlange(b, b.OutHx, b.OutHy, flL, s.FlangeLeft.Basis == DimBasis.Inside);
        }

        // Return lips — dimension ALONG the lip's own face, from the bend's outer corner to the lip's
        // free edge. The lip is the free-edge segment adjacent to the return bend: on the LEFT it precedes
        // the bend (incoming side → lip heading = -InH), on the RIGHT it follows it (outgoing side → OutH).
        // (Using the chain side, not an away-from-centroid guess, so an inward-folded lip is dimensioned
        // along the lip rather than down the flange.)
        foreach (var rb in fp.SectionBends.Where(b => b.IsReturn))
        {
            bool leftReturn = rb.X < cx;
            var rs = leftReturn ? s.ReturnLeft : s.ReturnRight;
            rs ??= s.ReturnLeft ?? s.ReturnRight;
            if (rs is null) continue;
            var (ux, uy) = leftReturn ? Um(-rb.InHx, -rb.InHy) : Um(rb.OutHx, rb.OutHy);
            bool inside = rs.Basis == DimBasis.Inside;
            bool hem = rb.AngleDeg >= 170;
            // A 180° hem folds back alongside its flange, so Corner() is degenerate (the two faces are
            // parallel). Start the hem lip from the flange's free-edge corner — the same point as that
            // flange's dimension top. A 90° return has a real outer corner.
            var c = hem
                ? ((leftReturn ? leftFlangeTop : rightFlangeTop) ?? Corner(rb, inside ? -1 : 1))
                : Corner(rb, inside ? -1 : 1);
            dims.Add(new CsDim(c.x, c.y, c.x + ux * rs.Length, c.y + uy * rs.Length, rs.Length, inside, DimKind.Lip, hem));
        }
        return dims;
    }

    private static (double x, double y) LineX(double px, double py, double dx, double dy, double qx, double qy, double ex, double ey)
    {
        double denom = dx * ey - dy * ex;
        if (Math.Abs(denom) < 1e-9) return (px, py);
        double a = ((qx - px) * ey - (qy - py) * ex) / denom;
        return (px + a * dx, py + a * dy);
    }

    // An aligned (rotated) dimension parallel to the face p1→p2, offset outward (away from
    // <paramref name="awayFrom"/>): witness lines, a parallel dim line + arrowheads, and a label
    // placed clear of the part + other labels. Keeps the accent (as-specified) vs muted styling.
    private static void DimAligned(XGraphics gfx, XFont font, XPoint p1, XPoint p2, double offset,
        XPoint awayFrom, string label, bool accent, List<XRect> placed, XRect panel,
        (double x, double y)? forceDir = null, string? tag = null)
    {
        double dx = p2.X - p1.X, dy = p2.Y - p1.Y, l = Math.Sqrt(dx * dx + dy * dy);
        if (l < 1e-6) return;
        double ux = dx / l, uy = dy / l;        // along the face
        double nx = -uy, ny = ux;               // face normal
        double mx = (p1.X + p2.X) / 2, my = (p1.Y + p2.Y) / 2;
        // Pick the perpendicular SIDE: either an explicit direction (lips force this) or away from the
        // part centroid (the default for web/flange dims).
        if (forceDir is { } fd && (fd.x * fd.x + fd.y * fd.y) > 1e-12)
        {
            if (nx * fd.x + ny * fd.y < 0) { nx = -nx; ny = -ny; }
        }
        else if (nx * (mx - awayFrom.X) + ny * (my - awayFrom.Y) < 0) { nx = -nx; ny = -ny; }   // point outward

        var col = accent ? AccentColor : DimColor;
        var brush = accent ? AccentBrush : DimBrush;
        var pen = new XPen(col, accent ? 0.9 : 0.6);
        var o1 = new XPoint(p1.X + nx * offset, p1.Y + ny * offset);
        var o2 = new XPoint(p2.X + nx * offset, p2.Y + ny * offset);
        // p1/p2 ARE the exact material corners; a short stub past the edge (into the material) marks
        // which edge we're leading from, then the witness runs out to the dimension line.
        const double stub = 4;
        gfx.DrawLine(pen, new XPoint(p1.X - nx * stub, p1.Y - ny * stub), new XPoint(o1.X + nx * ExtOver, o1.Y + ny * ExtOver));
        gfx.DrawLine(pen, new XPoint(p2.X - nx * stub, p2.Y - ny * stub), new XPoint(o2.X + nx * ExtOver, o2.Y + ny * ExtOver));
        gfx.DrawLine(pen, o1, o2);
        Arrow(gfx, o1, -ux, -uy, brush);
        Arrow(gfx, o2, ux, uy, brush);

        var lblFont = accent ? new XFont("Arial", font.Size, XFontStyleEx.Bold) : font;
        var lblBrush = accent ? AccentBrush : TextBrush;
        var sz = gfx.MeasureString(label, lblFont);
        var anchor = new XPoint((o1.X + o2.X) / 2 + nx * 8, (o1.Y + o2.Y) / 2 + ny * 8);
        var rect = PlaceLabel(placed, panel, anchor, new XSize(sz.Width + 3, sz.Height + 1), new XPoint(nx, ny));
        var dimMid = new XPoint((o1.X + o2.X) / 2, (o1.Y + o2.Y) / 2);
        var lc = new XPoint(rect.X + rect.Width / 2, rect.Y + rect.Height / 2);
        if (Dist(lc, dimMid) > sz.Height + 12) gfx.DrawLine(pen, dimMid, lc);   // connector only if displaced
        gfx.DrawString(label, lblFont, lblBrush, rect, XStringFormats.Center);

        // Calibration tag: a red circled letter just left of the value label, keyed to the `dims` map.
        if (tag is not null)
        {
            const double r = 8;
            var c = new XPoint(rect.X - r - 2, rect.Y + rect.Height / 2);
            gfx.DrawEllipse(new XPen(XColors.Red, 1.1), new XRect(c.X - r, c.Y - r, 2 * r, 2 * r));
            gfx.DrawString(tag, new XFont("Arial", 9, XFontStyleEx.Bold), new XSolidBrush(XColors.Red),
                new XRect(c.X - r, c.Y - r, 2 * r, 2 * r), XStringFormats.Center);
        }
    }

    /// <summary>FAB-note item stamp: a black circled letter (A, B, C…) at (x, y), matching the lettered
    /// note list in the ERP viewer. Drawn in the title block of each appended FAB drawing page.</summary>
    private static void DrawItemTag(XGraphics gfx, string tag, double x, double y)
    {
        const double r = 9;
        var box = new XRect(x, y, 2 * r, 2 * r);
        gfx.DrawEllipse(new XPen(XColors.Black, 1.2), box);
        gfx.DrawString(tag, new XFont("Arial", 10, XFontStyleEx.Bold), XBrushes.Black, box, XStringFormats.Center);
    }

    // A label carried off a feature point by a leader, placed clear of the part + other labels.
    private static void LeaderLabel(XGraphics gfx, XFont font, XPoint tip, XPoint pushDir, string label,
        bool accent, List<XRect> placed, XRect panel)
    {
        var pen = new XPen(accent ? AccentColor : DimColor, accent ? 0.9 : 0.7);
        var brush = accent ? AccentBrush : TextBrush;
        var lblFont = accent ? new XFont("Arial", font.Size, XFontStyleEx.Bold) : font;
        var m = gfx.MeasureString(label, lblFont);
        var sz = new XSize(m.Width + 4, m.Height + 2);

        // Try the preferred direction, then alternatives, choosing one whose LEADER doesn't run through
        // an existing label (so the angle leader never crosses the flange dimension, etc.).
        var dirs = new[] { Unit(pushDir.X, pushDir.Y), new XPoint(0, -1), new XPoint(-1, 0), new XPoint(1, 0), new XPoint(0, 1) };
        XRect chosen = default; bool ok = false;
        foreach (var d in dirs)
        {
            var cand = TryPlace(placed, panel, new XPoint(tip.X + d.X * 20, tip.Y + d.Y * 20), sz, d);
            var cc = new XPoint(cand.X + cand.Width / 2, cand.Y + cand.Height / 2);
            if (!LeaderCrosses(tip, cc, placed)) { chosen = cand; ok = true; break; }
        }
        if (!ok) chosen = TryPlace(placed, panel, new XPoint(tip.X + pushDir.X * 20, tip.Y + pushDir.Y * 20), sz, pushDir);
        placed.Add(chosen);

        var lc = new XPoint(chosen.X + chosen.Width / 2, chosen.Y + chosen.Height / 2);
        gfx.DrawLine(pen, lc, tip);
        var dir = Unit(tip.X - lc.X, tip.Y - lc.Y);
        Arrow(gfx, tip, dir.X, dir.Y, accent ? AccentBrush : DimBrush);
        gfx.DrawString(label, lblFont, brush, chosen, XStringFormats.Center);
    }

    // A small square right-angle symbol in the corner where faces d1 and d2 meet at the apex.
    private static void RightAngleMark(XGraphics gfx, XPen pen, XPoint apex, XPoint d1, XPoint d2, double size)
    {
        var u1 = Unit(d1.X, d1.Y);
        var u2 = Unit(d2.X, d2.Y);
        var p1 = new XPoint(apex.X + u1.X * size, apex.Y + u1.Y * size);
        var p2 = new XPoint(apex.X + (u1.X + u2.X) * size, apex.Y + (u1.Y + u2.Y) * size);
        var p3 = new XPoint(apex.X + u2.X * size, apex.Y + u2.Y * size);
        gfx.DrawLine(pen, p1, p2);
        gfx.DrawLine(pen, p2, p3);
    }

    // Greedy placement: step the rect outward along pushDir until it clears every placed rect (with
    // clearance) and stays inside the panel; record + return it. Falls back to the clamped anchor.
    // Greedy placement that does NOT commit — returns the candidate rect (for trying alternatives).
    private static XRect TryPlace(List<XRect> placed, XRect panel, XPoint anchor, XSize size, XPoint pushDir)
    {
        const double clr = 3, stepLen = 4;
        var u = Unit(pushDir.X, pushDir.Y);
        for (int step = 0; step <= 48; step++)
        {
            double cx = anchor.X + u.X * step * stepLen, cy = anchor.Y + u.Y * step * stepLen;
            double rx = Math.Max(panel.X + 1, Math.Min(cx - size.Width / 2, panel.X + panel.Width - size.Width - 1));
            double ry = Math.Max(panel.Y + 1, Math.Min(cy - size.Height / 2, panel.Y + panel.Height - size.Height - 1));
            var r = new XRect(rx, ry, size.Width, size.Height);
            bool clash = false;
            foreach (var q in placed) if (Overlaps(r, q, clr)) { clash = true; break; }
            if (!clash) return r;
        }
        double fx = Math.Max(panel.X + 1, Math.Min(anchor.X - size.Width / 2, panel.X + panel.Width - size.Width - 1));
        double fy = Math.Max(panel.Y + 1, Math.Min(anchor.Y - size.Height / 2, panel.Y + panel.Height - size.Height - 1));
        return new XRect(fx, fy, size.Width, size.Height);
    }

    // Greedy placement that commits the chosen rect to <paramref name="placed"/>.
    private static XRect PlaceLabel(List<XRect> placed, XRect panel, XPoint anchor, XSize size, XPoint pushDir)
    {
        var r = TryPlace(placed, panel, anchor, size, pushDir);
        placed.Add(r);
        return r;
    }

    // Does the leader segment a→b cross any already-placed LABEL rect? (Skips placed[0] = the part bbox.)
    private static bool LeaderCrosses(XPoint a, XPoint b, List<XRect> placed)
    {
        for (int i = 1; i < placed.Count; i++)
            if (SegIntersectsRect(a, b, placed[i])) return true;
        return false;
    }

    private static bool SegIntersectsRect(XPoint p, XPoint q, XRect r)
    {
        bool In(XPoint pt) => pt.X >= r.X && pt.X <= r.X + r.Width && pt.Y >= r.Y && pt.Y <= r.Y + r.Height;
        if (In(p) || In(q)) return true;
        var tl = new XPoint(r.X, r.Y); var tr = new XPoint(r.X + r.Width, r.Y);
        var br = new XPoint(r.X + r.Width, r.Y + r.Height); var bl = new XPoint(r.X, r.Y + r.Height);
        return SegSeg(p, q, tl, tr) || SegSeg(p, q, tr, br) || SegSeg(p, q, br, bl) || SegSeg(p, q, bl, tl);
    }

    private static bool SegSeg(XPoint a, XPoint b, XPoint c, XPoint d)
    {
        double o1 = Orient(a, b, c), o2 = Orient(a, b, d), o3 = Orient(c, d, a), o4 = Orient(c, d, b);
        return (o1 > 0) != (o2 > 0) && (o3 > 0) != (o4 > 0);
    }

    private static double Orient(XPoint p, XPoint q, XPoint r) => (q.X - p.X) * (r.Y - p.Y) - (q.Y - p.Y) * (r.X - p.X);

    private static bool Overlaps(XRect a, XRect b, double clr)
        => a.X < b.X + b.Width + clr && a.X + a.Width + clr > b.X
        && a.Y < b.Y + b.Height + clr && a.Y + a.Height + clr > b.Y;

    private static XRect BBox(XPoint[] pts)
    {
        double minx = pts.Min(p => p.X), maxx = pts.Max(p => p.X);
        double miny = pts.Min(p => p.Y), maxy = pts.Max(p => p.Y);
        return new XRect(minx, miny, maxx - minx, maxy - miny);
    }

    private static XPoint Unit(double dx, double dy)
    {
        double l = Math.Sqrt(dx * dx + dy * dy);
        return l < 1e-6 ? new XPoint(0, -1) : new XPoint(dx / l, dy / l);
    }

    private static double Dist(XPoint a, XPoint b) { double dx = a.X - b.X, dy = a.Y - b.Y; return Math.Sqrt(dx * dx + dy * dy); }

    private static string F(double v) => v.ToString("0.000", CultureInfo.InvariantCulture);

    /// <summary>Dimension label in fractional inches with the inch sign (to 1/16; decimal below that), e.g. 2-3/16".</summary>
    private static string Fq(double v) => DrawFormat.FracInch(v);

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
