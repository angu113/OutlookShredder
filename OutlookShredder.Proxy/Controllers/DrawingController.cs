using Microsoft.AspNetCore.Mvc;
using OutlookShredder.Proxy.Services;
using OutlookShredder.Proxy.Services.Drawing;
using PdfSharp.Drawing;
using PdfSharp.Pdf;

namespace OutlookShredder.Proxy.Controllers;

[ApiController]
public class DrawingController : ControllerBase
{
    private readonly ILogger<DrawingController> _log;

    public DrawingController(ILogger<DrawingController> log) => _log = log;

    public sealed class GenerateRequest
    {
        /// <summary>Plain-text part description, e.g. "U 4 x 2 x 36, 16ga CRS, inside".</summary>
        public string? Text { get; set; }

        /// <summary>Calibration mode: letters each cross-section dimension (A, B, C…) on the drawing and
        /// returns the letter→geometry mapping in <c>dims</c>, so the dimension anchoring can be calibrated.</summary>
        public bool Calibrate { get; set; }

        /// <summary>Order quantity from the wizard's Qty box — printed as the "xN" line on the DXF's
        /// no-cut Notes label. Null ⇒ no quantity line.</summary>
        public int? Quantity { get; set; }
    }

    // Material display name + the token the canonical text uses + its gauge-table family.
    private static readonly (string Name, string Token, MaterialFamily Fam)[] Materials =
    {
        ("Cold Rolled Steel", "CRS",    MaterialFamily.ColdRolled),
        ("Hot Rolled Steel",  "HRS",    MaterialFamily.HotRolled),
        ("Galvanized Steel",  "galv",   MaterialFamily.Galvanized),
        ("Stainless Steel",   "SS",     MaterialFamily.Stainless),
        ("Aluminum",          "alum",   MaterialFamily.Aluminium),
        ("Brass",             "brass",  MaterialFamily.Brass),
        ("Copper",            "copper", MaterialFamily.Copper),
    };

    /// <summary>
    /// Materials for the wizard: each with the canonical text token and its gauge list
    /// (gauge number + decimal thickness in inches), so the thickness picker can show both.
    /// </summary>
    [HttpGet("/api/drawing/materials")]
    public IActionResult MaterialList() => Ok(Materials.Select(m => new
    {
        name = m.Name,
        token = m.Token,
        gauges = GaugeTables.GaugesFor(m.Fam).Select(g => new { gauge = g.Gauge, thickness = g.Thickness }),
    }));

    /// <summary>
    /// ASME B16.48 paddle-blind ("frying pan") sizes for the wizard's NPS + class pickers and the
    /// standard-thickness hint. The proxy holds the table; the lookup happens at generate time.
    /// </summary>
    [HttpGet("/api/drawing/paddle-blanks")]
    public IActionResult PaddleBlanks() => Ok(PaddleBlankTable.All.Select(p => new
    {
        nps = p.Nps,
        npsValue = p.NpsValue,
        pressureClass = p.PressureClass,
        od = p.Od,
        centerToEnd = p.CenterToEnd,
        handleWidth = p.HandleWidth,
        thickness = p.Thickness,
    }));

    /// <summary>
    /// Parses a part description, develops the flat pattern, and returns the resolved spec,
    /// bend math, and cut geometry. (PDF/DXF bytes added in the next step.)
    /// </summary>
    [HttpPost("/api/drawing/generate")]
    public IActionResult Generate([FromBody] GenerateRequest req)
    {
        if (req is null || string.IsNullOrWhiteSpace(req.Text))
            return BadRequest(new { error = "text is required" });

        try
        {
            var spec = DrawingTextParser.Parse(req.Text);
            var fp = FlatPattern.Develop(spec, req.Quantity);

            byte[] pdf = DrawingPdfRenderer.Render(fp, req.Calibrate);
            byte[] dxf = DrawingDxfWriter.Write(fp.Cut);

            // Calibration: the cross-section dimensions in draw order, each keyed to its drawing letter,
            // so the user can say "letter B should be the return lip OD" and we fix that anchor.
            object? dims = null;
            if (req.Calibrate)
                dims = DrawingPdfRenderer.ComputeCrossSectionDims(fp)
                    .Select((d, i) => new
                    {
                        label = ((char)('A' + i)).ToString(),
                        kind  = d.Kind.ToString(),
                        value = d.Value,
                        basis = d.Inside ? "ID" : "OD",
                        hem   = d.Hem,
                        x1 = d.X1, y1 = d.Y1, x2 = d.X2, y2 = d.Y2,
                    }).ToList();

            return Ok(new
            {
                ok = true,
                input = req.Text,
                pdfBase64 = Convert.ToBase64String(pdf),
                dxfBase64 = Convert.ToBase64String(dxf),
                suggestedFileName = fp.Cut.Part,
                spec = new
                {
                    type = spec.Type.ToString(),
                    material = spec.Material,
                    units = spec.Units,
                    thickness = spec.Thickness,
                    insideRadius = spec.InsideRadius,
                    kFactor = spec.KFactor,
                    angleDeg = spec.AngleDeg,
                    web = new { spec.Web.Value, basis = spec.Web.Basis.ToString() },
                    flangeLeft = new { spec.FlangeLeft.Value, basis = spec.FlangeLeft.Basis.ToString() },
                    flangeRight = new { spec.FlangeRight.Value, basis = spec.FlangeRight.Basis.ToString() },
                    length = spec.Length,
                    polishDirection = spec.PolishDirection.ToString(),
                },
                math = new
                {
                    ossb = fp.Ossb,
                    bendAllowance = fp.BendAllowance,
                    bendDeduction = fp.BendDeduction,
                    webOutside = fp.WebOutside,
                    flangeLeftOutside = fp.FlangeLeftOutside,
                    flangeRightOutside = fp.FlangeRightOutside,
                    flatWidth = fp.FlatWidth,
                    flatHeight = fp.FlatHeight,
                    bendLinesX = fp.BendLinesX,
                },
                title = fp.Title,
                summary = fp.Summary,
                cut = fp.Cut,
                dims,
                bends = req.Calibrate
                    ? fp.SectionBends.Select(b => new
                    {
                        x = b.X, y = b.Y, inHx = b.InHx, inHy = b.InHy, outHx = b.OutHx, outHy = b.OutHy,
                        angle = b.AngleDeg, dir = b.Dir.ToString(), isReturn = b.IsReturn,
                    }).ToList()
                    : null,
            });
        }
        catch (DrawingParseException ex)
        {
            return BadRequest(new { error = ex.Message });
        }
        catch (Exception ex)
        {
            _log.LogWarning(ex, "[Drawing] generate failed for input: {Input}", req.Text);
            return StatusCode(500, new { error = ex.Message });
        }
    }

    /// <summary>Returns just the rendered PDF (visual-test harness). Body: { text }.</summary>
    [HttpPost("/api/drawing/pdf")]
    public IActionResult Pdf([FromBody] GenerateRequest req)
    {
        if (req is null || string.IsNullOrWhiteSpace(req.Text))
            return BadRequest(new { error = "text is required" });
        try
        {
            var fp = FlatPattern.Develop(DrawingTextParser.Parse(req.Text));
            return File(DrawingPdfRenderer.Render(fp), "application/pdf");
        }
        catch (DrawingParseException ex) { return BadRequest(new { error = ex.Message }); }
        catch (Exception ex)
        {
            _log.LogWarning(ex, "[Drawing] pdf failed for input: {Input}", req.Text);
            return StatusCode(500, new { error = ex.Message });
        }
    }

    /// <summary>
    /// Self-test for the FAB-append pipeline: synthesizes a one-page "slip" carrying a FAB note,
    /// runs <see cref="PickingSlipFabAppender"/>, and returns the combined PDF. The result should
    /// have the slip page plus one appended drawing page.
    /// </summary>
    [HttpGet("/api/drawing/fab-selftest")]
    public IActionResult FabSelfTest([FromQuery] string? text)
    {
        var desc = string.IsNullOrWhiteSpace(text) ? "U 4 x 2 x 36, 16ga CRS, finish outside" : text;
        PickingSlipEnricher.EnsureFontResolver();

        byte[] slip;
        using (var doc = new PdfDocument())
        {
            var page = doc.AddPage();
            using (var gfx = XGraphics.FromPdfPage(page))
            {
                var font = new XFont("Arial", 11);
                gfx.DrawString("PICKING SLIP (FAB self-test)", font, XBrushes.Black, new XPoint(50, 60));
                gfx.DrawString($"FAB: (2) {desc}", font, XBrushes.Black, new XPoint(50, 100));
            }
            using var ms = new MemoryStream();
            doc.Save(ms);
            slip = ms.ToArray();
        }

        var combined = PickingSlipFabAppender.AppendFabDrawings(slip, _log);
        return File(combined, "application/pdf");
    }
}
