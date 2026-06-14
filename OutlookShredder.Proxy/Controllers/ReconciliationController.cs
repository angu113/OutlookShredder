using Microsoft.AspNetCore.Mvc;
using OutlookShredder.Proxy.Services;

namespace OutlookShredder.Proxy.Controllers;

/// <summary>
/// ShadowCat Payment Reconciliation — matches OpenBravo card payments against the Heartland /
/// Global Payments transaction export. v1: the client uploads both CSVs (picked from Downloads);
/// the proxy parses + matches and returns the result. Resolution tracking lands in a later increment.
/// </summary>
[ApiController]
[Route("api/reconciliation")]
public class ReconciliationController(PaymentReconciliationService recon) : ControllerBase
{
    /// <summary>
    /// Run a reconciliation over two uploaded CSVs: <c>obFile</c> (OB Payment In "ExportedData") and
    /// <c>heartlandFile</c> (Heartland transaction export). Returns the full <c>ReconRunResult</c>
    /// (rows, per-day batch totals, per-pay-type subtotals, counts).
    /// </summary>
    [HttpPost("run")]
    [Consumes("multipart/form-data")]
    public async Task<IActionResult> Run([FromForm] IFormFile? obFile, [FromForm] IFormFile? heartlandFile, CancellationToken ct)
    {
        if (obFile is null || heartlandFile is null)
            return BadRequest(new { error = "Both obFile and heartlandFile are required." });
        try
        {
            var obCsv = await ReadAsync(obFile, ct);
            var hlCsv = await ReadAsync(heartlandFile, ct);
            return Ok(recon.Run(obCsv, hlCsv, obFile.FileName, heartlandFile.FileName));
        }
        catch (Exception ex)
        {
            return StatusCode(500, new { error = ex.Message });
        }
    }

    /// <summary>Last reconciliation result (in-memory cache), for re-opening the tool without re-running.</summary>
    [HttpGet("results")]
    public IActionResult Results() =>
        Ok(new { ready = recon.GetLastResult() is not null, status = recon.Status, lastRunAt = recon.LastRunAt, result = recon.GetLastResult() });

    private static async Task<string> ReadAsync(IFormFile f, CancellationToken ct)
    {
        using var s = f.OpenReadStream();
        using var r = new StreamReader(s);
        return await r.ReadToEndAsync(ct);
    }
}
