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
        // Each side: an uploaded file, else the CSV staged by a prior fetch (OB via Steve / Heartland via the GP portal).
        var obCsv  = obFile is not null        ? await ReadAsync(obFile, ct)        : recon.StagedObCsv;
        var hlCsv  = heartlandFile is not null  ? await ReadAsync(heartlandFile, ct) : recon.StagedHeartlandCsv;
        var obName = obFile?.FileName        ?? "OpenBravo (fetched)";
        var hlName = heartlandFile?.FileName ?? "Heartland (fetched)";
        if (obCsv is null)
            return BadRequest(new { error = "No OB payments — upload an OB file or use Fetch from OpenBravo first." });
        if (hlCsv is null)
            return BadRequest(new { error = "No Heartland data — upload a Heartland file or use Fetch Heartland first." });
        try
        {
            return Ok(recon.Run(obCsv, hlCsv, obName, hlName));
        }
        catch (Exception ex)
        {
            return StatusCode(500, new { error = ex.Message });
        }
    }

    /// <summary>
    /// Fetches the OB Payment-In export via the Steve OB automation and stages it as the OB side
    /// (content-validated as a Payment-In file). Requires OpenBravo open with the Shredder extension.
    /// Then call <c>run</c> with just the Heartland file.
    /// </summary>
    [HttpPost("ob-export/fetch")]
    public async Task<IActionResult> FetchObExport(CancellationToken ct)
    {
        var (ok, message, rows) = await recon.FetchObViaSteveAsync(ct);
        return ok ? Ok(new { ok, message, rows }) : StatusCode(504, new { ok, message });
    }

    /// <summary>
    /// Fetches the Heartland/GP "Merchant Batch Download" report via the GP portal (SSRS URL access,
    /// driven by the heartland.js content script) and stages it as the processor side. Requires the GP
    /// portal open and freshly 2FA-authenticated — the portal times out fast.
    /// </summary>
    [HttpPost("heartland/fetch")]
    public async Task<IActionResult> FetchHeartland(CancellationToken ct)
    {
        var (ok, message, rows) = await recon.FetchHeartlandViaPortalAsync(ct);
        return ok ? Ok(new { ok, message, rows }) : StatusCode(504, new { ok, message });
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
