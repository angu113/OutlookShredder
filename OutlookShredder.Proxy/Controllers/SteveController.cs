using Microsoft.AspNetCore.Cors;
using Microsoft.AspNetCore.Mvc;
using OutlookShredder.Proxy.Services;

namespace OutlookShredder.Proxy.Controllers;

/// <summary>
/// Relay for Steve — the RPA agent that automates OpenBravo tasks via the Chrome extension.
///
/// Shredder POST /api/steve/trigger { task }  → extension GET /api/steve/pending
///                                            → extension POST /api/steve/consume   (starts)
///                                            → extension POST /api/steve/complete  (done)
/// FileWatcherService calls SteveState.SetExportResult(path) when ExportedData*.csv lands.
/// StatementGenView polls GET /api/steve/export-result until the path appears.
///
/// CORS: same "PhoneSearch" policy so the content script on the OB origin can reach localhost:7000.
/// </summary>
[ApiController]
[Route("api/steve")]
[EnableCors("PhoneSearch")]
public class SteveController : ControllerBase
{
    private readonly ILogger<SteveController> _log;

    public SteveController(ILogger<SteveController> log) => _log = log;

    [HttpOptions]
    [HttpOptions("trigger")]
    [HttpOptions("pending")]
    [HttpOptions("consume")]
    [HttpOptions("complete")]
    [HttpOptions("export-result")]
    [HttpOptions("clear-result")]
    public IActionResult Preflight() => NoContent();

    // ── Trigger (Shredder → proxy) ────────────────────────────────────────────

    [HttpPost("trigger")]
    public IActionResult Trigger([FromBody] SteveTriggerRequest req)
    {
        if (string.IsNullOrWhiteSpace(req.Task))
            return BadRequest(new { error = "task is required" });
        SteveState.ClearExportResult();           // clear stale result from prior run
        SteveState.SetPending(req.Task);
        _log.LogInformation("[Steve] Task triggered: {Task}", req.Task);
        return Ok(new { ok = true });
    }

    // ── Pending / consume (extension polls these) ─────────────────────────────

    [HttpGet("pending")]
    public IActionResult GetPending()
    {
        var task = SteveState.GetPending();
        return Ok(new { task });
    }

    [HttpPost("consume")]
    public IActionResult Consume()
    {
        SteveState.ClearPending();
        return Ok(new { ok = true });
    }

    // ── Complete (extension reports back) ─────────────────────────────────────

    [HttpPost("complete")]
    public IActionResult Complete([FromBody] SteveCompleteRequest req)
    {
        _log.LogInformation("[Steve] Task complete: ok={Ok} file={File} error={Error}",
            req.Ok, req.FileName, req.Error);
        SteveState.ClearPending();
        return Ok(new { ok = true });
    }

    // ── Export result (StatementGenView polls this) ───────────────────────────

    [HttpGet("export-result")]
    public IActionResult GetExportResult()
    {
        var path = SteveState.GetExportResult();
        return Ok(new { path });
    }

    [HttpPost("clear-result")]
    public IActionResult ClearResult()
    {
        SteveState.ClearExportResult();
        return Ok(new { ok = true });
    }

    public record SteveTriggerRequest(string Task);
    public record SteveCompleteRequest(bool Ok, string? FileName, string? Error);
}
