using System.Text;
using System.Text.Json;
using Microsoft.AspNetCore.Mvc;
using OutlookShredder.Proxy.Services;

namespace OutlookShredder.Proxy.Controllers;

/// <summary>
/// Dev-only eval harness (wip/sidecar/project-classification-eval-harness.md).
/// Runs the real classifier in-process over the human-labeled MailGoldenLabels corpus
/// and reports per-leaf P/R/F1, confusion matrix, and confidence calibration.
/// Zero workflow side effects — never writes to MailClassifications, the bus, or matchers.
/// </summary>
[ApiController]
[Route("api/mail-eval")]
public sealed class MailEvalController : ControllerBase
{
    private readonly MailEvalService _eval;

    public MailEvalController(MailEvalService eval) => _eval = eval;

    /// <summary>Start an eval run in the background and return immediately.</summary>
    [HttpPost("run")]
    public IActionResult Run([FromBody] EvalRunRequest req)
        => Accepted(_eval.StartRun(req));

    /// <summary>Live snapshot: running flag, processed/total, rolling accuracy.</summary>
    [HttpGet("status")]
    public IActionResult Status() => Ok(_eval.GetSnapshot());

    /// <summary>Full metrics report from the most recent completed run (null when none yet).</summary>
    [HttpGet("report")]
    public IActionResult Report()
    {
        var r = _eval.GetReport();
        return r is null ? NoContent() : Ok(r);
    }

    /// <summary>Per-item results as JSONL (one JSON object per line).</summary>
    [HttpGet("results")]
    public IActionResult Results()
    {
        var run = _eval.GetResults();
        if (run is null) return NoContent();

        var sb = new StringBuilder();
        var opts = new JsonSerializerOptions { PropertyNamingPolicy = JsonNamingPolicy.CamelCase };
        foreach (var r in run.Items)
            sb.AppendLine(JsonSerializer.Serialize(r, opts));
        return Content(sb.ToString(), "application/x-ndjson", Encoding.UTF8);
    }

    /// <summary>
    /// Bootstrap MailGoldenLabels from the current in-memory AI classifications (one row per item).
    /// The seeded label is the AI's own guess — human correction is REQUIRED before running the eval.
    /// Pass ?overwrite=true to replace existing rows (default: skip existing human corrections).
    /// </summary>
    [HttpPost("seed-golden")]
    public IActionResult SeedGolden([FromQuery] bool overwrite = false)
    {
        // Fire-and-forget with CancellationToken.None — client disconnect must not abort mid-write.
        _ = Task.Run(() => _eval.SeedGoldenFromCurrentsAsync(overwrite, CancellationToken.None));
        return Accepted(new { message = "Seeding started in background. Watch proxy logs for [MailEval] seed N/total progress." });
    }
}
