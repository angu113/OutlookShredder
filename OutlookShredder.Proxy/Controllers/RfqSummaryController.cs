using Microsoft.AspNetCore.Mvc;
using OutlookShredder.Proxy.Services;

namespace OutlookShredder.Proxy.Controllers;

[ApiController]
[Route("api/rfq")]
public class RfqSummaryController : ControllerBase
{
    private readonly RfqSummaryService               _summary;
    private readonly ILogger<RfqSummaryController>    _log;

    public RfqSummaryController(RfqSummaryService summary, ILogger<RfqSummaryController> log)
    {
        _summary = summary;
        _log     = log;
    }

    /// <summary>
    /// Produces a short (≤3 bullet) AI summary of an RFQ from a client-assembled text input
    /// (requested items + each supplier's coverage / prices / regrets). Returns { bullets: [] }.
    /// On any AI failure the bullets array is empty so the client keeps its deterministic summary.
    /// </summary>
    [HttpPost("summarize")]
    public async Task<IActionResult> Summarize([FromBody] RfqSummarizeRequest req, CancellationToken ct)
    {
        if (string.IsNullOrWhiteSpace(req?.Input))
            return Ok(new { bullets = Array.Empty<string>() });

        var bullets = await _summary.SummarizeAsync(req.Input!, ct);
        return Ok(new { bullets });
    }

    public record RfqSummarizeRequest(string? Input);
}
