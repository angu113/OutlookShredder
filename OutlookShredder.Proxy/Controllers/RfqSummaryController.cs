using Microsoft.AspNetCore.Mvc;
using OutlookShredder.Proxy.Services;

namespace OutlookShredder.Proxy.Controllers;

[ApiController]
[Route("api/rfq")]
public class RfqSummaryController : ControllerBase
{
    private readonly RfqSummaryService             _summary;
    private readonly RfqStateOfPlayService         _state;
    private readonly SharePointService             _sp;
    private readonly IConfiguration                _config;
    private readonly ILogger<RfqSummaryController> _log;

    public RfqSummaryController(RfqSummaryService summary, RfqStateOfPlayService state,
        SharePointService sp, IConfiguration config, ILogger<RfqSummaryController> log)
    {
        _summary = summary;
        _state   = state;
        _sp      = sp;
        _config  = config;
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

    // ── GET /api/rfq/{rfqId}/state-of-play ────────────────────────────────────
    /// <summary>The cached AI "state of play" comparison for an RFQ. Generates on demand when the cache
    /// is missing or its InputsHash is stale (then persists it). `summary` is null when there are no
    /// quotes yet / AI is unavailable and nothing was cached.</summary>
    [HttpGet("{rfqId}/state-of-play")]
    public async Task<IActionResult> GetStateOfPlay(string rfqId, CancellationToken ct)
    {
        try
        {
            var rows   = await _sp.ReadSupplierItemsByRfqIdAsync(rfqId);
            bool pdfs  = _config.GetValue("RfqStateOfPlay:IncludePdfs", false);
            var hash   = _state.ComputeInputsHash(rows, pdfs);
            var cached = await _sp.ReadRfqSummaryAsync(rfqId);
            if (cached?.Summary is { Length: > 0 } && cached.InputsHash == hash)
                return Ok(new { rfqId, summary = cached.Summary, generatedAt = cached.GeneratedAt, model = cached.Model, mode = cached.Mode, fresh = true });

            var result = await _state.GenerateAsync(rfqId, rows, pdfs, null, ct);
            if (result is null)   // AI unavailable / no quotes — keep whatever was cached
                return Ok(new { rfqId, summary = cached?.Summary, generatedAt = cached?.GeneratedAt, model = cached?.Model, mode = cached?.Mode, fresh = false });

            await _sp.WriteRfqSummaryAsync(rfqId, result.Summary, result.InputsHash, result.Model, result.Mode);
            return Ok(new { rfqId, summary = result.Summary, generatedAt = DateTimeOffset.UtcNow.ToString("o"), model = result.Model, mode = result.Mode, fresh = true });
        }
        catch (Exception ex) { _log.LogError(ex, "GetStateOfPlay failed for {Rfq}", rfqId); return StatusCode(500, new { error = ex.Message }); }
    }

    // ── POST /api/rfq/{rfqId}/state-of-play/regenerate?mode=text|pdf ───────────
    /// <summary>Generates a state-of-play in the requested mode and RETURNS it WITHOUT touching the
    /// durable cache — for eyeballing text-only vs with-PDFs side by side.</summary>
    [HttpPost("{rfqId}/state-of-play/regenerate")]
    public async Task<IActionResult> RegenerateStateOfPlay(string rfqId, [FromQuery] string? mode, CancellationToken ct)
    {
        try
        {
            var rows   = await _sp.ReadSupplierItemsByRfqIdAsync(rfqId);
            bool pdfs  = string.Equals(mode, "pdf", StringComparison.OrdinalIgnoreCase);
            var result = await _state.GenerateAsync(rfqId, rows, pdfs, null, ct);
            return Ok(new { rfqId, mode = pdfs ? "pdf" : "text", summary = result?.Summary, model = result?.Model });
        }
        catch (Exception ex) { _log.LogError(ex, "RegenerateStateOfPlay failed for {Rfq}", rfqId); return StatusCode(500, new { error = ex.Message }); }
    }

    public record RfqSummarizeRequest(string? Input);
}
