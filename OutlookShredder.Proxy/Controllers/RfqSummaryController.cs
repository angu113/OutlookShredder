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
    private readonly RfqNotificationService        _notify;
    private readonly IConfiguration                _config;
    private readonly ILogger<RfqSummaryController> _log;

    public RfqSummaryController(RfqSummaryService summary, RfqStateOfPlayService state,
        SharePointService sp, RfqNotificationService notify, IConfiguration config, ILogger<RfqSummaryController> log)
    {
        _summary = summary;
        _state   = state;
        _sp      = sp;
        _notify  = notify;
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
            // Read-only by default — the response pipeline (queue) owns generation + keeps the cache
            // fresh. The UI only triggers a one-off TEXT fallback when nothing has been cached yet.
            var cached = await _sp.ReadRfqSummaryAsync(rfqId);
            if (cached?.Summary is { Length: > 0 })
                return Ok(new { rfqId, summary = cached.Summary, generatedAt = cached.GeneratedAt, model = cached.Model, mode = cached.Mode, fresh = true });

            var rows = await _sp.ReadSupplierItemsByRfqIdAsync(rfqId);
            if (RfqStateOfPlayService.CompetingSuppliers(rows) < 2)
                return Ok(new { rfqId, summary = (string?)null, fresh = false });   // nothing to compare yet

            var result = await _state.GenerateAsync(rfqId, rows, includePdfs: false, sentAgo: null, ct);
            if (result is null)
                return Ok(new { rfqId, summary = (string?)null, fresh = false });

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

    // ── POST /api/rfq/state-of-play/refresh-stale?days=&rfqIds=&dryRun= ────────
    /// <summary>Rechecks cached state-of-play summaries against current SLI data (via InputsHash) and
    /// regenerates + persists the STALE ones — an existing summary whose inputs changed, e.g. after a
    /// price re-extraction/backfill. dryRun=true reports which are stale WITHOUT spending AI. Scope =
    /// explicit rfqIds (comma-separated) or every reference created within `days`. Never bulk-creates
    /// brand-new summaries (skips RFQs that were never summarized).</summary>
    [HttpPost("state-of-play/refresh-stale")]
    public async Task<IActionResult> RefreshStaleSummaries(
        [FromQuery] int days = 7, [FromQuery] string? rfqIds = null, [FromQuery] bool dryRun = false,
        CancellationToken ct = default)
    {
        try
        {
            List<string> ids;
            if (!string.IsNullOrWhiteSpace(rfqIds))
                ids = rfqIds.Split(',', StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries).ToList();
            else
            {
                var refs = await _sp.ReadRfqReferencesAsync();
                var cut  = DateTimeOffset.UtcNow.AddDays(-Math.Abs(days));
                ids = refs.Where(r => DateTimeOffset.TryParse(
                            r.TryGetValue("DateCreated", out var d) ? d?.ToString() : null, out var dt) && dt >= cut)
                    .Select(r => r.TryGetValue("RFQ_ID", out var v) ? v?.ToString() : null)
                    .Where(s => !string.IsNullOrWhiteSpace(s)).Select(s => s!).Distinct().ToList();
            }

            bool pdfs = _config.GetValue("RfqStateOfPlay:IncludePdfs", false);
            _log.LogInformation("[refresh-stale] scanning {Total} RFQ(s) (dryRun={Dry})", ids.Count, dryRun);
            int withSummary = 0, stale = 0, regen = 0, idx = 0;
            var staleIds = new List<string>();
            foreach (var rfqId in ids)
            {
                idx++;
                var rows = await _sp.ReadSupplierItemsByRfqIdAsync(rfqId);
                if (RfqStateOfPlayService.CompetingSuppliers(rows) < 2) continue;
                var cached = await _sp.ReadRfqSummaryAsync(rfqId);
                if (cached?.Summary is not { Length: > 0 }) continue;          // never summarized — leave for on-demand
                withSummary++;
                if (cached.InputsHash == _state.ComputeInputsHash(rows, pdfs)) continue;   // fresh
                stale++; staleIds.Add(rfqId);
                if (dryRun) continue;
                var result = await _state.GenerateAsync(rfqId, rows, pdfs, null, ct);
                if (result is not null)
                {
                    await _sp.WriteRfqSummaryAsync(rfqId, result.Summary, result.InputsHash, result.Model, result.Mode);
                    _notify.NotifyRfqSummary(rfqId);
                    regen++;
                    _log.LogInformation("[refresh-stale] progress {Idx}/{Total}: regenerated {RfqId} ({Regen} done, stale #{Stale})",
                        idx, ids.Count, rfqId, regen, stale);
                }
            }
            _log.LogInformation("[refresh-stale] DONE: regenerated={Regen} of {Stale} stale ({WithSummary} had summaries, {Total} scanned, dryRun={Dry})",
                regen, stale, withSummary, ids.Count, dryRun);
            return Ok(new { scope = string.IsNullOrWhiteSpace(rfqIds) ? $"last {days}d" : "explicit",
                            dryRun, withSummary, stale, regenerated = regen, staleRfqs = staleIds });
        }
        catch (Exception ex) { _log.LogError(ex, "RefreshStaleSummaries failed"); return StatusCode(500, new { error = ex.Message }); }
    }

    // ── GET /api/rfq/{rfqId}/winner-block?dimsAware= (deterministic, NO AI — for backtesting) ──
    /// <summary>The raw deterministic winner block for an RFQ (no AI cost). dimsAware=false reproduces the
    /// legacy MSPC-alone pooling for A/B regression comparison against the new dimension-aware pooling.</summary>
    [HttpGet("{rfqId}/winner-block")]
    public async Task<IActionResult> WinnerBlock(string rfqId, [FromQuery] bool dimsAware = true)
    {
        try
        {
            var rows = await _sp.ReadSupplierItemsByRfqIdAsync(rfqId);
            return Ok(new { rfqId, dimsAware, block = _state.WinnerBlockForDiag(rows, dimsAware) });
        }
        catch (Exception ex) { _log.LogError(ex, "WinnerBlock diag failed for {Rfq}", rfqId); return StatusCode(500, new { error = ex.Message }); }
    }

    public record RfqSummarizeRequest(string? Input);
}
