using Microsoft.AspNetCore.Mvc;
using OutlookShredder.Proxy.Models;
using OutlookShredder.Proxy.Services;

namespace OutlookShredder.Proxy.Controllers;

[ApiController]
public class SynonymController : ControllerBase
{
    private readonly ProductSynonymService             _synonyms;
    private readonly ILogger<SynonymController>        _log;

    public SynonymController(
        ProductSynonymService synonyms,
        ILogger<SynonymController> log)
    {
        _synonyms = synonyms;
        _log      = log;
    }

    // ── GET /api/synonyms ─────────────────────────────────────────────────────

    [HttpGet("api/synonyms")]
    public IActionResult GetAll() => Ok(_synonyms.Groups);

    // ── POST /api/synonyms ────────────────────────────────────────────────────

    [HttpPost("api/synonyms")]
    public async Task<IActionResult> Add(
        [FromBody] SynonymGroup group,
        CancellationToken ct)
    {
        if (string.IsNullOrWhiteSpace(group.Canonical))
            return BadRequest(new { error = "Canonical is required" });
        try
        {
            var saved = await _synonyms.AddSynonymAsync(group, ct);
            return Ok(saved);
        }
        catch (Exception ex)
        {
            _log.LogError(ex, "[Synonyms] Failed to add synonym '{Canonical}'", group.Canonical);
            return StatusCode(500, new { error = ex.Message });
        }
    }

    // ── PUT /api/synonyms/{spItemId} ──────────────────────────────────────────

    [HttpPut("api/synonyms/{spItemId}")]
    public async Task<IActionResult> Update(
        string spItemId,
        [FromBody] SynonymGroup group,
        CancellationToken ct)
    {
        if (string.IsNullOrWhiteSpace(group.Canonical))
            return BadRequest(new { error = "Canonical is required" });
        try
        {
            var saved = await _synonyms.UpdateSynonymAsync(spItemId, group, ct);
            return Ok(saved);
        }
        catch (Exception ex)
        {
            _log.LogError(ex, "[Synonyms] Failed to update synonym {Id}", spItemId);
            return StatusCode(500, new { error = ex.Message });
        }
    }

    // ── POST /api/synonyms/reload ─────────────────────────────────────────────
    // Reloads the in-memory cache from SP without restarting the proxy.

    [HttpPost("api/synonyms/reload")]
    public async Task<IActionResult> Reload(CancellationToken ct)
    {
        try
        {
            await _synonyms.LoadAsync(ct);
            return Ok(new { loaded = _synonyms.Groups.Count });
        }
        catch (Exception ex)
        {
            _log.LogError(ex, "[Synonyms] Reload failed");
            return StatusCode(500, new { error = ex.Message });
        }
    }
}
