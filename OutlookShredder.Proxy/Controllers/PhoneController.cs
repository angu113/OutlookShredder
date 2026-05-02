using Microsoft.AspNetCore.Mvc;
using OutlookShredder.Proxy.Services;

namespace OutlookShredder.Proxy.Controllers;

[ApiController]
[Route("api/phone")]
public class PhoneController : ControllerBase
{
    private readonly SharePointService _sp;
    private readonly ILogger<PhoneController> _log;

    public PhoneController(SharePointService sp, ILogger<PhoneController> log)
    {
        _sp  = sp;
        _log = log;
    }

    /// <summary>Returns the most recent call log entries, newest first.</summary>
    [HttpGet("call-log")]
    public async Task<IActionResult> GetCallLog(
        [FromQuery] int top = 200,
        CancellationToken ct = default)
    {
        try
        {
            var records = await _sp.ReadPhoneCallLogAsync(top, ct);
            return Ok(records);
        }
        catch (Exception ex)
        {
            _log.LogWarning(ex, "[Phone] Failed to read call log");
            return StatusCode(500, new { error = ex.Message });
        }
    }

    /// <summary>Updates the free-text notes on a call log entry.</summary>
    [HttpPatch("call-log/{spItemId}/notes")]
    public async Task<IActionResult> UpdateNotes(
        string spItemId,
        [FromBody] UpdateNotesRequest req,
        CancellationToken ct = default)
    {
        try
        {
            await _sp.UpdateCallLogNotesAsync(spItemId, req.Notes ?? "", ct);
            return NoContent();
        }
        catch (Exception ex)
        {
            _log.LogWarning(ex, "[Phone] Failed to update notes for {Id}", spItemId);
            return StatusCode(500, new { error = ex.Message });
        }
    }
}

public record UpdateNotesRequest(string? Notes);
