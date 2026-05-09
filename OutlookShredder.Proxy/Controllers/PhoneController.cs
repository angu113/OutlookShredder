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
        [FromQuery] int top = 500,
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

    /// <summary>Deletes a single call log entry by its SP item ID.</summary>
    [HttpDelete("call-log/{spItemId}")]
    public async Task<IActionResult> DeleteCallLogEntry(string spItemId, CancellationToken ct = default)
    {
        try
        {
            await _sp.DeletePhoneCallLogItemAsync(spItemId, ct);
            return NoContent();
        }
        catch (Exception ex)
        {
            _log.LogWarning(ex, "[Phone] Failed to delete call log entry {Id}", spItemId);
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

    /// <summary>Patches BpName/ContactName/PopupMessage on all call log entries for a given phone number.</summary>
    [HttpPatch("call-log/update-crm")]
    public async Task<IActionResult> UpdateCrm(
        [FromQuery] string phone,
        [FromBody] UpdateCrmRequest req,
        CancellationToken ct = default)
    {
        if (string.IsNullOrWhiteSpace(phone))
            return BadRequest(new { error = "phone is required" });
        try
        {
            await _sp.UpdateCallLogCrmByPhoneAsync(phone, req.BpName, req.ContactName, req.PopupMessage, ct);
            return NoContent();
        }
        catch (Exception ex)
        {
            _log.LogWarning(ex, "[Phone] Failed to update CRM for phone {Phone}", phone);
            return StatusCode(500, new { error = ex.Message });
        }
    }
}

public record UpdateNotesRequest(string? Notes);
public record UpdateCrmRequest(string? BpName, string? ContactName, string? PopupMessage);
