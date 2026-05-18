using Microsoft.AspNetCore.Mvc;
using OutlookShredder.Proxy.Services;

namespace OutlookShredder.Proxy.Controllers;

[ApiController]
[Route("api/phone")]
public class PhoneController : ControllerBase
{
    private readonly SharePointService _sp;
    private readonly ProxyLeaseService _lease;
    private readonly ILogger<PhoneController> _log;

    public PhoneController(SharePointService sp, ProxyLeaseService lease, ILogger<PhoneController> log)
    {
        _sp    = sp;
        _lease = lease;
        _log   = log;
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

    /// <summary>
    /// Force-steals the Zoom watcher lease from whichever machine currently holds it.
    /// The ProxyLeaseService on this instance will start the Zoom hook within 30 s.
    /// </summary>
    [HttpPost("zoom/claim-lease")]
    public async Task<IActionResult> ClaimZoomLease(CancellationToken ct)
    {
        try
        {
            var machine = Environment.MachineName;
            var prev    = await _sp.ForceClaimLeaseAsync(ProxyLeaseService.ServiceName, machine, ct: ct);
            _log.LogInformation("[Phone] Zoom lease force-claimed on {Machine} (was: {Prev})", machine, prev ?? "none");
            return Ok(new { machine, previousHolder = prev ?? "none", message = "Lease claimed — Zoom watcher will start within 30 s" });
        }
        catch (Exception ex)
        {
            _log.LogWarning(ex, "[Phone] Failed to claim Zoom lease");
            return StatusCode(500, new { error = ex.Message });
        }
    }
}

public record UpdateNotesRequest(string? Notes);
public record UpdateCrmRequest(string? BpName, string? ContactName, string? PopupMessage);
