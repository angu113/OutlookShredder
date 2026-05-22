using Microsoft.AspNetCore.Mvc;
using OutlookShredder.Proxy.Services;

namespace OutlookShredder.Proxy.Controllers;

[ApiController]
[Route("api/phone")]
public class PhoneController : ControllerBase
{
    private readonly SharePointService    _sp;
    private readonly ProxyLeaseService    _lease;
    private readonly RfqNotificationService _notify;
    private readonly CustomerCacheService _crm;
    private readonly IConfiguration       _config;
    private readonly ILogger<PhoneController> _log;

    public PhoneController(
        SharePointService sp,
        ProxyLeaseService lease,
        RfqNotificationService notify,
        CustomerCacheService crm,
        IConfiguration config,
        ILogger<PhoneController> log)
    {
        _sp     = sp;
        _lease  = lease;
        _notify = notify;
        _crm    = crm;
        _config = config;
        _log    = log;
    }

    /// <summary>Returns call log entries for the last Phone:LookbackHours hours (default 36).</summary>
    [HttpGet("call-log")]
    public async Task<IActionResult> GetCallLog(CancellationToken ct = default)
    {
        try
        {
            var hours = _config.GetValue("Phone:LookbackHours", 36.0);
            var since = DateTimeOffset.UtcNow.AddHours(-hours);
            var records = await _sp.ReadPhoneCallLogAsync(since, ct);
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
    /// Injects a synthetic IncomingCall bus event — bypasses the WinEvent hook so the
    /// full Service Bus → Shredder toast pipeline can be verified without a real Zoom call.
    /// Optional query params: name, phone, bp, contact, popup.
    /// Writes a real call-log row to SP (use DELETE /api/phone/call-log/{id} to clean up).
    /// </summary>
    [HttpPost("test-call")]
    public async Task<IActionResult> TestCall(
        [FromQuery] string  name    = "Test Caller",
        [FromQuery] string  phone   = "(000) 000-0000",
        [FromQuery] string? bp      = null,
        [FromQuery] string? contact = null,
        [FromQuery] string? popup   = null,
        CancellationToken ct = default)
    {
        // If no bp supplied, try a CRM lookup so the test exercises the real lookup path
        string? resolvedBp      = bp;
        string? resolvedContact = contact;
        string? resolvedPopup   = popup;
        if (bp is null)
        {
            var digits = string.Concat(phone.Where(char.IsDigit));
            if (digits.Length >= 10)
            {
                var matches = _crm.LookupAllByPhone(digits[^10..]);
                if (matches.Count > 0)
                {
                    resolvedBp      = matches[0].BusinessPartner;
                    resolvedContact = matches[0].ContactName;
                    resolvedPopup   = matches[0].PopupMessage;
                }
            }
        }

        string spItemId = "";
        try
        {
            spItemId = await _sp.WritePhoneCallLogAsync(
                name, phone, resolvedBp, resolvedContact, resolvedPopup,
                DateTimeOffset.UtcNow, ct);
        }
        catch (Exception ex)
        {
            _log.LogWarning(ex, "[Phone] test-call: call-log write failed");
        }

        _notify.NotifyIncomingCall(name, phone, resolvedBp, resolvedPopup, resolvedContact,
            callLogSpItemId: spItemId);
        _log.LogInformation("[Phone] test-call fired — name='{Name}' phone='{Phone}' bp='{Bp}' spItemId={Id}",
            name, phone, resolvedBp ?? "none", spItemId);

        return Ok(new { name, phone, bp = resolvedBp, contact = resolvedContact, popup = resolvedPopup, spItemId });
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
