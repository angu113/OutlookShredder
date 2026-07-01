using Microsoft.AspNetCore.Mvc;
using OutlookShredder.Proxy.Models;
using OutlookShredder.Proxy.Services;

namespace OutlookShredder.Proxy.Controllers;

[ApiController]
[Route("api/phone")]
public class PhoneController : ControllerBase
{
    private readonly SharePointService    _sp;
    private readonly RfqNotificationService _notify;
    private readonly CustomerCacheService _crm;
    private readonly CallLogCrmBackfillService _backfill;
    private readonly IConfiguration       _config;
    private readonly ILogger<PhoneController> _log;

    public PhoneController(
        SharePointService sp,
        RfqNotificationService notify,
        CustomerCacheService crm,
        CallLogCrmBackfillService backfill,
        IConfiguration config,
        ILogger<PhoneController> log)
    {
        _sp     = sp;
        _notify = notify;
        _crm    = crm;
        _backfill = backfill;
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

    /// <summary>Targeted backfill for ONE phone right after a manual link: refreshes the CRM cache so the
    /// just-added contact is live, then writes the matched BpName/ContactName/PopupMessage onto every call-log
    /// row for that phone. (The team-wide variant is call-log/backfill-crm.)</summary>
    [HttpPost("call-log/backfill-customer")]
    public async Task<IActionResult> BackfillCustomer([FromQuery] string phone, CancellationToken ct = default)
    {
        if (string.IsNullOrWhiteSpace(phone))
            return BadRequest(new { error = "phone is required" });
        try
        {
            await _crm.RefreshNowAsync(ct);   // pick up the newly-linked contact before looking it up
            var match = _crm.LookupByPhone(phone);
            if (match is null)
                return Ok(new { phone, matched = false });

            await _sp.UpdateCallLogCrmByPhoneAsync(
                phone, match.BusinessPartner, match.ContactName, match.PopupMessage, ct);
            _log.LogInformation("[Phone] backfill-customer {Phone} -> bp='{Bp}' contact='{Contact}' popup={HasPopup}",
                phone, match.BusinessPartner, match.ContactName ?? "(none)", !string.IsNullOrEmpty(match.PopupMessage));
            return Ok(new
            {
                phone, matched = true, bp = match.BusinessPartner,
                contact = match.ContactName, hasPopup = !string.IsNullOrEmpty(match.PopupMessage),
            });
        }
        catch (Exception ex)
        {
            _log.LogWarning(ex, "[Phone] backfill-customer failed for {Phone}", phone);
            return StatusCode(500, new { error = ex.Message });
        }
    }

    /// <summary>
    /// Backfills CRM fields (BpName / ContactName / PopupMessage) on existing call-log rows using the SAME
    /// phone→customer lookup as live incoming-call detection (CustomerCacheService.LookupAllByPhone, primary
    /// match). Reads every historic row, normalises its CallerPhone, looks it up, and sets the matched CRM
    /// values. dryRun=true (default) reports what WOULD change without writing. includeChanged=false fills
    /// only blank BpName rows (never overwrites an existing value).
    /// </summary>
    [HttpPost("call-log/backfill-crm")]
    public async Task<IActionResult> BackfillCrm(
        [FromQuery] bool dryRun = true,
        [FromQuery] bool includeChanged = true,
        CancellationToken ct = default)
    {
        try
        {
            var r = await _backfill.RunAsync(dryRun, includeChanged, ct);
            return Ok(new
            {
                dryRun, includeChanged,
                totalRecords   = r.TotalRecords,
                wouldUpdate    = r.WouldUpdate,
                addBlank = r.AddBlank, changeExisting = r.ChangeExisting, unchanged = r.Unchanged,
                noMatch = r.NoMatch, noPhone = r.NoPhone,
                multiMatchPhones = r.MultiMatchPhones,
                updated = r.Updated, failed = r.Failed,
                addSamples = r.AddSamples, changeSamples = r.ChangeSamples,
            });
        }
        catch (Exception ex)
        {
            _log.LogWarning(ex, "[Phone] CRM backfill failed");
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

        var receivedAt = DateTimeOffset.UtcNow;
        string spItemId = "";
        try
        {
            spItemId = await _sp.WritePhoneCallLogAsync(
                name, phone, resolvedBp, resolvedContact, resolvedPopup,
                receivedAt, ct);
        }
        catch (Exception ex)
        {
            _log.LogWarning(ex, "[Phone] test-call: call-log write failed");
        }

        _notify.NotifyIncomingCall(name, phone, resolvedBp, resolvedPopup, resolvedContact,
            callLogSpItemId: spItemId, receivedAt: receivedAt.ToString("o"));
        _log.LogInformation("[Phone] test-call fired — name='{Name}' phone='{Phone}' bp='{Bp}' spItemId={Id}",
            name, phone, resolvedBp ?? "none", spItemId);

        return Ok(new { name, phone, bp = resolvedBp, contact = resolvedContact, popup = resolvedPopup, spItemId });
    }
}

public record UpdateNotesRequest(string? Notes);
public record UpdateCrmRequest(string? BpName, string? ContactName, string? PopupMessage);
