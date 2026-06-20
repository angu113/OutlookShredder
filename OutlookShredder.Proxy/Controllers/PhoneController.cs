using Microsoft.AspNetCore.Mvc;
using OutlookShredder.Proxy.Models;
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
            var records = await _sp.ReadAllPhoneCallLogAsync(ct);

            int noPhone = 0, noMatch = 0, unchanged = 0, addBlank = 0, changeExisting = 0, multiMatch = 0;
            var plan         = new List<(PhoneCallLogRecord Rec, CustomerLookupResult Match)>();
            var addSamples   = new List<string>();
            var changeSamples = new List<string>();

            foreach (var rec in records)
            {
                var digits = CustomerImportService.NormalizePhone(rec.CallerPhone);
                if (digits is null) { noPhone++; continue; }

                var matches = _crm.LookupAllByPhone(digits);
                if (matches.Count == 0) { noMatch++; continue; }
                if (matches.Count > 1) multiMatch++;
                var m = matches[0];   // primary match — same as live detection's call-log write

                switch (ClassifyBackfill(rec, m.BusinessPartner, m.ContactName, m.PopupMessage))
                {
                    case CrmBackfillAction.Unchanged:
                        unchanged++;
                        break;
                    case CrmBackfillAction.AddBp:
                        addBlank++;
                        plan.Add((rec, m));
                        if (addSamples.Count < 40)
                            addSamples.Add($"{rec.CallerPhone} -> '{m.BusinessPartner}' (was blank)");
                        break;
                    case CrmBackfillAction.ChangeBp:
                        changeExisting++;
                        if (includeChanged) plan.Add((rec, m));
                        if (changeSamples.Count < 40)
                            changeSamples.Add($"{rec.CallerPhone}: '{rec.BpName}' -> '{m.BusinessPartner}'");
                        break;
                }
            }

            int updated = 0, failed = 0;
            if (!dryRun)
            {
                foreach (var (rec, m) in plan)
                {
                    try
                    {
                        await _sp.UpdateCallLogCrmByItemIdAsync(
                            rec.SpItemId, m.BusinessPartner, m.ContactName, m.PopupMessage, ct);
                        updated++;
                    }
                    catch (Exception ex)
                    {
                        failed++;
                        _log.LogWarning(ex, "[Phone] backfill patch failed for item {Id}", rec.SpItemId);
                    }
                }
                _log.LogInformation("[Phone] CRM backfill: {Updated} updated, {Failed} failed (of {Plan} planned)",
                    updated, failed, plan.Count);
            }

            return Ok(new
            {
                dryRun, includeChanged,
                totalRecords   = records.Count,
                wouldUpdate    = plan.Count,
                addBlank, changeExisting, unchanged, noMatch, noPhone,
                multiMatchPhones = multiMatch,
                updated, failed,
                addSamples, changeSamples,
            });
        }
        catch (Exception ex)
        {
            _log.LogWarning(ex, "[Phone] CRM backfill failed");
            return StatusCode(500, new { error = ex.Message });
        }
    }

    internal enum CrmBackfillAction { Unchanged, AddBp, ChangeBp }

    /// <summary>Decides the backfill action for one call-log row given its matched CRM values. Pure +
    /// testable: Unchanged when all three CRM fields already equal the match (null/blank- and trim-
    /// insensitive); AddBp when BpName is currently blank (a gain); otherwise ChangeBp (would overwrite).</summary>
    internal static CrmBackfillAction ClassifyBackfill(
        PhoneCallLogRecord rec, string? matchBp, string? matchContact, string? matchPopup)
    {
        static bool Same(string? a, string? b) =>
            string.Equals((a ?? "").Trim(), (b ?? "").Trim(), StringComparison.Ordinal);

        if (Same(rec.BpName, matchBp) && Same(rec.ContactName, matchContact) && Same(rec.PopupMessage, matchPopup))
            return CrmBackfillAction.Unchanged;

        return string.IsNullOrWhiteSpace(rec.BpName) ? CrmBackfillAction.AddBp : CrmBackfillAction.ChangeBp;
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
