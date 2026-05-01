using Microsoft.AspNetCore.Mvc;
using OutlookShredder.Proxy.Services;

namespace OutlookShredder.Proxy.Controllers;

/// <summary>Dev-only helpers for triggering events without real hardware/mail.</summary>
[ApiController]
[Route("api/test")]
public class TestController(RfqNotificationService notify, SharePointService sp) : ControllerBase
{
    /// <summary>
    /// POST /api/test/zoom-call — fires an IncomingCall bus event as if Zoom detected an
    /// incoming call.  Useful for testing the Shredder call-toast UI without a real call.
    /// If bpName/contactName/popupMessage are not provided, performs a real CRM lookup
    /// so the test matches the live Zoom path.
    /// </summary>
    [HttpPost("zoom-call")]
    public async Task<IActionResult> ZoomCall(
        [FromQuery] string  callerName   = "Test Caller",
        [FromQuery] string  callerPhone  = "5550000000",
        [FromQuery] string? bpName       = null,
        [FromQuery] string? popupMessage = null,
        [FromQuery] string? contactName  = null,
        CancellationToken ct = default)
    {
        // If CRM fields not explicitly supplied, do the real lookup.
        if (bpName is null && contactName is null && popupMessage is null
            && !string.IsNullOrWhiteSpace(callerPhone))
        {
            var crm = await sp.LookupCustomerByPhoneAsync(callerPhone, ct);
            if (crm is not null)
            {
                bpName      = crm.BusinessPartner;
                popupMessage = crm.PopupMessage;
                contactName = crm.ContactName;
            }
        }

        notify.NotifyIncomingCall(callerName, callerPhone, bpName, popupMessage, contactName);
        return Ok(new { fired = true, callerName, callerPhone, bpName, popupMessage, contactName });
    }
}
