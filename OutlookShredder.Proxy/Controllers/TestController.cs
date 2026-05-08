using Microsoft.AspNetCore.Mvc;
using OutlookShredder.Proxy.Models;
using OutlookShredder.Proxy.Services;

namespace OutlookShredder.Proxy.Controllers;

/// <summary>Dev-only helpers for triggering events without real hardware/mail.</summary>
[ApiController]
[Route("api/test")]
public class TestController(RfqNotificationService notify, CustomerCacheService crmCache) : ControllerBase
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
            var crm = crmCache.LookupByPhone(callerPhone);
            if (crm is not null)
            {
                bpName       = crm.BusinessPartner;
                popupMessage = crm.PopupMessage;
                contactName  = crm.ContactName;
            }
        }

        notify.NotifyIncomingCall(callerName, callerPhone, bpName, popupMessage, contactName);
        return Ok(new { fired = true, callerName, callerPhone, bpName, popupMessage, contactName });
    }

    /// <summary>
    /// POST /api/test/erp-doc — fires an ErpDocument bus event as if the file watcher
    /// processed a new ERP PDF.  Triggers the Trigger window to restore and take focus.
    /// </summary>
    [HttpPost("erp-doc")]
    public IActionResult ErpDoc(
        [FromQuery] string documentNumber = "TEST-001",
        [FromQuery] string customerName   = "Test Customer",
        [FromQuery] string documentType   = "PickingSlip")
    {
        notify.NotifyRfqProcessed(new RfqProcessedNotification { EventType = "ErpDocument" });
        return Ok(new { fired = true, documentNumber, customerName, documentType });
    }

    /// <summary>
    /// POST /api/test/supplier-response — fires an SR bus event as if a supplier email
    /// was processed.  Triggers the RFQ tab to refresh and show a toast.
    /// </summary>
    [HttpPost("supplier-response")]
    public IActionResult SupplierResponse(
        [FromQuery] string rfqId        = "TEST01",
        [FromQuery] string supplierName = "Test Supplier",
        [FromQuery] string product      = "Test Product",
        [FromQuery] double totalPrice   = 999.99)
    {
        notify.NotifyRfqProcessed(new RfqProcessedNotification
        {
            EventType    = "SR",
            RfqId        = rfqId,
            SupplierName = supplierName,
            MessageId    = $"test-{Guid.NewGuid():N}",
            Products     = [new RfqNotificationProduct { Name = product, TotalPrice = (double)totalPrice }],
        });
        return Ok(new { fired = true, rfqId, supplierName, product, totalPrice });
    }
}
