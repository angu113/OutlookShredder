using Microsoft.AspNetCore.Mvc;
using OutlookShredder.Proxy.Services;

namespace OutlookShredder.Proxy.Controllers;

[ApiController]
[Route("api/customers")]
public class CustomersController(
    CustomerImportService importer,
    SharePointService     sp,
    CustomerCacheService  crmCache) : ControllerBase
{
    /// <summary>
    /// POST /api/customers/setup-lists — provisions Customers + CustomerContacts SP lists.
    /// Idempotent; safe to re-run.
    /// </summary>
    [HttpPost("setup-lists")]
    public async Task<IActionResult> SetupLists(CancellationToken ct)
    {
        var results = await sp.EnsureCustomerListsAsync(ct);
        return Ok(results);
    }

    /// <summary>
    /// POST /api/customers/import-partners — reads a business-partner CSV (Name + Popup Message)
    /// and upserts into the Customers SP list.  Send the CSV as the raw request body.
    /// </summary>
    [HttpPost("import-partners")]
    [RequestSizeLimit(50_000_000)]
    public async Task<IActionResult> ImportPartners(CancellationToken ct)
    {
        var csv = await ReadBodyAsync(ct);
        if (csv is null)
            return BadRequest("Send the CSV file as the raw request body.");

        var parsed = importer.ParsePartners(csv);
        if (parsed.Rows.Count == 0 && parsed.Warnings.Count > 0)
            return BadRequest(new { warnings = parsed.Warnings });

        var (added, updated, skipped) = await sp.UpsertBusinessPartnersAsync(parsed.Rows, ct);
        return Ok(new { parsed = parsed.Rows.Count, added, updated, skipped, warnings = parsed.Warnings });
    }

    /// <summary>
    /// POST /api/customers/import-contacts — reads a contacts CSV (Customer Name, Contact Name, Phone)
    /// and upserts into the CustomerContacts SP list.  For each (CustomerName, ContactName) pair in
    /// the file, existing rows are wiped and replaced with the freshly parsed phones.
    /// Send the CSV as the raw request body.
    /// </summary>
    [HttpPost("import-contacts")]
    [RequestSizeLimit(50_000_000)]
    public async Task<IActionResult> ImportContacts(CancellationToken ct)
    {
        var csv = await ReadBodyAsync(ct);
        if (csv is null)
            return BadRequest("Send the CSV file as the raw request body.");

        var parsed = importer.ParseContacts(csv);
        if (parsed.Rows.Count == 0 && parsed.Warnings.Count > 0)
            return BadRequest(new { warnings = parsed.Warnings });

        var (added, unchanged) = await sp.UpsertContactsAsync(parsed.Rows, ct);
        return Ok(new { parsed = parsed.Rows.Count, added, unchanged, warnings = parsed.Warnings });
    }

    /// <summary>
    /// GET /api/customers/lookup?phone=XXXXXXXXXX — looks up a business partner and contact
    /// by phone number.  Phone is normalised (strips formatting, drops leading country code).
    /// Returns 404 when no match is found.
    /// </summary>
    [HttpGet("lookup")]
    public async Task<IActionResult> Lookup([FromQuery] string phone, CancellationToken ct)
    {
        if (string.IsNullOrWhiteSpace(phone))
            return BadRequest("phone query parameter is required");

        var result = await sp.LookupCustomerByPhoneAsync(phone, ct);
        if (result is null) return NotFound();
        return Ok(result);
    }

    /// <summary>
    /// GET /api/customers/contacts — returns all CustomerContacts rows (CustomerName, ContactName, Phone).
    /// </summary>
    [HttpGet("contacts")]
    public IActionResult GetContacts()
    {
        return Ok(crmCache.GetAllContacts());
    }

    /// <summary>
    /// GET /api/customers/business-partners — returns all BP names from the Customers list, sorted.
    /// </summary>
    [HttpGet("business-partners")]
    public IActionResult GetBusinessPartners()
    {
        return Ok(crmCache.GetAllPartnerNames());
    }

    /// <summary>
    /// POST /api/customers/contacts/add — adds a single contact row directly (bypassing CSV import).
    /// isErpOrphan=true marks contacts added manually so they can be reported separately.
    /// </summary>
    [HttpPost("contacts/add")]
    public async Task<IActionResult> AddContact(
        [FromBody] AddContactRequest req,
        CancellationToken ct)
    {
        if (string.IsNullOrWhiteSpace(req.CustomerName) || string.IsNullOrWhiteSpace(req.ContactName))
            return BadRequest("customerName and contactName are required.");

        var phone = CustomerImportService.NormalizePhone(req.Phone);
        if (phone is null)
            return BadRequest(new { error = $"Invalid phone number: {req.Phone}" });

        await sp.AddContactDirectAsync(req.CustomerName, req.ContactName, phone, req.IsErpOrphan, ct);
        crmCache.Invalidate();
        return Ok(new { customerName = req.CustomerName, contactName = req.ContactName, phone, isErpOrphan = req.IsErpOrphan });
    }

    /// <summary>
    /// DELETE /api/customers/contact?bp=...&amp;contact=...&amp;phone=... — removes a specific
    /// (CustomerName, ContactName, Phone) triple from CustomerContacts.
    /// Returns 404 if no matching row is found.
    /// </summary>
    [HttpDelete("contact")]
    public async Task<IActionResult> DeleteContact(
        [FromQuery] string bp,
        [FromQuery] string contact,
        [FromQuery] string phone,
        CancellationToken ct)
    {
        if (string.IsNullOrWhiteSpace(bp) || string.IsNullOrWhiteSpace(contact) || string.IsNullOrWhiteSpace(phone))
            return BadRequest("bp, contact and phone query parameters are all required.");

        var deleted = await sp.DeleteContactAsync(bp, contact, phone, ct);
        return deleted ? Ok(new { deleted = true }) : NotFound();
    }

    private async Task<string?> ReadBodyAsync(CancellationToken ct)
    {
        if (Request.ContentLength is 0 or null) return null;
        using var ms = new System.IO.MemoryStream();
        await Request.Body.CopyToAsync(ms, ct);
        return ms.Length == 0 ? null : System.Text.Encoding.UTF8.GetString(ms.ToArray());
    }
}

public record AddContactRequest(string CustomerName, string ContactName, string Phone, bool IsErpOrphan = false);
