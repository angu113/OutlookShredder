using Microsoft.AspNetCore.Mvc;
using OutlookShredder.Proxy.Models;
using OutlookShredder.Proxy.Services;

namespace OutlookShredder.Proxy.Controllers;

[ApiController]
[Route("api")]
public class SupplierConversationsController : ControllerBase
{
    private readonly SharePointService    _sp;
    private readonly MailService          _mail;
    private readonly SupplierCacheService _suppliers;
    private readonly RfqNotificationService _notify;
    private readonly ILogger<SupplierConversationsController> _log;

    public SupplierConversationsController(
        SharePointService    sp,
        MailService          mail,
        SupplierCacheService suppliers,
        RfqNotificationService notify,
        ILogger<SupplierConversationsController> log)
    {
        _sp        = sp;
        _mail      = mail;
        _suppliers = suppliers;
        _notify    = notify;
        _log       = log;
    }

    /// <summary>Request body for POST /api/supplier-conversations/mark-read.</summary>
    public sealed class MarkReadRequest
    {
        public string  RfqId        { get; set; } = "";
        public string  SupplierName { get; set; } = "";
        public string  MessageId    { get; set; } = "";
        public bool    Read         { get; set; } = true;
        public string? ReadBy       { get; set; }   // display only ("read by … · time")
        public string? UserId       { get; set; }   // per-user read-state key (Environment.UserName)
    }

    /// <summary>
    /// Returns the merged conversation thread for the given RFQ/supplier pair, ordered oldest-first.
    /// By default includes both outbound (SupplierConversations list) and inbound (SupplierResponses list);
    /// set <paramref name="outboundOnly"/>=true to skip the SR query — useful when the caller already has the
    /// inbound data in memory from the RFQ grid and just needs the follow-ups.
    /// </summary>
    [HttpGet("supplier-conversations")]
    public async Task<IActionResult> Get(
        [FromQuery] string rfqId,
        [FromQuery] string supplierName,
        [FromQuery] bool   outboundOnly = false,
        [FromQuery] string? userId = null)
    {
        if (string.IsNullOrWhiteSpace(rfqId) || string.IsNullOrWhiteSpace(supplierName))
            return BadRequest(new { error = "rfqId and supplierName are required" });

        try
        {
            // userId optional here — when absent, per-user read flags are simply not applied (badge path
            // requires it). The client always sends it (ships in lockstep with this proxy).
            var messages = outboundOnly
                ? await _sp.ReadOutboundConversationAsync(userId ?? "", rfqId, supplierName)
                : await _sp.ReadConversationAsync(userId ?? "", rfqId, supplierName);
            return Ok(new { rfqId, supplierName, messages });
        }
        catch (Exception ex)
        {
            _log.LogError(ex, "[Conv] Read failed for {RfqId} / {Supplier}", rfqId, supplierName);
            return StatusCode(500, new { error = ex.Message });
        }
    }

    /// <summary>
    /// Marks one inbound supplier message read/unread (team-wide) by MessageId, then publishes a
    /// "SupplierMsgRead" bus event so peers update the message + the unread badge cascade.
    /// </summary>
    [HttpPost("supplier-conversations/mark-read")]
    public async Task<IActionResult> MarkRead([FromBody] MarkReadRequest req)
    {
        if (req is null || string.IsNullOrWhiteSpace(req.RfqId) ||
            string.IsNullOrWhiteSpace(req.SupplierName) || string.IsNullOrWhiteSpace(req.MessageId))
            return BadRequest(new { error = "rfqId, supplierName and messageId are required" });
        if (string.IsNullOrWhiteSpace(req.UserId))
            return BadRequest(new { error = "userId is required" });

        try
        {
            var patched = await _sp.MarkSupplierMessageReadAsync(req.UserId, req.MessageId, req.Read);
            // Publish scoped to the user so only THAT user's other machines refresh (read is per-user now).
            _notify.NotifySupplierMessageRead(req.UserId, req.RfqId, req.SupplierName, req.MessageId, req.Read, req.ReadBy);
            return Ok(new { ok = true, patched });
        }
        catch (Exception ex)
        {
            _log.LogError(ex, "[Conv] mark-read failed for {RfqId}/{Supplier} msg={Msg}",
                req.RfqId, req.SupplierName, req.MessageId);
            return StatusCode(500, new { error = ex.Message });
        }
    }

    /// <summary>
    /// Unread inbound-supplier-message tallies across all RFQs: grand total + per-RFQ + per-(rfq|supplier).
    /// Powers the RFQ-tab / RFQ-header / SR-header unread badges.
    /// </summary>
    [HttpGet("supplier-conversations/unread")]
    public async Task<IActionResult> GetUnread([FromQuery] string userId)
    {
        if (string.IsNullOrWhiteSpace(userId))
            return BadRequest(new { error = "userId is required" });
        try
        {
            var counts = await _sp.GetSupplierUnreadCountsAsync(userId);
            return Ok(new { total = counts.Total, byRfq = counts.ByRfq, bySupplier = counts.BySupplier });
        }
        catch (Exception ex)
        {
            _log.LogError(ex, "[Conv] unread tally failed");
            return StatusCode(500, new { error = ex.Message });
        }
    }

    /// <summary>
    /// One-time clean-slate: stamp every existing inbound supplier message read (no reader attribution),
    /// so only messages arriving after this start unread. Idempotent.
    /// </summary>
    [HttpPost("supplier-conversations/backfill-read")]
    public IActionResult BackfillRead() =>
        // Deprecated no-op: the team-wide clean-slate is replaced by the per-user lazy watermark — each
        // user's profile is created with AdoptedAt=now on their first unread fetch.
        Ok(new { ok = true, stamped = 0, deprecated = true });

    /// <summary>
    /// Returns the three contact email addresses (primary, manager, OOO) for the given supplier.
    /// Always returns 200 — fields are null when the supplier is not in the cache or the
    /// column has no value.
    /// </summary>
    [HttpGet("supplier-contacts")]
    public IActionResult GetContacts([FromQuery] string supplierName)
    {
        if (string.IsNullOrWhiteSpace(supplierName))
            return BadRequest(new { error = "supplierName required" });

        var contacts = _suppliers.GetContactsForSupplier(supplierName);
        return Ok(new
        {
            contactEmail    = contacts?.ContactEmail,
            managerContact  = contacts?.ManagerContact,
            oooContact      = contacts?.OooContact,
            primaryFirstName = contacts?.PrimaryFirstName,
            managerFirstName = contacts?.ManagerFirstName,
            oooFirstName     = contacts?.OooFirstName,
        });
    }

    /// <summary>
    /// Returns all suppliers that have at least one contact email configured.
    /// </summary>
    [HttpGet("suppliers/data")]
    public IActionResult GetAllSuppliersData() => Ok(_suppliers.GetAllSuppliersFull());

    [HttpPatch("suppliers/{spItemId}")]
    public async Task<IActionResult> PatchSupplier(string spItemId, [FromBody] Dictionary<string, string?> fields)
    {
        await _suppliers.PatchSupplierAsync(spItemId, fields);
        return Ok(new { patched = spItemId });
    }

    [HttpPost("suppliers/create")]
    public async Task<IActionResult> CreateSupplier([FromBody] Dictionary<string, string?> fields)
    {
        var id = await _suppliers.CreateSupplierAsync(fields);
        return Ok(new { created = id });
    }

    [HttpGet("supplier-contacts/all")]
    public IActionResult GetAllContacts()
    {
        var result = _suppliers.CachedNames
            .Select(name => { var c = _suppliers.GetContactsForSupplier(name); return new
            {
                supplierName     = name,
                contactEmail     = c?.ContactEmail,
                managerContact   = c?.ManagerContact,
                oooContact       = c?.OooContact,
                primaryFirstName = c?.PrimaryFirstName,
                managerFirstName = c?.ManagerFirstName,
                oooFirstName     = c?.OooFirstName,
            }; })
            .Where(x => x.contactEmail != null || x.managerContact != null || x.oooContact != null)
            .OrderBy(x => x.supplierName)
            .ToList();
        return Ok(result);
    }

    /// <summary>
    /// Returns a summary list of all MSG conversations, ordered most-recently-updated first.
    /// </summary>
    [HttpGet("supplier-conversations/msg-list")]
    public async Task<IActionResult> GetMsgList()
    {
        try
        {
            var summaries = await _sp.ReadMsgConversationSummariesAsync();
            return Ok(summaries);
        }
        catch (Exception ex)
        {
            _log.LogError(ex, "[Conv] Failed to read MSG conversation list");
            return StatusCode(500, new { error = ex.Message });
        }
    }

    /// <summary>
    /// Sends a follow-up email to a supplier about an existing RFQ and appends an
    /// outbound row to SupplierConversations.
    /// </summary>
    [HttpPost("supplier-inquiry/send")]
    public async Task<IActionResult> Send([FromBody] SupplierInquiryRequest req)
    {
        if (req is null) return BadRequest(new { error = "body required" });
        if (string.IsNullOrWhiteSpace(req.To))           return BadRequest(new { error = "to required" });
        if (string.IsNullOrWhiteSpace(req.RfqId))        return BadRequest(new { error = "rfqId required" });
        if (string.IsNullOrWhiteSpace(req.SupplierName)) return BadRequest(new { error = "supplierName required" });
        if (string.IsNullOrWhiteSpace(req.Subject))      return BadRequest(new { error = "subject required" });

        if (req.To.EndsWith("@mithrilmetals.com", StringComparison.OrdinalIgnoreCase))
            return BadRequest(new { error = "@mithrilmetals.com is not a valid supplier address" });

        byte[]? attachmentBytes = null;
        if (!string.IsNullOrEmpty(req.AttachmentContentBase64))
        {
            try   { attachmentBytes = Convert.FromBase64String(req.AttachmentContentBase64); }
            catch { return BadRequest(new { error = "attachmentContentBase64 is not valid base64" }); }
        }

        // Merge multi-BCC list with legacy single-BCC field.
        var bccAddresses = req.BccAddresses
            ?.Where(a => !string.IsNullOrWhiteSpace(a) &&
                         !a.EndsWith("@mithrilmetals.com", StringComparison.OrdinalIgnoreCase))
            .ToList();
        if ((bccAddresses is null || bccAddresses.Count == 0) && !string.IsNullOrWhiteSpace(req.Bcc))
            bccAddresses = [req.Bcc];

        try
        {
            var graphConvId = await _mail.SendSupplierInquiryAsync(
                req.To, req.Subject, req.Body,
                req.AttachmentName, attachmentBytes, req.AttachmentContentType,
                bccAddresses: bccAddresses?.Count > 0 ? bccAddresses : null);

            var sliVer = await _sp.GetCurrentSliVersionAsync(req.RfqId, req.SupplierName);
            var spId   = await _sp.WriteConversationMessageAsync(new ConversationMessage
            {
                RfqId               = req.RfqId,
                SupplierName        = req.SupplierName,
                SupplierResponseId  = req.SupplierResponseId,
                Direction           = "out",
                InReplyTo           = req.InReplyTo,
                SentAt              = DateTimeOffset.UtcNow,
                Subject             = req.Subject,
                BodyText            = req.Body,
                HasAttachments      = attachmentBytes is { Length: > 0 },
                ExtractedPricing    = false,
                SliVersionAtSend    = sliVer,
                GraphConversationId = graphConvId,
                ContactEmail        = req.To,
                BccAddresses        = bccAddresses is { Count: > 0 } ? string.Join(",", bccAddresses) : null,
            });

            return Ok(new { success = true, spItemId = spId });
        }
        catch (Exception ex)
        {
            _log.LogError(ex, "[Conv] Send failed to {To} for {RfqId}", req.To, req.RfqId);
            return StatusCode(500, new { error = ex.Message });
        }
    }
}
