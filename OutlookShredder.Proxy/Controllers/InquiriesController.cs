using Microsoft.AspNetCore.Mvc;
using OutlookShredder.Proxy.Services;

namespace OutlookShredder.Proxy.Controllers;

/// <summary>
/// CX SMS customer-inquiry endpoints (Phase 3). Thin over <see cref="InquiryService"/> — all threading,
/// opt-out, draft, and notification logic lives in the service. (Split-to-new-thread is deferred to a
/// follow-up; see wip/customer-experience-sms-inquiry.md.)
/// </summary>
[ApiController]
[Route("api/inquiries")]
public class InquiriesController : ControllerBase
{
    private readonly InquiryService            _inquiries;
    private readonly IConfiguration            _config;
    private readonly ILogger<InquiriesController> _log;

    public InquiriesController(InquiryService inquiries, IConfiguration config, ILogger<InquiriesController> log)
    {
        _inquiries = inquiries;
        _config    = config;
        _log       = log;
    }

    [HttpGet]
    public async Task<IActionResult> List([FromQuery] string? status, [FromQuery] string? q, CancellationToken ct)
    {
        try { return Ok(await _inquiries.ListAsync(status, q, ct)); }
        catch (Exception ex) { return Fail(ex, "list"); }
    }

    /// <summary>Phase 7b: total unread inbound across active inquiries — the app-level (taskbar + Pulse-icon) badge.</summary>
    [HttpGet("unread-total")]
    public IActionResult UnreadTotal() => Ok(new { total = _inquiries.UnreadTotal() });

    /// <summary>New-SMS: does a conversation already exist for this number? Returns its id so the client opens it
    /// instead of starting a duplicate (one thread per customer).</summary>
    [HttpGet("find")]
    public async Task<IActionResult> Find([FromQuery] string phone, CancellationToken ct)
    {
        try { var id = await _inquiries.FindInquiryIdByPhoneAsync(phone, ct); return Ok(new { found = id is not null, inquiryId = id }); }
        catch (Exception ex) { return Fail(ex, "find"); }
    }

    /// <summary>New-SMS: get-or-create-or-reopen the one-thread inquiry for a number (operator-initiated, called on
    /// first send). Returns the inquiry; 400 if the phone isn't a valid US number.</summary>
    [HttpPost("start")]
    public async Task<IActionResult> Start([FromBody] StartInquiryRequest req, CancellationToken ct)
    {
        if (string.IsNullOrWhiteSpace(req?.Phone)) return BadRequest(new { error = "phone is required" });
        try
        {
            var inq = await _inquiries.StartInquiryAsync(req.Phone!, ct);
            return inq is null ? BadRequest(new { error = "Not a valid US phone number" }) : Ok(inq);
        }
        catch (Exception ex) { return Fail(ex, "start"); }
    }

    public sealed class StartInquiryRequest { public string? Phone { get; set; } }

    [HttpGet("{id}")]
    public async Task<IActionResult> Get(string id, CancellationToken ct)
    {
        try
        {
            var detail = await _inquiries.GetDetailAsync(id, ct);
            return detail is null ? NotFound() : Ok(detail);
        }
        catch (Exception ex) { return Fail(ex, "get"); }
    }

    /// <summary>Streams a stored inbound media file (image/pdf preview, or cad/file download). The proxy holds
    /// the bytes durably (app-only Graph creds) so the client never needs the carrier's auth or expiring URL.</summary>
    [HttpGet("{id}/media")]
    public async Task<IActionResult> GetMedia(string id, [FromQuery] string name, CancellationToken ct)
    {
        try
        {
            var res = await _inquiries.GetMediaAsync(id, name, ct);
            return res is null ? NotFound() : File(res.Value.Bytes, res.Value.ContentType);
        }
        catch (Exception ex) { return Fail(ex, "media"); }
    }

    /// <summary>Dev/recovery: re-pull + attach media for an existing message by SID (gated by SignalWire:AllowDevSeed).
    /// Body: { sid, mediaJson } where mediaJson is a [{url,contentType}] array of carrier media parts.</summary>
    [HttpPost("{id}/backfill-media")]
    public async Task<IActionResult> BackfillMedia(string id, [FromBody] BackfillMediaRequest? req, CancellationToken ct)
    {
        if (!_config.GetValue("SignalWire:AllowDevSeed", false)) return NotFound();
        if (string.IsNullOrWhiteSpace(req?.Sid) || string.IsNullOrWhiteSpace(req.MediaJson))
            return BadRequest(new { error = "sid and mediaJson are required" });
        try
        {
            var ok = await _inquiries.BackfillMessageMediaAsync(id, req.Sid!, req.MediaJson!, ct);
            return ok ? Ok(new { ok = true }) : NotFound(new { error = "no message matched that sid" });
        }
        catch (Exception ex) { return Fail(ex, "backfill-media"); }
    }

    public sealed class BackfillMediaRequest
    {
        public string? Sid       { get; set; }
        public string? MediaJson { get; set; }
    }

    /// <summary>Admin/one-time: rewrite an operator identity (old Windows login -> Shredder app username) across
    /// Inquiries.AssignedTo, InquiryNotes.NoteAuthor and InquiryQuotations.LinkedBy. Dry-run unless apply=true.
    /// Body: { from, to, apply }. Returns per-list matched/patched counts + affected inquiry ids.</summary>
    [HttpPost("backfill-identity")]
    public async Task<IActionResult> BackfillIdentity([FromBody] BackfillIdentityRequest? req, CancellationToken ct)
    {
        if (string.IsNullOrWhiteSpace(req?.From) || string.IsNullOrWhiteSpace(req.To))
            return BadRequest(new { error = "from and to are required" });
        try
        {
            var r = await _inquiries.BackfillIdentityAsync(req.From!, req.To!, req.Apply, ct);
            return Ok(new
            {
                applied  = req.Apply,
                from     = req.From,
                to       = req.To,
                assigned = new { matched = r.AssignedMatched, patched = r.AssignedPatched },
                notes    = new { matched = r.NotesMatched,    patched = r.NotesPatched },
                quotes   = new { matched = r.QuotesMatched,   patched = r.QuotesPatched },
                affectedInquiries = r.AffectedInquiryIds,
            });
        }
        catch (Exception ex) { return Fail(ex, "backfill-identity"); }
    }

    /// <summary>One-time migration: populate native *Dt dateTime columns from the legacy text date columns
    /// across the inquiry lists (Inquiries/Drafts/InquiryNotes/InquiryQuotations). Idempotent.</summary>
    [HttpPost("backfill-datetime-columns")]
    public async Task<IActionResult> BackfillDateTimeColumns(CancellationToken ct)
    {
        try
        {
            var (scanned, patched, failed) = await _inquiries.BackfillDateTimeColumnsAsync(ct);
            return Ok(new { scanned, patched, failed });
        }
        catch (Exception ex) { return Fail(ex, "backfill-datetime-columns"); }
    }

    public sealed class BackfillIdentityRequest
    {
        public string? From  { get; set; }
        public string? To    { get; set; }
        public bool    Apply { get; set; }
    }

    /// <summary>One-time backfill for outbound messages stuck on "queued" from before the SmsStatus callback
    /// was wired up (2026-07-01) — looks up each one's current status directly via SignalWire and applies it.
    /// dryRun=true (default) reports counts without writing.</summary>
    [HttpPost("backfill-message-statuses")]
    public async Task<IActionResult> BackfillMessageStatuses([FromQuery] bool dryRun = true, CancellationToken ct = default)
    {
        try { return Ok(await _inquiries.BackfillMessageStatusesAsync(dryRun, ct)); }
        catch (Exception ex) { return Fail(ex, "backfill-message-statuses"); }
    }

    [HttpPost("{id}/messages")]
    public async Task<IActionResult> SendMessage(string id, [FromBody] SendReplyRequest req, CancellationToken ct)
    {
        if (string.IsNullOrWhiteSpace(req.Body)) return BadRequest(new { error = "Body is required" });
        try
        {
            var msg = await _inquiries.SendOperatorReplyAsync(id, req.Body, req.FromDraftSpItemId, req.From, null, ct);
            return msg is null ? NotFound() : Ok(msg);
        }
        catch (InvalidOperationException ex) { return Conflict(new { error = ex.Message }); }   // opted out / no gateway
        catch (Exception ex) { return Fail(ex, "send"); }
    }

    /// <summary>Outbound MMS: text + image/PDF attachments (multipart/form-data; field 'files'). Images send as
    /// MMS; a PDF is rasterized to one image per page. The durable copy is kept in SharePoint (rendered in the
    /// thread). Form fields: body, from, fromDraftSpItemId, files[].</summary>
    [HttpPost("{id}/messages/mms")]
    [RequestSizeLimit(30_000_000)]
    public async Task<IActionResult> SendMms(string id, [FromForm] string? body, [FromForm] string? from,
        [FromForm] int? fromDraftSpItemId, CancellationToken ct)
    {
        try
        {
            var attachments = new List<InquiryService.OutboundAttachment>();
            foreach (var f in Request.Form.Files)
            {
                if (f.Length <= 0) continue;
                using var ms = new MemoryStream();
                await f.CopyToAsync(ms, ct);
                attachments.Add(new InquiryService.OutboundAttachment(
                    f.FileName, string.IsNullOrWhiteSpace(f.ContentType) ? "application/octet-stream" : f.ContentType, ms.ToArray()));
            }
            if (string.IsNullOrWhiteSpace(body) && attachments.Count == 0)
                return BadRequest(new { error = "Body or at least one attachment is required" });

            var msg = await _inquiries.SendOperatorReplyAsync(id, body ?? "", fromDraftSpItemId, from, attachments, ct);
            return msg is null ? NotFound() : Ok(msg);
        }
        catch (InvalidOperationException ex) { return Conflict(new { error = ex.Message }); }   // opted out / no gateway / unsupported type
        catch (Exception ex) { return Fail(ex, "send-mms"); }
    }

    [HttpPost("{id}/notes")]
    public async Task<IActionResult> AddNote(string id, [FromBody] AddNoteRequest req, CancellationToken ct)
    {
        if (string.IsNullOrWhiteSpace(req.Body)) return BadRequest(new { error = "Body is required" });
        try
        {
            var note = await _inquiries.AddNoteAsync(id, req.Author ?? "", req.Body, ct);
            return note is null ? NotFound() : Ok(note);
        }
        catch (Exception ex) { return Fail(ex, "note"); }
    }

    [HttpPost("{id}/quotations")]
    public async Task<IActionResult> LinkQuotation(string id, [FromBody] LinkQuotationRequest req, CancellationToken ct)
    {
        if (!InquiryRules.IsValidHsk(req.HskNumber))
            return BadRequest(new { error = "HskNumber must look like (HSK-)?(SO|PO|Q)<digits>, e.g. SO1036432" });
        try
        {
            var q = await _inquiries.LinkQuotationAsync(id, req.HskNumber!, req.LinkedBy ?? "", ct);
            return q is null ? NotFound() : Ok(q);
        }
        catch (Exception ex) { return Fail(ex, "quotation"); }
    }

    [HttpPatch("{id}")]
    public async Task<IActionResult> Patch(string id, [FromBody] PatchInquiryRequest req, CancellationToken ct)
    {
        try
        {
            var inquiry = await _inquiries.UpdateInquiryAsync(id, req.Status, req.AssignedTo, ct);
            return inquiry is null ? NotFound() : Ok(inquiry);
        }
        catch (Exception ex) { return Fail(ex, "patch"); }
    }

    [HttpPost("{id}/read")]
    public async Task<IActionResult> MarkRead(string id, CancellationToken ct)
    {
        try
        {
            var inquiry = await _inquiries.MarkReadAsync(id, ct);
            return inquiry is null ? NotFound() : Ok(inquiry);
        }
        catch (Exception ex) { return Fail(ex, "read"); }
    }

    /// <summary>Phase 7: set one message's read flag (per-message Mark Read/Unread toggle). Body: { read }.</summary>
    [HttpPost("{id}/messages/{messageSpItemId:int}/read")]
    public async Task<IActionResult> SetMessageRead(string id, int messageSpItemId, [FromBody] ReadFlagRequest? req, CancellationToken ct)
    {
        try
        {
            var inquiry = await _inquiries.SetMessageReadAsync(id, messageSpItemId, req?.Read ?? true, ct);
            return inquiry is null ? NotFound() : Ok(inquiry);
        }
        catch (Exception ex) { return Fail(ex, "msg-read"); }
    }

    /// <summary>Phase 7: mark every message in the inquiry read/unread (Mark all read / Mark all unread). Body: { read }.</summary>
    [HttpPost("{id}/read-all")]
    public async Task<IActionResult> MarkAll(string id, [FromBody] ReadFlagRequest? req, CancellationToken ct)
    {
        try
        {
            var inquiry = await _inquiries.MarkAllAsync(id, req?.Read ?? true, ct);
            return inquiry is null ? NotFound() : Ok(inquiry);
        }
        catch (Exception ex) { return Fail(ex, "read-all"); }
    }

    /// <summary>Operator: regenerate the AI suggestion for the latest inbound (dismisses prior pending drafts).</summary>
    [HttpPost("{id}/regenerate-draft")]
    public async Task<IActionResult> RegenerateDraft(string id, CancellationToken ct)
    {
        try
        {
            var ok = await _inquiries.RegenerateDraftAsync(id, ct);
            return ok ? Ok(new { ok = true }) : NotFound(new { error = "no inbound message to draft from" });
        }
        catch (Exception ex) { return Fail(ex, "regenerate-draft"); }
    }

    [HttpPost("{id}/drafts/{draftId:int}/accept")]
    public async Task<IActionResult> AcceptDraft(string id, int draftId, [FromBody] AcceptDraftRequest? req, CancellationToken ct)
    {
        try
        {
            var msg = await _inquiries.AcceptDraftAsync(id, draftId, req?.From, ct);
            return msg is null ? NotFound() : Ok(msg);
        }
        catch (InvalidOperationException ex) { return Conflict(new { error = ex.Message }); }
        catch (Exception ex) { return Fail(ex, "accept-draft"); }
    }

    [HttpPost("{id}/drafts/{draftId:int}/dismiss")]
    public async Task<IActionResult> DismissDraft(string id, int draftId, CancellationToken ct)
    {
        try { await _inquiries.DismissDraftAsync(id, draftId, ct); return Ok(); }
        catch (Exception ex) { return Fail(ex, "dismiss-draft"); }
    }

    private IActionResult Fail(Exception ex, string op)
    {
        _log.LogWarning(ex, "[Inquiries] {Op} failed", op);
        return StatusCode(500, new { error = ex.Message });
    }

    public sealed class SendReplyRequest     { public string? Body { get; set; } public int? FromDraftSpItemId { get; set; } public string? From { get; set; } }
    public sealed class AcceptDraftRequest   { public string? From { get; set; } }
    public sealed class ReadFlagRequest      { public bool Read { get; set; } }
    public sealed class AddNoteRequest       { public string? Author { get; set; } public string? Body { get; set; } }
    public sealed class LinkQuotationRequest { public string? HskNumber { get; set; } public string? LinkedBy { get; set; } }
    public sealed class PatchInquiryRequest  { public string? Status { get; set; } public string? AssignedTo { get; set; } }
}
