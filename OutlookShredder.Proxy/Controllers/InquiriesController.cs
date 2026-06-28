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
    private readonly ILogger<InquiriesController> _log;

    public InquiriesController(InquiryService inquiries, ILogger<InquiriesController> log)
    {
        _inquiries = inquiries;
        _log       = log;
    }

    [HttpGet]
    public async Task<IActionResult> List([FromQuery] string? status, [FromQuery] string? q, CancellationToken ct)
    {
        try { return Ok(await _inquiries.ListAsync(status, q, ct)); }
        catch (Exception ex) { return Fail(ex, "list"); }
    }

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

    [HttpPost("{id}/messages")]
    public async Task<IActionResult> SendMessage(string id, [FromBody] SendReplyRequest req, CancellationToken ct)
    {
        if (string.IsNullOrWhiteSpace(req.Body)) return BadRequest(new { error = "Body is required" });
        try
        {
            var msg = await _inquiries.SendOperatorReplyAsync(id, req.Body, req.FromDraftSpItemId, req.From, ct);
            return msg is null ? NotFound() : Ok(msg);
        }
        catch (InvalidOperationException ex) { return Conflict(new { error = ex.Message }); }   // opted out / no gateway
        catch (Exception ex) { return Fail(ex, "send"); }
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
        try { await _inquiries.DismissDraftAsync(draftId, ct); return Ok(); }
        catch (Exception ex) { return Fail(ex, "dismiss-draft"); }
    }

    private IActionResult Fail(Exception ex, string op)
    {
        _log.LogWarning(ex, "[Inquiries] {Op} failed", op);
        return StatusCode(500, new { error = ex.Message });
    }

    public sealed class SendReplyRequest     { public string? Body { get; set; } public int? FromDraftSpItemId { get; set; } public string? From { get; set; } }
    public sealed class AcceptDraftRequest   { public string? From { get; set; } }
    public sealed class AddNoteRequest       { public string? Author { get; set; } public string? Body { get; set; } }
    public sealed class LinkQuotationRequest { public string? HskNumber { get; set; } public string? LinkedBy { get; set; } }
    public sealed class PatchInquiryRequest  { public string? Status { get; set; } public string? AssignedTo { get; set; } }
}
