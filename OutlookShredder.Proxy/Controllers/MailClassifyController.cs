using Microsoft.AspNetCore.Mvc;
using OutlookShredder.Proxy.Models;
using OutlookShredder.Proxy.Services;

namespace OutlookShredder.Proxy.Controllers;

/// <summary>
/// Phase 1a of the mail workbench (wip/mail-classification.md): validate classification quality
/// on real forwarded mail BEFORE building SP persistence. Preview only — no writes.
/// </summary>
[ApiController]
[Route("api/mail-classify")]
public sealed class MailClassifyController : ControllerBase
{
    private readonly MailboxBridgeService _bridge;
    private readonly MailClassifierService _classifier;

    public MailClassifyController(MailboxBridgeService bridge, MailClassifierService classifier)
    {
        _bridge     = bridge;
        _classifier = classifier;
    }

    /// <summary>The taxonomy the classifier targets (for UI/inspection).</summary>
    [HttpGet("taxonomy")]
    public IActionResult GetTaxonomy() =>
        Ok(MailTaxonomy.Leaves.Select(l => new { l.Top, l.Sub, l.Path, l.Description }));

    /// <summary>
    /// Classify a message without persisting. Provide either { upn?, id } to pull a cached mirror
    /// message, or raw { subject, from, body, attachmentNames } for ad-hoc testing.
    /// </summary>
    [HttpPost("preview")]
    public async Task<IActionResult> Preview([FromBody] PreviewRequest req, CancellationToken ct)
    {
        MailClassifyInput input;

        if (!string.IsNullOrWhiteSpace(req.Id))
        {
            var upn = string.IsNullOrWhiteSpace(req.Upn)
                ? _bridge.GetStatuses().FirstOrDefault()?.WatchedUpn
                : req.Upn;
            if (string.IsNullOrWhiteSpace(upn)) return BadRequest(new { error = "No watched mailbox configured." });

            var body = _bridge.GetMessage(upn, req.Id);
            if (body is null) return NotFound(new { error = "Message not in cache (may not have been polled yet)." });

            input = new MailClassifyInput
            {
                Subject         = body.Subject,
                FromAddress     = body.FromAddress,
                FromName        = body.FromName,
                ToLine          = body.ToLine,
                BodyText        = body.BodyText,
                AttachmentNames = body.Attachments.Select(a => a.Name).ToList(),
            };
        }
        else
        {
            input = new MailClassifyInput
            {
                Subject         = req.Subject ?? "",
                FromAddress     = req.From ?? "",
                FromName        = req.FromName ?? "",
                ToLine          = req.ToLine ?? "",
                BodyText        = req.Body ?? "",
                AttachmentNames = req.AttachmentNames ?? [],
            };
            if (string.IsNullOrWhiteSpace(input.Subject) && string.IsNullOrWhiteSpace(input.BodyText))
                return BadRequest(new { error = "Provide an id, or a subject/body to classify." });
        }

        var result = await _classifier.ClassifyAsync(input, ct);
        if (result is null) return StatusCode(503, new { error = "Both AI providers unavailable or failed." });

        return Ok(new
        {
            input = new { input.Subject, input.FromAddress, input.ToLine, attachments = input.AttachmentNames },
            classification = result,
        });
    }

    public sealed class PreviewRequest
    {
        public string? Upn { get; set; }
        public string? Id  { get; set; }
        public string? Subject  { get; set; }
        public string? From     { get; set; }
        public string? FromName { get; set; }
        public string? ToLine   { get; set; }
        public string? Body     { get; set; }
        public List<string>? AttachmentNames { get; set; }
    }
}
