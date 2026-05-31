using Microsoft.AspNetCore.Mvc;
using OutlookShredder.Proxy.Services;

namespace OutlookShredder.Proxy.Controllers;

/// <summary>
/// Read-side HTTP surface for the mailbox bridge (wip/mailbox-bridge.md, Phase 1).
/// All routes key on the WATCHED mailbox upn (e.g. hackensack@metalsupermarkets.com);
/// the bridge maps that to the destination mirror folder internally.
/// Outbound send (POST .../send) is Phase 1.1 and intentionally not implemented yet.
/// </summary>
[ApiController]
[Route("api/mailbox")]
public sealed class MailboxController : ControllerBase
{
    private readonly MailboxBridgeService _bridge;

    public MailboxController(MailboxBridgeService bridge) => _bridge = bridge;

    /// <summary>List configured watched mailboxes + per-mailbox poll state.</summary>
    [HttpGet]
    public IActionResult GetMailboxes() => Ok(_bridge.GetStatuses());

    /// <summary>Folder tree for a watched mailbox (v1: a single Inbox node).</summary>
    [HttpGet("{upn}/folders")]
    public IActionResult GetFolders(string upn)
    {
        var folders = _bridge.GetFolders(upn);
        return folders is null ? NotFoundUpn(upn) : Ok(folders);
    }

    /// <summary>Recent messages in a folder (v1 ignores folder id — only the mirror Inbox exists).</summary>
    [HttpGet("{upn}/folders/{folder}/messages")]
    public IActionResult GetMessages(string upn, string folder, [FromQuery] int top = 100)
    {
        var msgs = _bridge.GetMessages(upn, Math.Clamp(top, 1, 250));
        return msgs is null ? NotFoundUpn(upn) : Ok(msgs);
    }

    /// <summary>Full message detail (original headers + plain-text body + attachment metadata).</summary>
    [HttpGet("{upn}/messages/{id}")]
    public IActionResult GetMessage(string upn, string id)
    {
        if (!_bridge.IsWatched(upn)) return NotFoundUpn(upn);
        var body = _bridge.GetMessage(upn, id);
        return body is null ? NotFound(new { error = "Message not in cache (may not have been polled yet)" }) : Ok(body);
    }

    /// <summary>Attachment bytes from the embedded original message.</summary>
    [HttpGet("{upn}/messages/{id}/attachments/{name}")]
    public async Task<IActionResult> GetAttachment(string upn, string id, string name, CancellationToken ct)
    {
        if (!_bridge.IsWatched(upn)) return NotFoundUpn(upn);
        var att = await _bridge.GetAttachmentAsync(upn, id, name, ct);
        if (att is null) return NotFound(new { error = $"Attachment '{name}' not found" });
        return File(att.Value.Bytes, att.Value.ContentType, att.Value.FileName);
    }

    /// <summary>Mark a message read/unread on the mirror copy.</summary>
    [HttpPost("{upn}/messages/{id}/mark-read")]
    public async Task<IActionResult> MarkRead(string upn, string id, [FromBody] MarkReadRequest req, CancellationToken ct)
    {
        var ok = await _bridge.SetReadAsync(upn, id, req?.Read ?? true, ct);
        return ok ? Ok(new { success = true }) : NotFoundUpn(upn);
    }

    private IActionResult NotFoundUpn(string upn) =>
        NotFound(new { error = $"Mailbox '{upn}' is not a configured watched mailbox" });

    public sealed class MarkReadRequest { public bool Read { get; set; } = true; }
}
