using Microsoft.AspNetCore.Mvc;
using OutlookShredder.Proxy.Models;
using OutlookShredder.Proxy.Services;

namespace OutlookShredder.Proxy.Controllers;

[ApiController]
[Route("api/messages")]
public class MessagingController : ControllerBase
{
    private readonly MessagingService            _messaging;
    private readonly SharePointService           _sp;
    private readonly ILogger<MessagingController> _log;

    public MessagingController(MessagingService messaging, SharePointService sp,
        ILogger<MessagingController> log)
    {
        _messaging = messaging;
        _sp        = sp;
        _log       = log;
    }

    [HttpGet("conversations")]
    public async Task<IActionResult> GetConversations([FromQuery] int top = 200, CancellationToken ct = default)
    {
        try   { return Ok(await _sp.GetConversationSummariesAsync(top, ct)); }
        catch (Exception ex)
        {
            _log.LogWarning(ex, "[Messaging] GetConversations failed");
            return StatusCode(500, new { error = ex.Message });
        }
    }

    [HttpGet("conversation/{conversationId}")]
    public async Task<IActionResult> GetConversation(string conversationId,
        [FromQuery] int top = 50, CancellationToken ct = default)
    {
        try   { return Ok(await _sp.GetConversationMessagesAsync(conversationId, top, ct)); }
        catch (Exception ex)
        {
            _log.LogWarning(ex, "[Messaging] GetConversation failed for '{Id}'", conversationId);
            return StatusCode(500, new { error = ex.Message });
        }
    }

    [HttpPost("send")]
    public async Task<IActionResult> Send([FromBody] SendMessageRequest req, CancellationToken ct = default)
    {
        if (string.IsNullOrWhiteSpace(req.To) || string.IsNullOrWhiteSpace(req.Body))
            return BadRequest(new { error = "To and Body are required" });

        try
        {
            var ok = string.Equals(req.Channel, "sms", StringComparison.OrdinalIgnoreCase)
                ? await _messaging.SendSmsAsync(req.From, req.To, req.Body, ct)
                : await _messaging.SendInternalAsync(req.From, req.To, req.Body, ct);
            return Ok(new { ok });
        }
        catch (Exception ex)
        {
            _log.LogWarning(ex, "[Messaging] Send failed");
            return StatusCode(500, new { error = ex.Message });
        }
    }

    [HttpPost("read/{conversationId}")]
    public async Task<IActionResult> MarkRead(string conversationId, CancellationToken ct = default)
    {
        try   { await _sp.MarkConversationReadAsync(conversationId, ct); return Ok(); }
        catch (Exception ex)
        {
            _log.LogWarning(ex, "[Messaging] MarkRead failed for '{Id}'", conversationId);
            return StatusCode(500, new { error = ex.Message });
        }
    }

    [HttpGet("users")]
    public IActionResult GetUsers() => Ok(_messaging.KnownUsers);
}

[ApiController]
[Route("api/sms")]
public class SmsWebhookController : ControllerBase
{
    private readonly MessagingService              _messaging;
    private readonly ILogger<SmsWebhookController> _log;

    public SmsWebhookController(MessagingService messaging, ILogger<SmsWebhookController> log)
    {
        _messaging = messaging;
        _log       = log;
    }

    /// <summary>
    /// SignalWire inbound SMS webhook. Returns empty TwiML so SignalWire does not attempt a reply.
    /// Wire this URL in the SignalWire dashboard: POST https://your-proxy/api/sms/inbound
    /// </summary>
    [HttpPost("inbound")]
    public IActionResult Inbound([FromForm] string? From, [FromForm] string? To,
        [FromForm] string? Body, [FromForm] string? MessageSid)
    {
        if (!string.IsNullOrWhiteSpace(From) && !string.IsNullOrWhiteSpace(Body))
        {
            _log.LogInformation("[SMS] Inbound from {From}: {Preview}",
                From, Body.Length > 80 ? Body[..80] + "…" : Body);
            _ = Task.Run(async () =>
            {
                try { await _messaging.HandleInboundSmsAsync(From, To ?? "", Body, MessageSid); }
                catch (Exception ex) { _log.LogWarning(ex, "[SMS] HandleInboundSms failed"); }
            });
        }
        return Content("<Response/>", "application/xml");
    }
}
