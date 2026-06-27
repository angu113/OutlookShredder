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
            bool ok;
            if (string.Equals(req.Channel, "email", StringComparison.OrdinalIgnoreCase))
            {
                if (string.IsNullOrWhiteSpace(req.Subject))
                    return BadRequest(new { error = "Subject is required for email channel" });
                ok = await _messaging.SendEmailAsync(req.From, req.To, req.Subject!, req.Body, ct);
            }
            else if (string.Equals(req.Channel, "sms", StringComparison.OrdinalIgnoreCase))
            {
                ok = await _messaging.SendSmsAsync(req.From, req.To, req.Body, ct);
            }
            else
            {
                ok = await _messaging.SendInternalAsync(req.From, req.To, req.Body, req.Subject, ct);
            }
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
        try   { await _messaging.MarkReadAsync(conversationId, ct); return Ok(); }
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
    private readonly SmsInboundQueue               _queue;
    private readonly IConfiguration                _config;
    private readonly ILogger<SmsWebhookController> _log;

    public SmsWebhookController(SmsInboundQueue queue, IConfiguration config, ILogger<SmsWebhookController> log)
    {
        _queue  = queue;
        _config = config;
        _log    = log;
    }

    /// <summary>
    /// SignalWire inbound SMS webhook. Any proxy may receive it (Cloudflare routes to any healthy tunnel
    /// replica); we validate the signature, ack fast, and enqueue to the dedup queue so exactly one proxy
    /// processes it. Returns empty TwiML so SignalWire does not attempt a reply.
    /// Wire in the SignalWire dashboard: POST https://&lt;public-host&gt;/api/sms/inbound
    /// </summary>
    [HttpPost("inbound")]
    public async Task<IActionResult> Inbound()
    {
        var form = await Request.ReadFormAsync();
        if (!ValidateSignature(form)) return StatusCode(403);

        var from  = form["From"].ToString();
        var to    = form["To"].ToString();
        var body  = form["Body"].ToString();
        var sid   = form["MessageSid"].ToString();
        var media = CollectMediaUrls(form);

        if (!string.IsNullOrWhiteSpace(from) && (!string.IsNullOrWhiteSpace(body) || media is not null))
        {
            _log.LogInformation("[SMS] inbound from {From} sid {Sid}", from, sid);
            try
            {
                await _queue.EnqueueAsync(new SmsInboundQueue.Job(
                    from, to, body, string.IsNullOrEmpty(sid) ? null : sid, media));
            }
            catch (Exception ex) { _log.LogWarning(ex, "[SMS] enqueue inbound failed"); }
        }
        return Content("<Response/>", "application/xml");
    }

    /// <summary>
    /// SignalWire delivery-status callback for outbound messages. Validates the signature; Phase 1 updates
    /// the stored message status by SID. Wire as the outbound statusCallback / number status URL.
    /// </summary>
    [HttpPost("status")]
    public async Task<IActionResult> Status()
    {
        var form = await Request.ReadFormAsync();
        if (!ValidateSignature(form)) return StatusCode(403);

        var sid    = form["MessageSid"].ToString();
        var status = form["MessageStatus"].ToString();
        _log.LogInformation("[SMS] status sid {Sid} -> {Status}", sid, status);
        // Phase 1: persist status to the message row by SID.
        return Content("<Response/>", "application/xml");
    }

    /// <summary>Validates the SignalWire HMAC-SHA1 webhook signature. Fails closed when no token is
    /// configured unless SignalWire:AllowUnsignedWebhooks=true (local dev only).</summary>
    private bool ValidateSignature(IFormCollection form)
    {
        var token = _config["SignalWire:SigningKey"] ?? _config["SignalWire:ApiToken"];
        if (string.IsNullOrWhiteSpace(token))
        {
            if (_config.GetValue("SignalWire:AllowUnsignedWebhooks", false)) return true;
            _log.LogWarning("[SMS] no SignalWire token configured — rejecting webhook");
            return false;
        }

        var sig = Request.Headers["X-SignalWire-Signature"].FirstOrDefault()
               ?? Request.Headers["X-Twilio-Signature"].FirstOrDefault();
        var url    = BuildPublicUrl();
        var prms   = form.Select(kv => new KeyValuePair<string, string>(kv.Key, kv.Value.ToString()));
        var ok     = SignalWireSignatureValidator.IsValid(url, prms, sig, token);
        if (!ok) _log.LogWarning("[SMS] webhook signature rejected (url={Url})", url);
        return ok;
    }

    /// <summary>The PUBLIC url SignalWire signed — a configured base wins (deterministic, spoof-proof);
    /// else reconstruct from forwarded headers (Cloudflare Tunnel sets X-Forwarded-Proto + Host).</summary>
    private string BuildPublicUrl()
    {
        var configured = _config["SignalWire:WebhookBaseUrl"];
        if (!string.IsNullOrWhiteSpace(configured))
            return configured.TrimEnd('/') + Request.Path.ToString();

        var proto = Request.Headers["X-Forwarded-Proto"].FirstOrDefault() ?? "https";
        var host  = Request.Headers["X-Forwarded-Host"].FirstOrDefault()
                 ?? (Request.Host.HasValue ? Request.Host.Value : "");
        return $"{proto}://{host}{Request.Path}";
    }

    private static string? CollectMediaUrls(IFormCollection form)
    {
        if (!int.TryParse(form["NumMedia"].ToString(), out var n) || n <= 0) return null;
        var urls = new List<string>();
        for (int i = 0; i < n; i++)
        {
            var u = form[$"MediaUrl{i}"].ToString();
            if (!string.IsNullOrWhiteSpace(u)) urls.Add(u);
        }
        return urls.Count > 0 ? System.Text.Json.JsonSerializer.Serialize(urls) : null;
    }
}
