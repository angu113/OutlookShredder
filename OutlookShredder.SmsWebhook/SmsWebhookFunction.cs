using System.Net;
using System.Text;
using System.Text.Json;
using Azure.Messaging.ServiceBus;
using Microsoft.Azure.Functions.Worker;
using Microsoft.Azure.Functions.Worker.Http;
using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.Logging;

namespace OutlookShredder.SmsWebhook;

/// <summary>
/// Public SMS ingress (Azure Function). Validates the SignalWire HMAC signature, then enqueues the inbound
/// message onto the existing <c>sms-inbound-jobs</c> Service Bus queue (MessageId = MessageSid for dedup) —
/// the desktop proxies' <c>SmsInboundQueueProcessor</c> (competing consumers) pick it up and run the full
/// inquiry pipeline. Nothing here touches SharePoint / AI; it's a thin, always-on receiver that decouples
/// ingress from the proxies (inbound survives proxies being down). The JSON payload mirrors the proxy's
/// <c>SmsInboundQueue.Job(From, To, Body, Sid, MediaUrls)</c> — keep the two in sync.
/// </summary>
public class SmsWebhookFunction
{
    private readonly ServiceBusClient _sb;
    private readonly ServiceBusClient _sbStatus;
    private readonly IConfiguration   _config;
    private readonly ILogger<SmsWebhookFunction> _log;

    private static readonly JsonSerializerOptions _json = new() { PropertyNamingPolicy = JsonNamingPolicy.CamelCase };

    public SmsWebhookFunction(ServiceBusClient sb, StatusServiceBusClient sbStatus,
        IConfiguration config, ILogger<SmsWebhookFunction> log)
    {
        _sb       = sb;
        _sbStatus = sbStatus.Client;
        _config   = config;
        _log      = log;
    }

    /// <summary>Inbound SMS webhook. Wire in SignalWire: POST https://&lt;func&gt;.azurewebsites.net/api/sms/inbound</summary>
    [Function("SmsInbound")]
    public async Task<HttpResponseData> Inbound(
        [HttpTrigger(AuthorizationLevel.Anonymous, "post", Route = "sms/inbound")] HttpRequestData req)
    {
        var form = await ReadFormAsync(req);
        if (!ValidateRequest(req, form)) return req.CreateResponse(HttpStatusCode.Forbidden);

        var from  = Get(form, "From");
        var to    = Get(form, "To");
        var body  = Get(form, "Body");
        var sid   = Get(form, "MessageSid");
        var media = CollectMediaUrls(form);

        if (!string.IsNullOrWhiteSpace(from) && (!string.IsNullOrWhiteSpace(body) || media is not null))
        {
            _log.LogInformation("[SmsWebhook] inbound from {From} sid {Sid}", from, sid);
            try { await EnqueueAsync(from, to, body, string.IsNullOrEmpty(sid) ? null : sid, media); }
            catch (Exception ex) { _log.LogError(ex, "[SmsWebhook] enqueue failed"); }
        }
        return Twiml(req);
    }

    /// <summary>Outbound delivery-status callback. Validates + acks, then enqueues onto <c>sms-status-jobs</c>
    /// — this Function has no SharePoint access, so applying the status (the Messages row + a proxy's
    /// in-memory cache) is deferred to a proxy consumer (<c>SmsStatusQueueProcessor</c>).</summary>
    [Function("SmsStatus")]
    public async Task<HttpResponseData> Status(
        [HttpTrigger(AuthorizationLevel.Anonymous, "post", Route = "sms/status")] HttpRequestData req)
    {
        var form = await ReadFormAsync(req);
        if (!ValidateRequest(req, form)) return req.CreateResponse(HttpStatusCode.Forbidden);

        var sid    = Get(form, "MessageSid");
        var status = Get(form, "MessageStatus");
        _log.LogInformation("[SmsWebhook] status sid {Sid} -> {Status}", sid, status);
        if (!string.IsNullOrWhiteSpace(sid) && !string.IsNullOrWhiteSpace(status))
        {
            try { await EnqueueStatusAsync(sid, status); }
            catch (Exception ex) { _log.LogError(ex, "[SmsWebhook] status enqueue failed"); }
        }
        return Twiml(req);
    }

    /// <summary>Keep-warm: a cheap timer tick every 5 min keeps the Consumption instance + isolated worker
    /// hot, so the FIRST inbound SMS after an idle period doesn't cold-start past SignalWire's ~15s webhook
    /// timeout (the cause of the early "11200 HTTP retrieval failure"). Tiny + within the monthly free grant.</summary>
    [Function("KeepWarm")]
    public void KeepWarm([TimerTrigger("0 */5 * * * *")] TimerInfo timer)
        => _log.LogInformation("[SmsWebhook] keep-warm tick");

    private async Task EnqueueAsync(string from, string to, string body, string? sid, string? mediaUrls)
    {
        var queue = _config["ServiceBus:InboundQueueName"] ?? "sms-inbound-jobs";
        await using var sender = _sb.CreateSender(queue);
        var payload = JsonSerializer.Serialize(new { from, to, body, sid, mediaUrls }, _json);
        var msg = new ServiceBusMessage(payload) { ContentType = "application/json" };
        if (!string.IsNullOrWhiteSpace(sid)) msg.MessageId = sid;   // duplicate-detection key (the proxy queue dedups on this)
        await sender.SendMessageAsync(msg);
    }

    // No MessageId/dedup here (unlike EnqueueAsync above): a status callback legitimately repeats for the
    // SAME sid with a DIFFERENT status (queued -> sent -> delivered), so a sid-keyed dedup would collapse
    // real, distinct events. The queue (sms-status-jobs) is provisioned by the proxy without duplicate
    // detection to match. If the queue doesn't exist yet (proxy hasn't started since this was wired), the
    // send throws and the caller logs + drops it — never fatal to the webhook ack.
    private async Task EnqueueStatusAsync(string sid, string status)
    {
        var queue = _config["ServiceBus:SmsStatusQueueName"] ?? "sms-status-jobs";
        await using var sender = _sbStatus.CreateSender(queue);
        var payload = JsonSerializer.Serialize(new { sid, status }, _json);
        var msg = new ServiceBusMessage(payload) { ContentType = "application/json" };
        await sender.SendMessageAsync(msg);
    }

    /// <summary>
    /// Authenticates the inbound webhook. PRIMARY: an unguessable shared secret in the URL (<c>?t=&lt;secret&gt;</c>,
    /// app setting <c>SignalWire:WebhookSecret</c>) — SignalWire Call Fabric's cXML webhook does not emit a
    /// Twilio-style signature we can validate, so the secret URL is the gate (HTTPS in transit; rotate by
    /// changing the app setting + the SignalWire resource URL). FALLBACK (when no secret is set): the
    /// Twilio-compatible HMAC signature.
    /// </summary>
    private bool ValidateRequest(HttpRequestData req, IReadOnlyDictionary<string, string> form)
    {
        var secret = _config["SignalWire:WebhookSecret"];
        if (!string.IsNullOrWhiteSpace(secret))
        {
            var t  = QueryParam(req.Url, "t") ?? "";
            var ok = System.Security.Cryptography.CryptographicOperations.FixedTimeEquals(
                Encoding.UTF8.GetBytes(t), Encoding.UTF8.GetBytes(secret));
            if (!ok) _log.LogWarning("[SmsWebhook] URL secret missing/mismatch");
            return ok;
        }

        var token = _config["SignalWire:SigningToken"] ?? _config["SignalWire:ApiToken"];
        if (string.IsNullOrWhiteSpace(token)) { _log.LogWarning("[SmsWebhook] no auth configured — rejecting"); return false; }
        var sig = Header(req, "X-SignalWire-Signature") ?? Header(req, "X-Twilio-Signature");
        var url = BuildPublicUrl(req);
        var valid = SignalWireSignatureValidator.IsValid(url, form, sig, token);
        if (!valid) _log.LogWarning("[SmsWebhook] signature rejected (url={Url})", url);
        return valid;
    }

    private static string? QueryParam(Uri url, string key)
    {
        foreach (var pair in url.Query.TrimStart('?').Split('&', StringSplitOptions.RemoveEmptyEntries))
        {
            var i = pair.IndexOf('=');
            var k = Uri.UnescapeDataString(i < 0 ? pair : pair[..i]);
            if (string.Equals(k, key, StringComparison.Ordinal))
                return Uri.UnescapeDataString(i < 0 ? "" : pair[(i + 1)..]);
        }
        return null;
    }

    /// <summary>The PUBLIC url SignalWire signed. A configured base wins (deterministic); else reconstruct from
    /// the request (Azure Functions sees the real public host).</summary>
    private string BuildPublicUrl(HttpRequestData req)
    {
        var path       = req.Url.AbsolutePath;   // e.g. /api/sms/inbound
        var configured = _config["SignalWire:WebhookBaseUrl"];
        return !string.IsNullOrWhiteSpace(configured)
            ? configured.TrimEnd('/') + path
            : $"{req.Url.Scheme}://{req.Url.Authority}{path}";
    }

    // ── helpers ────────────────────────────────────────────────────────────────
    private static async Task<Dictionary<string, string>> ReadFormAsync(HttpRequestData req)
    {
        using var reader = new StreamReader(req.Body, Encoding.UTF8);
        return ParseForm(await reader.ReadToEndAsync());
    }

    private static Dictionary<string, string> ParseForm(string body)
    {
        var d = new Dictionary<string, string>(StringComparer.Ordinal);
        if (string.IsNullOrEmpty(body)) return d;
        foreach (var pair in body.Split('&', StringSplitOptions.RemoveEmptyEntries))
        {
            var i = pair.IndexOf('=');
            var k = i < 0 ? pair : pair[..i];
            var v = i < 0 ? "" : pair[(i + 1)..];
            d[Decode(k)] = Decode(v);
        }
        return d;
    }

    private static string Decode(string s) => Uri.UnescapeDataString(s.Replace('+', ' '));
    private static string Get(IReadOnlyDictionary<string, string> f, string k) => f.TryGetValue(k, out var v) ? v : "";

    // Forwards each MMS media part as {url, contentType}. The proxy downloads them at ingest (carrier auth),
    // stores them durably, promotes a text/plain caption to the body, and drops the SMIL layout part.
    private static string? CollectMediaUrls(IReadOnlyDictionary<string, string> form)
    {
        if (!int.TryParse(Get(form, "NumMedia"), out var n) || n <= 0) return null;
        var parts = new List<object>();
        for (int i = 0; i < n; i++)
        {
            var u = Get(form, $"MediaUrl{i}");
            if (string.IsNullOrWhiteSpace(u)) continue;
            parts.Add(new { url = u, contentType = Get(form, $"MediaContentType{i}") });
        }
        return parts.Count > 0 ? JsonSerializer.Serialize(parts) : null;
    }

    private static string? Header(HttpRequestData req, string name)
        => req.Headers.TryGetValues(name, out var vals) ? vals.FirstOrDefault() : null;

    private static HttpResponseData Twiml(HttpRequestData req)
    {
        var r = req.CreateResponse(HttpStatusCode.OK);
        r.Headers.Add("Content-Type", "application/xml");
        r.WriteString("<Response/>");
        return r;
    }
}
