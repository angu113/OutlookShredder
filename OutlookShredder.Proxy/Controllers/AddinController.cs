using System;
using System.Collections.Generic;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Extensions.Logging;
using OutlookShredder.Proxy.Models;
using OutlookShredder.Proxy.Services;

namespace OutlookShredder.Proxy.Controllers;

[ApiController]
[Route("api/addin")]
public class AddinController : ControllerBase
{
    private readonly RfqNotificationService        _notify;
    private readonly ILogger<AddinController>      _log;

    public AddinController(RfqNotificationService notify, ILogger<AddinController> log)
    {
        _notify = notify;
        _log    = log;
    }

    /// <summary>
    /// Office.js heartbeat — called every 15 seconds from the task pane.
    /// Returns ack + server time so the add-in can confirm proxy connectivity.
    /// </summary>
    [HttpPost("heartbeat")]
    public IActionResult Heartbeat([FromBody] HeartbeatRequest? req)
    {
        _log.LogDebug("Add-in heartbeat #{Seq} clientTime={ClientTime}", req?.Seq ?? 0, req?.ClientTime);
        return Ok(new { ack = true, serverTime = DateTimeOffset.UtcNow });
    }

    /// <summary>
    /// Receives a full email payload pushed by the VSTO add-in when a new message
    /// arrives in a monitored mailbox.  Publishes an AddinNewEmail Service Bus event
    /// so connected Shredder clients are notified in real time.
    /// Large payloads (100 MB body limit) are accepted; attachment bytes are carried
    /// in the payload for the caller to consume — the proxy does not persist them.
    /// </summary>
    [HttpPost("email-received")]
    [RequestSizeLimit(104_857_600)] // 100 MB
    public IActionResult EmailReceived([FromBody] AddinEmailPayload payload)
    {
        var attCount = payload.Attachments?.Count ?? 0;
        var attBytes = 0L;
        if (payload.Attachments != null)
            foreach (var a in payload.Attachments)
                attBytes += a.SizeBytes;

        _log.LogInformation(
            "Addin email received: From={From} Subject={Subject} Attachments={Count} TotalBytes={Bytes}",
            payload.FromAddress, payload.Subject, attCount, attBytes);

        _notify.NotifyRfqProcessed(new RfqProcessedNotification
        {
            EventType  = "AddinNewEmail",
            MessageId  = payload.InternetMessageId,
            MsgFrom    = payload.FromAddress,
            MsgTo      = payload.ToAddress,
            MsgBody    = payload.Subject,
            MsgChannel = payload.MailboxDisplayName,
            MsgDirection  = "in",
            MsgTimestamp  = payload.ReceivedAt.ToString("o")
        });

        return Ok(new { received = true, attachments = attCount });
    }

    /// <summary>
    /// Forwards a fetch-by-EntryId request to the add-in's HTTP listener (port 7002)
    /// and returns the full message payload.  The add-in must be running.
    /// </summary>
    [HttpPost("fetch")]
    public async Task<IActionResult> Fetch([FromBody] AddinFetchRequest req,
        [FromServices] IHttpClientFactory http)
    {
        using var client = http.CreateClient();
        client.Timeout = TimeSpan.FromSeconds(60);
        try
        {
            var json     = System.Text.Json.JsonSerializer.Serialize(req);
            var content  = new System.Net.Http.StringContent(json, System.Text.Encoding.UTF8, "application/json");
            var response = await client.PostAsync("http://localhost:7002/fetch", content);
            var body     = await response.Content.ReadAsStringAsync();
            return Content(body, "application/json");
        }
        catch (Exception ex)
        {
            _log.LogWarning("Addin fetch failed: {Msg}", ex.Message);
            return StatusCode(502, new { error = "Add-in unreachable", detail = ex.Message });
        }
    }

    /// <summary>
    /// Forwards a send-email request to the add-in's HTTP listener.
    /// </summary>
    [HttpPost("send")]
    public async Task<IActionResult> Send([FromBody] AddinSendRequest req,
        [FromServices] IHttpClientFactory http)
    {
        using var client = http.CreateClient();
        client.Timeout = TimeSpan.FromSeconds(60);
        try
        {
            var json     = System.Text.Json.JsonSerializer.Serialize(req);
            var content  = new System.Net.Http.StringContent(json, System.Text.Encoding.UTF8, "application/json");
            var response = await client.PostAsync("http://localhost:7002/send", content);
            var body     = await response.Content.ReadAsStringAsync();
            return Content(body, "application/json");
        }
        catch (Exception ex)
        {
            _log.LogWarning("Addin send failed: {Msg}", ex.Message);
            return StatusCode(502, new { error = "Add-in unreachable", detail = ex.Message });
        }
    }
}

// ── Payload models ──────────────────────────────────────────────────────────

public record HeartbeatRequest(int Seq, string? ClientTime);

public class AddinEmailPayload
{
    public string   EntryId            { get; set; } = string.Empty;
    public string?  StoreId            { get; set; }
    public string?  InternetMessageId  { get; set; }
    public string?  Subject            { get; set; }
    public string?  FromAddress        { get; set; }
    public string?  FromName           { get; set; }
    public string?  ToAddress          { get; set; }
    public DateTime ReceivedAt         { get; set; }
    public string?  BodyText           { get; set; }
    public string?  BodyHtml           { get; set; }
    public string?  MailboxDisplayName { get; set; }
    public List<AddinAttachment>? Attachments { get; set; }
}

public class AddinAttachment
{
    public string  FileName       { get; set; } = string.Empty;
    public string? ContentType    { get; set; }
    public int     SizeBytes      { get; set; }
    public string  ContentBase64  { get; set; } = string.Empty;
}

public class AddinFetchRequest
{
    public string  EntryId { get; set; } = string.Empty;
    public string? StoreId { get; set; }
}

public class AddinSendRequest
{
    public string? FromAccount { get; set; }
    public string? To          { get; set; }
    public string? Cc          { get; set; }
    public string? Bcc         { get; set; }
    public string? Subject     { get; set; }
    public string? BodyHtml    { get; set; }
    public string? BodyText    { get; set; }
}
