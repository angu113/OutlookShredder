using System.Net.Http.Headers;
using System.Text;
using System.Text.Json;

namespace OutlookShredder.Proxy.Services;

public class SignalWireService
{
    private readonly IConfiguration             _config;
    private readonly HttpClient                 _http;
    private readonly ILogger<SignalWireService> _log;

    public SignalWireService(IConfiguration config, ILogger<SignalWireService> log)
    {
        _config = config;
        _log    = log;
        _http   = new HttpClient { Timeout = TimeSpan.FromSeconds(30) };
    }

    public bool IsConfigured =>
        !string.IsNullOrWhiteSpace(_config["SignalWire:ProjectId"])  &&
        !string.IsNullOrWhiteSpace(_config["SignalWire:ApiToken"])   &&
        !string.IsNullOrWhiteSpace(_config["SignalWire:FromNumber"]) &&
        !string.IsNullOrWhiteSpace(_config["SignalWire:SpaceUrl"]);

    public string? FromNumber => _config["SignalWire:FromNumber"];

    public async Task<string?> SendSmsAsync(string to, string body, CancellationToken ct = default)
        => await SendSmsAsync(to, body, null, null, ct);

    public async Task<string?> SendSmsAsync(string to, string body, string? statusCallback, CancellationToken ct = default)
        => await SendSmsAsync(to, body, statusCallback, null, ct);

    /// <summary>Sends an SMS/MMS; when <paramref name="statusCallback"/> is set SignalWire POSTs delivery-status
    /// updates there (our /api/sms/status). Supplying <paramref name="mediaUrls"/> (publicly-fetchable URLs —
    /// SignalWire downloads them) sends an MMS (SendAsMms=true; up to 8). Returns the MessageSid, or null on
    /// failure.</summary>
    public async Task<string?> SendSmsAsync(string to, string body, string? statusCallback,
        IReadOnlyList<string>? mediaUrls, CancellationToken ct = default)
    {
        var projectId  = _config["SignalWire:ProjectId"]!;
        var apiToken   = _config["SignalWire:ApiToken"]!;
        var fromNumber = _config["SignalWire:FromNumber"]!;
        var spaceUrl   = _config["SignalWire:SpaceUrl"]!.TrimEnd('/');

        var url  = $"https://{spaceUrl}/api/laml/2010-04-01/Accounts/{projectId}/Messages.json";
        var auth = Convert.ToBase64String(Encoding.ASCII.GetBytes($"{projectId}:{apiToken}"));

        // A List (not Dictionary) so MediaUrl can repeat — the LaML API accepts multiple MediaUrl form fields.
        var form = new List<KeyValuePair<string, string>>
        {
            new("From", fromNumber),
            new("To",   to),
        };
        if (!string.IsNullOrEmpty(body)) form.Add(new("Body", body));   // MMS may be media-only (empty body)
        if (mediaUrls is { Count: > 0 })
        {
            foreach (var u in mediaUrls)
                if (!string.IsNullOrWhiteSpace(u)) form.Add(new("MediaUrl", u));
            form.Add(new("SendAsMms", "true"));
        }
        if (!string.IsNullOrWhiteSpace(statusCallback)) form.Add(new("StatusCallback", statusCallback));

        using var req = new HttpRequestMessage(HttpMethod.Post, url);
        req.Headers.Authorization = new AuthenticationHeaderValue("Basic", auth);
        req.Content = new FormUrlEncodedContent(form);

        try
        {
            using var resp = await _http.SendAsync(req, ct);
            var content = await resp.Content.ReadAsStringAsync(ct);
            if (!resp.IsSuccessStatusCode)
            {
                _log.LogWarning("[SignalWire] SMS send failed {Status}: {Error}", (int)resp.StatusCode, content);
                return null;
            }
            using var doc = JsonDocument.Parse(content);
            return doc.RootElement.TryGetProperty("sid", out var sid) ? sid.GetString() : null;
        }
        catch (Exception ex)
        {
            _log.LogWarning(ex, "[SignalWire] SMS send exception to {To}", to);
            return null;
        }
    }

    /// <summary>Looks up a message's CURRENT delivery status by its SignalWire MessageSid — for backfilling
    /// rows whose status callback never landed (see SmsStatusQueueProcessor / the 2026-07-01 "stuck on
    /// Queued" fix). Returns null on any failure/not-found so the caller can skip that row, not crash a batch.</summary>
    public async Task<string?> GetMessageStatusAsync(string sid, CancellationToken ct = default)
    {
        if (string.IsNullOrWhiteSpace(sid)) return null;
        var projectId = _config["SignalWire:ProjectId"];
        var apiToken  = _config["SignalWire:ApiToken"];
        var spaceUrl  = _config["SignalWire:SpaceUrl"]?.TrimEnd('/');
        if (string.IsNullOrWhiteSpace(projectId) || string.IsNullOrWhiteSpace(apiToken) || string.IsNullOrWhiteSpace(spaceUrl))
            return null;

        var url  = $"https://{spaceUrl}/api/laml/2010-04-01/Accounts/{projectId}/Messages/{sid}.json";
        var auth = Convert.ToBase64String(Encoding.ASCII.GetBytes($"{projectId}:{apiToken}"));

        using var req = new HttpRequestMessage(HttpMethod.Get, url);
        req.Headers.Authorization = new AuthenticationHeaderValue("Basic", auth);
        try
        {
            using var resp = await _http.SendAsync(req, ct);
            var content = await resp.Content.ReadAsStringAsync(ct);
            if (!resp.IsSuccessStatusCode)
            {
                _log.LogWarning("[SignalWire] status lookup {Status} for {Sid}: {Error}", (int)resp.StatusCode, sid, content);
                return null;
            }
            using var doc = JsonDocument.Parse(content);
            return doc.RootElement.TryGetProperty("status", out var status) ? status.GetString() : null;
        }
        catch (Exception ex)
        {
            _log.LogWarning(ex, "[SignalWire] status lookup exception for {Sid}", sid);
            return null;
        }
    }

    /// <summary>Downloads an inbound MMS media part from SignalWire's media store (Basic auth, ProjectId:ApiToken).
    /// Returns (contentType, bytes), or null on failure. The media URLs arrive on the inbound webhook and expire
    /// after the carrier's retention window — callers store the bytes durably at ingest.</summary>
    public async Task<(string ContentType, byte[] Bytes)?> DownloadMediaAsync(string mediaUrl, CancellationToken ct = default)
    {
        var projectId = _config["SignalWire:ProjectId"];
        var apiToken  = _config["SignalWire:ApiToken"];
        if (string.IsNullOrWhiteSpace(mediaUrl) || string.IsNullOrWhiteSpace(projectId) || string.IsNullOrWhiteSpace(apiToken))
            return null;

        using var req = new HttpRequestMessage(HttpMethod.Get, mediaUrl);
        var auth = Convert.ToBase64String(Encoding.ASCII.GetBytes($"{projectId}:{apiToken}"));
        req.Headers.Authorization = new AuthenticationHeaderValue("Basic", auth);
        try
        {
            using var resp = await _http.SendAsync(req, ct);
            if (!resp.IsSuccessStatusCode)
            {
                _log.LogWarning("[SignalWire] media GET {Status} for {Url}", (int)resp.StatusCode, mediaUrl);
                return null;
            }
            var contentType = resp.Content.Headers.ContentType?.MediaType ?? "application/octet-stream";
            var bytes       = await resp.Content.ReadAsByteArrayAsync(ct);
            return (contentType, bytes);
        }
        catch (Exception ex)
        {
            _log.LogWarning(ex, "[SignalWire] media download failed for {Url}", mediaUrl);
            return null;
        }
    }
}
