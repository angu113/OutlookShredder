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
        => await SendSmsAsync(to, body, null, ct);

    /// <summary>Sends an SMS; when <paramref name="statusCallback"/> is set, SignalWire POSTs delivery-status
    /// updates there (our /api/sms/status). Returns the MessageSid, or null on failure.</summary>
    public async Task<string?> SendSmsAsync(string to, string body, string? statusCallback, CancellationToken ct = default)
    {
        var projectId  = _config["SignalWire:ProjectId"]!;
        var apiToken   = _config["SignalWire:ApiToken"]!;
        var fromNumber = _config["SignalWire:FromNumber"]!;
        var spaceUrl   = _config["SignalWire:SpaceUrl"]!.TrimEnd('/');

        var url  = $"https://{spaceUrl}/api/laml/2010-04-01/Accounts/{projectId}/Messages.json";
        var auth = Convert.ToBase64String(Encoding.ASCII.GetBytes($"{projectId}:{apiToken}"));

        var form = new Dictionary<string, string>
        {
            ["From"] = fromNumber,
            ["To"]   = to,
            ["Body"] = body,
        };
        if (!string.IsNullOrWhiteSpace(statusCallback)) form["StatusCallback"] = statusCallback;

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
