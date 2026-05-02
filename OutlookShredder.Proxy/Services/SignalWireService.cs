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
    {
        var projectId  = _config["SignalWire:ProjectId"]!;
        var apiToken   = _config["SignalWire:ApiToken"]!;
        var fromNumber = _config["SignalWire:FromNumber"]!;
        var spaceUrl   = _config["SignalWire:SpaceUrl"]!.TrimEnd('/');

        var url  = $"https://{spaceUrl}/api/laml/2010-04-01/Accounts/{projectId}/Messages.json";
        var auth = Convert.ToBase64String(Encoding.ASCII.GetBytes($"{projectId}:{apiToken}"));

        using var req = new HttpRequestMessage(HttpMethod.Post, url);
        req.Headers.Authorization = new AuthenticationHeaderValue("Basic", auth);
        req.Content = new FormUrlEncodedContent(new Dictionary<string, string>
        {
            ["From"] = fromNumber,
            ["To"]   = to,
            ["Body"] = body,
        });

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
}
