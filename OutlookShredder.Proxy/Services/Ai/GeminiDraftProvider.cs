using System.Text;
using System.Text.Json;
using OutlookShredder.Proxy.Models;

namespace OutlookShredder.Proxy.Services.Ai;

/// <summary>Gemini (JSON mode / responseSchema) draft provider — fallback by registration order. Returns the
/// same {reply, intent, urgency, needsQuote} shape as Claude, so <see cref="InquiryDraftPrompt.MapResult"/>
/// parses either provider.</summary>
public sealed class GeminiDraftProvider : IInquiryDraftProvider
{
    private readonly IConfiguration               _config;
    private readonly ILogger<GeminiDraftProvider> _log;
    private readonly HttpClient                   _http;

    private const string ApiUrlTemplate =
        "https://generativelanguage.googleapis.com/v1beta/models/{0}:generateContent?key={1}";

    public string Name => "Gemini";
    public bool IsConfigured => !string.IsNullOrWhiteSpace(_config["Google:ApiKey"]);

    public GeminiDraftProvider(IConfiguration config, ILogger<GeminiDraftProvider> log)
    {
        _config = config;
        _log    = log;
        var timeoutSeconds = int.TryParse(_config["Claude:TimeoutSeconds"], out var t) ? t : 60;
        _http = new HttpClient { Timeout = TimeSpan.FromSeconds(timeoutSeconds) };
    }

    public async Task<InquiryDraftResult?> DraftAsync(InquiryDraftInput input, CancellationToken ct = default)
    {
        var apiKey = _config["Google:ApiKey"];
        if (string.IsNullOrWhiteSpace(apiKey)) return null;

        var model = _config["InquiryDraft:GeminiModel"] ?? "gemini-2.5-flash";
        var url   = string.Format(ApiUrlTemplate, model, apiKey);

        var schema = new
        {
            type = "object",
            required = new[] { "reply", "intent", "urgency", "needsQuote" },
            properties = new Dictionary<string, object>
            {
                ["reply"]      = new { type = "string" },
                ["intent"]     = new { type = "string", @enum = InquiryDraftPrompt.Intents },
                ["urgency"]    = new { type = "string", @enum = InquiryDraftPrompt.Urgencies },
                ["needsQuote"] = new { type = "boolean" },
            }
        };
        var body = new
        {
            systemInstruction = new { parts = new[] { new { text = InquiryDraftPrompt.SystemPrompt(_config) } } },
            contents = new[] { new { role = "user", parts = new[] { new { text = InquiryDraftPrompt.BuildUserText(input) } } } },
            generationConfig = new { responseMimeType = "application/json", responseSchema = schema }
        };

        var bodyJson   = JsonSerializer.Serialize(body);
        var maxRetries = int.TryParse(_config["Gemini:MaxRetries"], out var mr) ? mr : 3;

        HttpResponseMessage? response = null;
        for (int attempt = 0; attempt <= maxRetries; attempt++)
        {
            if (attempt > 0) await InquiryDraftPrompt.DelayAsync(attempt - 1, ct);
            try
            {
                var req = new HttpRequestMessage(HttpMethod.Post, url) { Content = new StringContent(bodyJson, Encoding.UTF8, "application/json") };
                response = await _http.SendAsync(req, ct);
            }
            catch (Exception ex) when (attempt < maxRetries)
            {
                _log.LogWarning(ex, "[InquiryDraft] Gemini attempt {A}/{T} threw — retrying", attempt + 1, maxRetries + 1);
                continue;
            }
            if (response is null) continue;
            var status = (int)response.StatusCode;
            if (response.IsSuccessStatusCode) break;
            if ((status == 429 || status >= 500) && attempt < maxRetries) { response = null; continue; }
            break;
        }

        if (response is null) return null;
        var raw = await response.Content.ReadAsStringAsync(ct);
        if (!response.IsSuccessStatusCode) { _log.LogError("[InquiryDraft] Gemini {Status}: {Body}", response.StatusCode, raw); return null; }

        try
        {
            using var doc = JsonDocument.Parse(raw);
            var jsonText = doc.RootElement.GetProperty("candidates")[0].GetProperty("content")
                              .GetProperty("parts")[0].GetProperty("text").GetString();
            return string.IsNullOrWhiteSpace(jsonText) ? null : InquiryDraftPrompt.MapResult(jsonText, model);
        }
        catch (Exception ex) { _log.LogError(ex, "[InquiryDraft] Failed to parse Gemini response: {Body}", raw); return null; }
    }
}
