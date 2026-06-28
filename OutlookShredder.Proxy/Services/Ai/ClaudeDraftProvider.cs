using System.Text;
using System.Text.Json;
using OutlookShredder.Proxy.Models;

namespace OutlookShredder.Proxy.Services.Ai;

/// <summary>Claude (tool_use) draft provider — primary by registration order. Uses the shared
/// <see cref="AiRateLimitHandler"/> + retry/jitter + prompt caching, mirroring <see cref="MailClassifierService"/>.</summary>
public sealed class ClaudeDraftProvider : IInquiryDraftProvider
{
    private readonly IConfiguration              _config;
    private readonly ILogger<ClaudeDraftProvider> _log;
    private readonly AiRateLimitTracker          _rateLimits;
    private readonly HttpClient                  _http;

    private const string ApiUrl = "https://api.anthropic.com/v1/messages";

    public string Name => "Claude";
    public bool IsConfigured => !string.IsNullOrEmpty(_config["Anthropic:ApiKey"]);

    public ClaudeDraftProvider(IConfiguration config, ILogger<ClaudeDraftProvider> log,
        AiRateLimitTracker rateLimits, ILogger<AiRateLimitHandler> handlerLog)
    {
        _config     = config;
        _log        = log;
        _rateLimits = rateLimits;

        var timeoutSeconds = int.TryParse(_config["Claude:TimeoutSeconds"], out var t) ? t : 60;
        var handler = new AiRateLimitHandler("Claude", rateLimits, handlerLog) { InnerHandler = new HttpClientHandler() };
        _http = new HttpClient(handler) { Timeout = TimeSpan.FromSeconds(timeoutSeconds) };
    }

    public async Task<InquiryDraftResult?> DraftAsync(InquiryDraftInput input, CancellationToken ct = default)
    {
        var apiKey = _config["Anthropic:ApiKey"];
        if (string.IsNullOrEmpty(apiKey)) return null;

        var maxTokens  = int.TryParse(_config["InquiryDraft:MaxTokens"], out var mt) ? mt : 512;
        var maxRetries = int.TryParse(_config["Claude:MaxRetries"], out var mr) ? mr : 3;
        var model      = _config["InquiryDraft:ClaudeModel"] ?? "claude-haiku-4-5-20251001";

        var body = new
        {
            model,
            max_tokens  = maxTokens,
            system      = new object[] { new { type = "text", text = InquiryDraftPrompt.SystemPrompt(_config), cache_control = new { type = "ephemeral" } } },
            tools       = new[] { ToolDefinition() },
            tool_choice = new { type = "tool", name = "draft_reply" },
            messages    = new[] { new { role = "user", content = BuildContent(input) } },
        };

        HttpResponseMessage response;
        try { response = await SendWithRetryAsync(apiKey, JsonSerializer.Serialize(body), maxRetries, ct); }
        catch (Exception ex) { _log.LogError(ex, "[InquiryDraft] Claude call failed"); return null; }

        var raw = await response.Content.ReadAsStringAsync(ct);
        if (!response.IsSuccessStatusCode) { _log.LogError("[InquiryDraft] Claude {Status}: {Body}", response.StatusCode, raw); return null; }

        using var doc = JsonDocument.Parse(raw);
        if (!doc.RootElement.TryGetProperty("content", out var content)) return null;
        foreach (var block in content.EnumerateArray())
        {
            if (!block.TryGetProperty("type", out var typeEl) || typeEl.GetString() != "tool_use") continue;
            if (!block.TryGetProperty("input", out var inputEl)) continue;
            return InquiryDraftPrompt.MapResult(inputEl.GetRawText(), model);
        }
        _log.LogWarning("[InquiryDraft] Claude returned no draft_reply block");
        return null;
    }

    /// <summary>User-turn content: plain text when there are no attachments (cheapest), else a content-block
    /// array with the text plus an image/document block per attachment (mirrors ClaudeExtractionService).</summary>
    private static object BuildContent(InquiryDraftInput input)
    {
        var text = InquiryDraftPrompt.BuildUserText(input);
        if (input.Attachments is not { Count: > 0 }) return text;

        var blocks = new List<object> { new { type = "text", text } };
        foreach (var a in input.Attachments)
        {
            if (a.MimeType.StartsWith("image/", StringComparison.OrdinalIgnoreCase))
                blocks.Add(new { type = "image", source = new { type = "base64", media_type = a.MimeType, data = a.Base64 } });
            else if (a.MimeType.Equals("application/pdf", StringComparison.OrdinalIgnoreCase))
                blocks.Add(new { type = "document", source = new { type = "base64", media_type = "application/pdf", data = a.Base64 }, title = a.FileName ?? "attachment" });
        }
        return blocks.ToArray();
    }

    private static object ToolDefinition() => new
    {
        name = "draft_reply",
        description = "Record a suggested SMS reply to the customer plus a classification of their message.",
        cache_control = new { type = "ephemeral" },
        input_schema = new
        {
            type = "object",
            required = new[] { "reply", "intent", "urgency", "needsQuote" },
            properties = new Dictionary<string, object>
            {
                ["reply"]      = new { type = "string", description = "The suggested reply text, ready to send as an SMS (1-2 short sentences)." },
                ["intent"]     = new { type = "string", @enum = InquiryDraftPrompt.Intents, description = "The customer's primary intent." },
                ["urgency"]    = new { type = "string", @enum = InquiryDraftPrompt.Urgencies, description = "How time-sensitive the message is." },
                ["needsQuote"] = new { type = "boolean", description = "True if answering requires raising a price quotation." },
                ["options"]    = new { type = "array", items = new { type = "string" }, description = "Small discrete choices for the customer (e.g. [\"Steel\",\"Stainless\",\"Aluminum\"]) when asking them to pick one. Empty otherwise." },
            },
        },
    };

    private async Task<HttpResponseMessage> SendWithRetryAsync(string apiKey, string bodyJson, int maxRetries, CancellationToken ct)
    {
        HttpResponseMessage? last = null;
        for (int attempt = 0; attempt <= maxRetries; attempt++)
        {
            await _rateLimits.ThrottleIfNeededAsync("Claude", ct);
            var req = new HttpRequestMessage(HttpMethod.Post, ApiUrl);
            req.Headers.Add("x-api-key", apiKey);
            req.Headers.Add("anthropic-version", "2023-06-01");
            req.Headers.Add("anthropic-beta", "prompt-caching-2024-07-31");
            req.Content = new StringContent(bodyJson, Encoding.UTF8, "application/json");

            try { last = await _http.SendAsync(req, ct); }
            catch (Exception ex) when (attempt < maxRetries)
            {
                _log.LogWarning(ex, "[InquiryDraft] Claude attempt {A}/{T} threw — retrying", attempt + 1, maxRetries + 1);
                await InquiryDraftPrompt.DelayAsync(attempt, ct); continue;
            }

            if (last.IsSuccessStatusCode) return last;
            var status = (int)last.StatusCode;
            if ((status == 429 || status >= 500) && attempt < maxRetries) { await InquiryDraftPrompt.DelayAsync(attempt, ct); continue; }
            return last;
        }
        return last!;
    }
}
