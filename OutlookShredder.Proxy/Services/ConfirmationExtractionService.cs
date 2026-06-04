using System.Text;
using System.Text.Json;
using OutlookShredder.Proxy.Models;

namespace OutlookShredder.Proxy.Services;

/// <summary>
/// Second-pass extraction over a SUPPLIER Sales Order Confirmation / acknowledgement PDF, run AFTER
/// the cheap text classifier has already routed the email into Supplier/Order Confirmations. Pulls the
/// field the PO-confirmation matcher needs but that usually lives in the PDF, not the email text: the
/// supplier's promised ship / delivery (ETA) date — which becomes the PO ExpectedDate and schedules
/// the waiting card out of the Prioritize bucket.
///
/// Mirrors BillExtractionService's provider pattern (Claude Haiku document -> Gemini Flash fallback,
/// via the shared AiRateLimitHandler) with a confirmation-specific prompt/schema. Kept distinct from
/// BillExtractionService (bills/receipts) and ErpAiService (our OUTBOUND ERP documents).
/// </summary>
public sealed class ConfirmationExtractionService
{
    private readonly IConfiguration _config;
    private readonly ILogger<ConfirmationExtractionService> _log;
    private readonly AiRateLimitTracker _rateLimits;
    private readonly HttpClient _claudeHttp;
    private readonly HttpClient _geminiHttp;

    private const string ClaudeApiUrl = "https://api.anthropic.com/v1/messages";
    private const string GeminiApiUrlTemplate =
        "https://generativelanguage.googleapis.com/v1beta/models/{0}:generateContent?key={1}";

    private const string SystemPromptText = """
        You are a data-extraction assistant for Metal Supermarkets Hackensack, a metals distribution
        company. The attached PDF is a SUPPLIER's Sales Order Confirmation / Order Acknowledgement for a
        purchase order WE placed (we are the buyer / recipient). Extract the following:

        expected_date:   The supplier's promised SHIP date or DELIVERY / due date for this order — when
                         the material will ship or arrive. Use an explicit "Ship Date", "Delivery
                         Date", "Due Date", "Estimated/Expected Delivery", "Promise Date", or
                         "Scheduled" date printed on the confirmation. If ONLY a lead time is given
                         (e.g. "ships in 2 weeks", "5-7 business days"), leave null — do NOT compute a
                         date. If several dates appear, prefer the delivery/arrival date over the ship
                         date. Return ISO format YYYY-MM-DD (no time).
        our_po_number:   OUR purchase-order number if printed, e.g. HSK-PO0001234 (a "Your PO" /
                         "Customer PO" / "PO #" field). Null if not shown.
        supplier_name:   The supplier/vendor that issued the confirmation (the seller). NEVER us (Metal
                         Supermarkets / the buyer / the "Bill To" / "Ship To" party).

        Set is_confirmation = true if this is a sales order confirmation / acknowledgement; false
        otherwise. Respond ONLY by calling the record_confirmation tool.
        """;

    // Claude tool definition — cache_control marks it for prompt caching alongside the system prompt.
    private static readonly JsonElement _toolJson = JsonDocument.Parse("""
        {
          "name": "record_confirmation",
          "description": "Record the key fields from a supplier sales order confirmation PDF for matching to our purchase order and setting its expected ship/delivery date.",
          "cache_control": { "type": "ephemeral" },
          "input_schema": {
            "type": "object",
            "required": ["is_confirmation"],
            "properties": {
              "is_confirmation": { "type": "boolean",         "description": "True if this is a supplier sales order confirmation / acknowledgement." },
              "expected_date":   { "type": ["string","null"], "description": "Promised ship/delivery/due date as ISO YYYY-MM-DD. Null if only a lead time is stated (do not compute a date)." },
              "our_po_number":   { "type": ["string","null"], "description": "Our PO number if printed, e.g. HSK-PO0001234. Null if absent." },
              "supplier_name":   { "type": ["string","null"], "description": "The supplier that issued the confirmation (the seller), not us the buyer." }
            }
          }
        }
        """).RootElement;

    // Gemini JSON-mode schema (nullable:true instead of ["type","null"]).
    private static readonly JsonElement _geminiSchema = JsonDocument.Parse("""
        {
          "type": "object",
          "required": ["is_confirmation"],
          "properties": {
            "is_confirmation": { "type": "boolean" },
            "expected_date":   { "type": "string", "nullable": true },
            "our_po_number":   { "type": "string", "nullable": true },
            "supplier_name":   { "type": "string", "nullable": true }
          }
        }
        """).RootElement;

    private static readonly (int MinMs, int MaxMs)[] RetryDelays =
    [
        (2_000,  4_000),
        (5_000, 10_000),
        (15_000, 25_000),
    ];

    private static readonly JsonSerializerOptions _jsonOpts = new() { PropertyNameCaseInsensitive = true };

    public ConfirmationExtractionService(
        IConfiguration config,
        ILogger<ConfirmationExtractionService> log,
        AiRateLimitTracker rateLimits,
        ILogger<AiRateLimitHandler> handlerLog)
    {
        _config     = config;
        _log        = log;
        _rateLimits = rateLimits;

        var timeoutSeconds = int.TryParse(_config["Claude:TimeoutSeconds"], out var t) ? t : 60;
        var handler = new AiRateLimitHandler("Claude", rateLimits, handlerLog) { InnerHandler = new HttpClientHandler() };
        _claudeHttp = new HttpClient(handler) { Timeout = TimeSpan.FromSeconds(timeoutSeconds) };
        _geminiHttp = new HttpClient { Timeout = TimeSpan.FromSeconds(timeoutSeconds) };
    }

    /// <summary>Claude (primary) -> Gemini (fallback). Null only if both are unavailable/failed.</summary>
    public async Task<ConfirmationExtraction?> ExtractConfirmationAsync(string base64Pdf, string fileName, CancellationToken ct = default)
    {
        var result = await TryClaudeAsync(base64Pdf, fileName, ct);
        if (result is not null) return result;

        _log.LogWarning("[Confirm] Claude unavailable or failed for {File} — trying Gemini fallback", fileName);
        return await TryGeminiAsync(base64Pdf, fileName, ct);
    }

    // ── Claude ────────────────────────────────────────────────────────────────

    private async Task<ConfirmationExtraction?> TryClaudeAsync(string base64Pdf, string fileName, CancellationToken ct)
    {
        var apiKey = _config["Anthropic:ApiKey"];
        if (string.IsNullOrEmpty(apiKey)) { _log.LogDebug("[Confirm] Anthropic:ApiKey not configured"); return null; }

        var maxTokens  = int.TryParse(_config["Claude:MaxTokens"],  out var mt) ? mt : 1024;
        var maxRetries = int.TryParse(_config["Claude:MaxRetries"], out var mr) ? mr : 3;
        var model      = _config["ConfirmationExtraction:ClaudeModel"] ?? _config["Claude:ErpModel"] ?? "claude-haiku-4-5-20251001";

        var userContent = new object[]
        {
            new {
                type   = "document",
                source = new { type = "base64", media_type = "application/pdf", data = base64Pdf },
                title  = fileName
            },
            new { type = "text", text = $"Extract the order-confirmation fields from the attached PDF: {fileName}" }
        };

        var body = new
        {
            model,
            max_tokens  = maxTokens,
            system      = new object[] { new { type = "text", text = SystemPromptText, cache_control = new { type = "ephemeral" } } },
            tools       = new[] { _toolJson },
            tool_choice = new { type = "tool", name = "record_confirmation" },
            messages    = new[] { new { role = "user", content = userContent } }
        };

        var bodyJson = JsonSerializer.Serialize(body);
        HttpResponseMessage response;
        try { response = await SendClaudeWithRetryAsync(apiKey, bodyJson, maxRetries, ct); }
        catch (Exception ex) { _log.LogError(ex, "[Confirm] Claude call failed for {File}", fileName); return null; }

        var raw = await response.Content.ReadAsStringAsync(ct);
        if (!response.IsSuccessStatusCode) { _log.LogError("[Confirm] Claude {Status} for {File}: {Body}", response.StatusCode, fileName, raw); return null; }

        using var doc = JsonDocument.Parse(raw);
        if (doc.RootElement.TryGetProperty("stop_reason", out var stop) && stop.GetString() == "max_tokens")
            _log.LogWarning("[Confirm] Claude hit max_tokens for {File} — response may be incomplete", fileName);
        if (!doc.RootElement.TryGetProperty("content", out var content)) return null;

        foreach (var block in content.EnumerateArray())
        {
            if (!block.TryGetProperty("type", out var typeEl) || typeEl.GetString() != "tool_use") continue;
            if (!block.TryGetProperty("name", out var nameEl) || nameEl.GetString() != "record_confirmation") continue;
            if (!block.TryGetProperty("input", out var inputEl)) continue;

            var extraction = JsonSerializer.Deserialize<ConfirmationExtraction>(inputEl.GetRawText(), _jsonOpts);
            if (extraction is not null)
                _log.LogInformation("[Confirm] Claude: isConfirmation={B} expectedDate={D} po={P} supplier={S} ({File})",
                    extraction.IsConfirmation, extraction.ExpectedDate, extraction.OurPoNumber, extraction.SupplierName, fileName);
            return extraction;
        }
        _log.LogWarning("[Confirm] Claude returned no record_confirmation block for {File}", fileName);
        return null;
    }

    private async Task<HttpResponseMessage> SendClaudeWithRetryAsync(string apiKey, string bodyJson, int maxRetries, CancellationToken ct)
    {
        HttpResponseMessage? last = null;
        for (int attempt = 0; attempt <= maxRetries; attempt++)
        {
            await _rateLimits.ThrottleIfNeededAsync("Claude", ct);
            var req = new HttpRequestMessage(HttpMethod.Post, ClaudeApiUrl);
            req.Headers.Add("x-api-key", apiKey);
            req.Headers.Add("anthropic-version", "2023-06-01");
            req.Headers.Add("anthropic-beta", "prompt-caching-2024-07-31");
            req.Content = new StringContent(bodyJson, Encoding.UTF8, "application/json");

            try { last = await _claudeHttp.SendAsync(req, ct); }
            catch (Exception ex) when (attempt < maxRetries)
            {
                _log.LogWarning(ex, "[Confirm] Claude attempt {A}/{T} threw — retrying", attempt + 1, maxRetries + 1);
                await DelayAsync(attempt, ct); continue;
            }

            if (last.IsSuccessStatusCode) return last;
            var status = (int)last.StatusCode;
            if ((status == 429 || status >= 500) && attempt < maxRetries) { await DelayAsync(attempt, ct); continue; }
            return last;
        }
        return last!;
    }

    // ── Gemini fallback ─────────────────────────────────────────────────────────

    private async Task<ConfirmationExtraction?> TryGeminiAsync(string base64Pdf, string fileName, CancellationToken ct)
    {
        var apiKey = _config["Google:ApiKey"];
        if (string.IsNullOrWhiteSpace(apiKey)) { _log.LogDebug("[Confirm] Google:ApiKey not configured — Gemini fallback unavailable"); return null; }

        var model = _config["Gemini:Model"] ?? "gemini-2.0-flash";
        var url   = string.Format(GeminiApiUrlTemplate, model, apiKey);

        var body = new
        {
            systemInstruction = new { parts = new[] { new { text = SystemPromptText } } },
            contents = new[]
            {
                new
                {
                    role  = "user",
                    parts = new object[]
                    {
                        new { inlineData = new { mimeType = "application/pdf", data = base64Pdf } },
                        new { text = $"Extract the order-confirmation fields from the attached PDF: {fileName}" }
                    }
                }
            },
            generationConfig = new { responseMimeType = "application/json", responseSchema = _geminiSchema }
        };

        var bodyJson   = JsonSerializer.Serialize(body);
        var maxRetries = int.TryParse(_config["Gemini:MaxRetries"], out var mr) ? mr : 3;

        HttpResponseMessage? response = null;
        for (int attempt = 0; attempt <= maxRetries; attempt++)
        {
            if (attempt > 0) await DelayAsync(attempt - 1, ct);
            try
            {
                var req = new HttpRequestMessage(HttpMethod.Post, url) { Content = new StringContent(bodyJson, Encoding.UTF8, "application/json") };
                response = await _geminiHttp.SendAsync(req, ct);
            }
            catch (Exception ex) when (attempt < maxRetries)
            {
                _log.LogWarning(ex, "[Confirm] Gemini attempt {A}/{T} threw — retrying", attempt + 1, maxRetries + 1);
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
        if (!response.IsSuccessStatusCode) { _log.LogError("[Confirm] Gemini {Status} for {File}: {Body}", response.StatusCode, fileName, raw); return null; }

        try
        {
            using var doc = JsonDocument.Parse(raw);
            var jsonText = doc.RootElement.GetProperty("candidates")[0].GetProperty("content")
                              .GetProperty("parts")[0].GetProperty("text").GetString();
            if (string.IsNullOrWhiteSpace(jsonText)) return null;
            var extraction = JsonSerializer.Deserialize<ConfirmationExtraction>(jsonText, _jsonOpts);
            if (extraction is not null)
                _log.LogInformation("[Confirm] Gemini: isConfirmation={B} expectedDate={D} po={P} supplier={S} ({File})",
                    extraction.IsConfirmation, extraction.ExpectedDate, extraction.OurPoNumber, extraction.SupplierName, fileName);
            return extraction;
        }
        catch (Exception ex) { _log.LogError(ex, "[Confirm] Failed to parse Gemini response for {File}: {Body}", fileName, raw); return null; }
    }

    private async Task DelayAsync(int attempt, CancellationToken ct)
    {
        var (min, max) = RetryDelays[Math.Min(attempt, RetryDelays.Length - 1)];
        await Task.Delay(Random.Shared.Next(min, max), ct);
    }
}
