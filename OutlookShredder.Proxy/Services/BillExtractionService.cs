using System.Text;
using System.Text.Json;
using OutlookShredder.Proxy.Models;

namespace OutlookShredder.Proxy.Services;

/// <summary>
/// Second-pass extraction over a SUPPLIER bill/invoice/receipt PDF, run AFTER the cheap text
/// classifier has already routed the email into a financial leaf (Supplier/Invoices and Bills,
/// Supplier/Receipts, Supplier/Statements). Pulls the fields the bill -> PO matcher needs but that
/// live in the PDF, not the email text: the bill total (amount), the supplier's OWN reference, and
/// our PO# if printed.
///
/// Mirrors ErpAiService's provider pattern (Claude Haiku document -> Gemini Flash fallback, via the
/// shared AiRateLimitHandler) but with a bill-specific prompt/schema. Kept distinct from ErpAiService
/// because that service is oriented to OUR OUTBOUND ERP documents, where the roles are reversed (on a
/// supplier bill the "customer" is us, the buyer).
/// </summary>
public sealed class BillExtractionService
{
    private readonly IConfiguration _config;
    private readonly ILogger<BillExtractionService> _log;
    private readonly AiRateLimitTracker _rateLimits;
    private readonly HttpClient _claudeHttp;
    private readonly HttpClient _geminiHttp;

    private const string ClaudeApiUrl = "https://api.anthropic.com/v1/messages";
    private const string GeminiApiUrlTemplate =
        "https://generativelanguage.googleapis.com/v1beta/models/{0}:generateContent?key={1}";

    private const string SystemPromptText = """
        You are a data-extraction assistant for Metal Supermarkets Hackensack, a metals distribution
        company. The attached PDF is a SUPPLIER's bill, invoice, or payment receipt addressed to us
        (we are the buyer / recipient). It may have been delivered via a billing processor such as
        Enmark, QuickBooks/Intuit, Bill.com, or Melio. Extract the following:

        supplier_name:       The supplier/vendor that ISSUED the bill (the seller). Read the REAL
                             supplier even if a processor delivered it. NEVER us (Metal Supermarkets /
                             the buyer / the "Bill To" party).
        amount:              The grand total / amount due, as a plain numeric string with no currency
                             symbol or thousands separators, e.g. "1234.56".
        supplier_reference:  The SUPPLIER's OWN reference printed on the document — their invoice
                             number, sales-order, or order/quote reference (e.g. "3060256",
                             "2808273"). This is the supplier's number, NOT our PO/SO number. If
                             several appear, prefer the invoice number, then the sales-order number.
                             Return just the reference value, not its label.
        our_po_number:       OUR purchase-order number if it is printed on the bill, e.g. HSK-PO0001234
                             (a "Your PO" / "Customer PO" / "PO #" field). Null if not shown.
        currency:            Currency code, e.g. "USD". Default "USD".

        Set is_bill = true if this is a supplier bill/invoice/receipt; false otherwise. Respond ONLY by
        calling the record_bill tool.
        """;

    // Claude tool definition — cache_control marks it for prompt caching alongside the system prompt.
    private static readonly JsonElement _toolJson = JsonDocument.Parse("""
        {
          "name": "record_bill",
          "description": "Record the key fields from a supplier bill/invoice/receipt PDF for matching to our purchase order.",
          "cache_control": { "type": "ephemeral" },
          "input_schema": {
            "type": "object",
            "required": ["is_bill"],
            "properties": {
              "is_bill":            { "type": "boolean",           "description": "True if this is a supplier bill/invoice/receipt." },
              "supplier_name":      { "type": ["string","null"],   "description": "The supplier that issued the bill (the seller), not us the buyer." },
              "amount":             { "type": ["string","null"],   "description": "Grand total / amount due as a plain numeric string, e.g. '1234.56'." },
              "supplier_reference": { "type": ["string","null"],   "description": "The supplier's own invoice/sales-order/quote reference value (not its label). NOT our PO/SO." },
              "our_po_number":      { "type": ["string","null"],   "description": "Our PO number if printed, e.g. HSK-PO0001234. Null if absent." },
              "currency":           { "type": ["string","null"],   "description": "Currency code, default USD." }
            }
          }
        }
        """).RootElement;

    // Gemini JSON-mode schema (nullable:true instead of ["type","null"]).
    private static readonly JsonElement _geminiSchema = JsonDocument.Parse("""
        {
          "type": "object",
          "required": ["is_bill"],
          "properties": {
            "is_bill":            { "type": "boolean" },
            "supplier_name":      { "type": "string", "nullable": true },
            "amount":             { "type": "string", "nullable": true },
            "supplier_reference": { "type": "string", "nullable": true },
            "our_po_number":      { "type": "string", "nullable": true },
            "currency":           { "type": "string", "nullable": true }
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

    public BillExtractionService(
        IConfiguration config,
        ILogger<BillExtractionService> log,
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
    public async Task<BillExtraction?> ExtractBillAsync(string base64Pdf, string fileName, CancellationToken ct = default)
    {
        var result = await TryClaudeAsync(base64Pdf, fileName, ct);
        if (result is not null) return result;

        _log.LogWarning("[Bill] Claude unavailable or failed for {File} — trying Gemini fallback", fileName);
        return await TryGeminiAsync(base64Pdf, fileName, ct);
    }

    // ── Claude ────────────────────────────────────────────────────────────────

    private async Task<BillExtraction?> TryClaudeAsync(string base64Pdf, string fileName, CancellationToken ct)
    {
        var apiKey = _config["Anthropic:ApiKey"];
        if (string.IsNullOrEmpty(apiKey)) { _log.LogDebug("[Bill] Anthropic:ApiKey not configured"); return null; }

        var maxTokens  = int.TryParse(_config["Claude:MaxTokens"],  out var mt) ? mt : 1024;
        var maxRetries = int.TryParse(_config["Claude:MaxRetries"], out var mr) ? mr : 3;
        var model      = _config["BillExtraction:ClaudeModel"] ?? _config["Claude:ErpModel"] ?? "claude-haiku-4-5-20251001";

        var userContent = new object[]
        {
            new {
                type   = "document",
                source = new { type = "base64", media_type = "application/pdf", data = base64Pdf },
                title  = fileName
            },
            new { type = "text", text = $"Extract the supplier bill fields from the attached PDF: {fileName}" }
        };

        var body = new
        {
            model,
            max_tokens  = maxTokens,
            system      = new object[] { new { type = "text", text = SystemPromptText, cache_control = new { type = "ephemeral" } } },
            tools       = new[] { _toolJson },
            tool_choice = new { type = "tool", name = "record_bill" },
            messages    = new[] { new { role = "user", content = userContent } }
        };

        var bodyJson = JsonSerializer.Serialize(body);
        HttpResponseMessage response;
        try { response = await SendClaudeWithRetryAsync(apiKey, bodyJson, maxRetries, ct); }
        catch (Exception ex) { _log.LogError(ex, "[Bill] Claude call failed for {File}", fileName); return null; }

        var raw = await response.Content.ReadAsStringAsync(ct);
        if (!response.IsSuccessStatusCode) { _log.LogError("[Bill] Claude {Status} for {File}: {Body}", response.StatusCode, fileName, raw); return null; }

        using var doc = JsonDocument.Parse(raw);
        if (doc.RootElement.TryGetProperty("stop_reason", out var stop) && stop.GetString() == "max_tokens")
            _log.LogWarning("[Bill] Claude hit max_tokens for {File} — response may be incomplete", fileName);
        if (!doc.RootElement.TryGetProperty("content", out var content)) return null;

        foreach (var block in content.EnumerateArray())
        {
            if (!block.TryGetProperty("type", out var typeEl) || typeEl.GetString() != "tool_use") continue;
            if (!block.TryGetProperty("name", out var nameEl) || nameEl.GetString() != "record_bill") continue;
            if (!block.TryGetProperty("input", out var inputEl)) continue;

            var extraction = JsonSerializer.Deserialize<BillExtraction>(inputEl.GetRawText(), _jsonOpts);
            if (extraction is not null)
                _log.LogInformation("[Bill] Claude: isBill={B} supplier={S} amount={A} ref={R} po={P} ({File})",
                    extraction.IsBill, extraction.SupplierName, extraction.Amount, extraction.SupplierReference, extraction.OurPoNumber, fileName);
            return extraction;
        }
        _log.LogWarning("[Bill] Claude returned no record_bill block for {File}", fileName);
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
                _log.LogWarning(ex, "[Bill] Claude attempt {A}/{T} threw — retrying", attempt + 1, maxRetries + 1);
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

    private async Task<BillExtraction?> TryGeminiAsync(string base64Pdf, string fileName, CancellationToken ct)
    {
        var apiKey = _config["Google:ApiKey"];
        if (string.IsNullOrWhiteSpace(apiKey)) { _log.LogDebug("[Bill] Google:ApiKey not configured — Gemini fallback unavailable"); return null; }

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
                        new { text = $"Extract the supplier bill fields from the attached PDF: {fileName}" }
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
                _log.LogWarning(ex, "[Bill] Gemini attempt {A}/{T} threw — retrying", attempt + 1, maxRetries + 1);
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
        if (!response.IsSuccessStatusCode) { _log.LogError("[Bill] Gemini {Status} for {File}: {Body}", response.StatusCode, fileName, raw); return null; }

        try
        {
            using var doc = JsonDocument.Parse(raw);
            var jsonText = doc.RootElement.GetProperty("candidates")[0].GetProperty("content")
                              .GetProperty("parts")[0].GetProperty("text").GetString();
            if (string.IsNullOrWhiteSpace(jsonText)) return null;
            var extraction = JsonSerializer.Deserialize<BillExtraction>(jsonText, _jsonOpts);
            if (extraction is not null)
                _log.LogInformation("[Bill] Gemini: isBill={B} supplier={S} amount={A} ref={R} po={P} ({File})",
                    extraction.IsBill, extraction.SupplierName, extraction.Amount, extraction.SupplierReference, extraction.OurPoNumber, fileName);
            return extraction;
        }
        catch (Exception ex) { _log.LogError(ex, "[Bill] Failed to parse Gemini response for {File}: {Body}", fileName, raw); return null; }
    }

    private async Task DelayAsync(int attempt, CancellationToken ct)
    {
        var (min, max) = RetryDelays[Math.Min(attempt, RetryDelays.Length - 1)];
        await Task.Delay(Random.Shared.Next(min, max), ct);
    }
}
