using System.Text;
using System.Text.Json;
using OutlookShredder.Proxy.Models;

namespace OutlookShredder.Proxy.Services;

/// <summary>
/// Calls Claude (primary) or Gemini (fallback) to classify a PDF as an ERP document
/// and extract its structured data.
/// Separate from IAiExtractionService — ERP documents are outbound company records,
/// not inbound supplier quotes, and require a different extraction schema.
/// </summary>
public class ErpAiService
{
    private readonly IConfiguration _config;
    private readonly ILogger<ErpAiService> _log;
    private readonly AiRateLimitTracker _rateLimits;
    private readonly HttpClient _claudeHttp;
    private readonly HttpClient _geminiHttp;

    private const string ClaudeApiUrl = "https://api.anthropic.com/v1/messages";
    private const string GeminiApiUrlTemplate =
        "https://generativelanguage.googleapis.com/v1beta/models/{0}:generateContent?key={1}";

    private const string SystemPromptText = """
        You are a data extraction assistant for Metal Supermarkets Hackensack, a metals
        distribution company.

        The PDF you receive is a confirmed ERP-generated document (invoice, sales order, quotation,
        picking slip, shipping note, purchase order, or payment receipt). The document type and our
        internal reference number have already been determined from the filename — you do NOT need to
        classify the document or extract our own reference number.

        Your job is to extract the following fields:

        CUSTOMER INFORMATION
        customer_name:      The company name the document is addressed to (the buyer or recipient).
                            On picking slips: look for the "Ship To:" label, then take the company
                            name from the line BELOW it (not the text to its right on the same line).
                            Do NOT use the Customer Rep name or store name.
                            For purchase orders this will be our supplier.
        customer_reference: Any PO or reference number assigned BY the customer, e.g. fields labelled
                            "Customer PO", "Your Ref", "Cust. Order", "Customer Reference", "PO #".
                            Leave null if not present. NEVER put our own ERP reference here.

        DATE
        document_date: The date printed on the document (invoice date, order date, quotation date).

        TOTALS
        total_amount: Grand total as a plain numeric string with no currency symbol, e.g. "1234.56".
                      For picking slips and shipping notes that carry no pricing, leave null.
        currency:     Currency code, e.g. "USD". Default "USD" if not stated.

        LINE ITEMS
        For each product line: description (include alloy, form, and dimensions), product code if
        shown, quantity, unit (PC, FT, LB, etc.), unit_price, total_price.
        Prices as plain numeric strings. Omit non-product lines (freight, tax, discounts, etc.).

        NOTES
        Any payment terms, delivery instructions, or other notable text worth recording.

        Always set is_erp_document = true.
        Leave document_type and document_number as null — they are set from the filename automatically.
        """;

    // Claude tool definition — cache_control marks this for prompt caching alongside the system prompt.
    private static readonly JsonElement _toolJson = JsonDocument.Parse("""
        {
          "name": "record_erp_document",
          "description": "Record detail data extracted from a confirmed Metal Supermarkets ERP PDF. Document type and reference number are already known from the filename — focus on customer, date, totals, and line items.",
          "cache_control": { "type": "ephemeral" },
          "input_schema": {
            "type": "object",
            "required": ["is_erp_document"],
            "properties": {
              "is_erp_document":    { "type": "boolean",           "description": "Always true — the file has been confirmed as ERP from its filename." },
              "document_type":      { "type": ["string","null"],   "description": "Leave null — set from filename automatically.", "enum": ["Quotation","SalesOrder","PickingSlip","PurchaseOrder","Invoice","CreditNote","Payment","ShippingNote","Unknown",null] },
              "document_number":    { "type": ["string","null"],   "description": "Our ERP reference (HSK-... or 020803-...). Leave null for most documents — set from filename. Only populate if the filename did not contain a record number (e.g. a bare PurchaseOrder.pdf)." },
              "customer_name":      { "type": ["string","null"],   "description": "Company the document is addressed to. On picking slips: the company name on the line BELOW 'Ship To:' (not the text to the right of it on the same line). Not the Customer Rep or store name." },
              "customer_reference": { "type": ["string","null"],   "description": "Customer's own PO or reference number. Never our own ERP reference." },
              "document_date":      { "type": ["string","null"],   "description": "Date as printed on the document." },
              "total_amount":       { "type": ["string","null"],   "description": "Grand total as a plain numeric string, e.g. '1234.56'. Null for documents with no pricing (picking slips, shipping notes)." },
              "currency":           { "type": ["string","null"],   "description": "Currency code, e.g. USD. Default USD." },
              "line_items": {
                "type": "array",
                "items": {
                  "type": "object",
                  "properties": {
                    "description": { "type": ["string","null"] },
                    "code":        { "type": ["string","null"] },
                    "quantity":    { "type": ["string","null"] },
                    "unit":        { "type": ["string","null"] },
                    "unit_price":  { "type": ["string","null"] },
                    "total_price": { "type": ["string","null"] }
                  }
                }
              },
              "notes": { "type": ["string","null"], "description": "Payment terms, delivery instructions, or other notable text." }
            }
          }
        }
        """).RootElement;

    // Gemini JSON-mode response schema (nullable: true instead of ["type","null"])
    private static readonly JsonElement _geminiSchema = JsonDocument.Parse("""
        {
          "type": "object",
          "required": ["is_erp_document"],
          "properties": {
            "is_erp_document":    { "type": "boolean" },
            "document_type":      { "type": "string",  "nullable": true },
            "document_number":    { "type": "string",  "nullable": true },
            "customer_name":      { "type": "string",  "nullable": true },
            "customer_reference": { "type": "string",  "nullable": true },
            "document_date":      { "type": "string",  "nullable": true },
            "total_amount":       { "type": "string",  "nullable": true },
            "currency":           { "type": "string",  "nullable": true },
            "line_items": {
              "type": "array",
              "nullable": true,
              "items": {
                "type": "object",
                "properties": {
                  "description": { "type": "string", "nullable": true },
                  "code":        { "type": "string", "nullable": true },
                  "quantity":    { "type": "string", "nullable": true },
                  "unit":        { "type": "string", "nullable": true },
                  "unit_price":  { "type": "string", "nullable": true },
                  "total_price": { "type": "string", "nullable": true }
                }
              }
            },
            "notes": { "type": "string", "nullable": true }
          }
        }
        """).RootElement;

    private static readonly (int MinMs, int MaxMs)[] RetryDelays =
    [
        (2_000,  4_000),
        (5_000, 10_000),
        (15_000, 25_000),
    ];

    private static readonly JsonSerializerOptions _jsonOpts =
        new() { PropertyNameCaseInsensitive = true };

    public ErpAiService(
        IConfiguration config,
        ILogger<ErpAiService> log,
        AiRateLimitTracker rateLimits,
        ILogger<AiRateLimitHandler> handlerLog)
    {
        _config     = config;
        _log        = log;
        _rateLimits = rateLimits;

        var timeoutSeconds = int.TryParse(_config["Claude:TimeoutSeconds"], out var t) ? t : 60;

        var handler = new AiRateLimitHandler("Claude", rateLimits, handlerLog)
        {
            InnerHandler = new HttpClientHandler()
        };
        _claudeHttp = new HttpClient(handler) { Timeout = TimeSpan.FromSeconds(timeoutSeconds) };
        _geminiHttp = new HttpClient { Timeout = TimeSpan.FromSeconds(timeoutSeconds) };
    }

    /// <summary>
    /// Classifies the PDF and extracts its ERP data.
    /// Tries Claude first; falls back to Gemini if Claude is unavailable or fails.
    /// Returns null only when both providers are unavailable or both fail.
    /// Returns an extraction with IsErpDocument=false for non-ERP PDFs.
    /// </summary>
    public async Task<ErpExtraction?> ExtractAsync(string base64Pdf, string fileName, CancellationToken ct = default)
    {
        var result = await TryClaudeAsync(base64Pdf, fileName, ct);
        if (result is not null) return result;

        _log.LogWarning("[ERP] Claude unavailable or failed for {File} — trying Gemini fallback", fileName);
        return await TryGeminiAsync(base64Pdf, fileName, ct);
    }

    // ── Claude ────────────────────────────────────────────────────────────────

    private async Task<ErpExtraction?> TryClaudeAsync(string base64Pdf, string fileName, CancellationToken ct)
    {
        var apiKey = _config["Anthropic:ApiKey"];
        if (string.IsNullOrEmpty(apiKey))
        {
            _log.LogDebug("[ERP] Anthropic:ApiKey not configured");
            return null;
        }

        var maxTokens  = int.TryParse(_config["Claude:MaxTokens"],  out var mt) ? mt : 2048;
        var maxRetries = int.TryParse(_config["Claude:MaxRetries"], out var mr) ? mr : 3;
        var model      = _config["Claude:ErpModel"] ?? "claude-haiku-4-5-20251001";

        var userContent = new object[]
        {
            new {
                type   = "document",
                source = new { type = "base64", media_type = "application/pdf", data = base64Pdf },
                title  = fileName
            },
            new { type = "text", text = $"Classify and extract data from the attached PDF: {fileName}" }
        };

        var body = new
        {
            model,
            max_tokens  = maxTokens,
            system      = new object[]
            {
                new { type = "text", text = SystemPromptText, cache_control = new { type = "ephemeral" } }
            },
            tools       = new[] { _toolJson },
            tool_choice = new { type = "tool", name = "record_erp_document" },
            messages    = new[] { new { role = "user", content = userContent } }
        };

        var bodyJson = JsonSerializer.Serialize(body);
        HttpResponseMessage response;
        try
        {
            response = await SendClaudeWithRetryAsync(apiKey, bodyJson, maxRetries, ct);
        }
        catch (Exception ex)
        {
            _log.LogError(ex, "[ERP] Claude call failed for {File}", fileName);
            return null;
        }

        var raw = await response.Content.ReadAsStringAsync(ct);

        if (!response.IsSuccessStatusCode)
        {
            _log.LogError("[ERP] Claude API error {Status} for {File}: {Body}", response.StatusCode, fileName, raw);
            return null;
        }

        using var doc = JsonDocument.Parse(raw);
        var root      = doc.RootElement;

        if (root.TryGetProperty("stop_reason", out var stop) && stop.GetString() == "max_tokens")
            _log.LogWarning("[ERP] Claude hit max_tokens for {File} — response may be incomplete", fileName);

        if (!root.TryGetProperty("content", out var content)) return null;

        foreach (var block in content.EnumerateArray())
        {
            if (!block.TryGetProperty("type", out var typeEl) || typeEl.GetString() != "tool_use") continue;
            if (!block.TryGetProperty("name", out var nameEl) || nameEl.GetString() != "record_erp_document") continue;
            if (!block.TryGetProperty("input", out var inputEl)) continue;

            var extraction = JsonSerializer.Deserialize<ErpExtraction>(inputEl.GetRawText(), _jsonOpts);
            if (extraction is not null)
                _log.LogInformation("[ERP] Claude: IsErp={IsErp} Type={Type} Number={Number} File={File}",
                    extraction.IsErpDocument, extraction.DocumentType, extraction.DocumentNumber, fileName);
            return extraction;
        }

        _log.LogWarning("[ERP] Claude returned no record_erp_document block for {File}", fileName);
        return null;
    }

    private async Task<HttpResponseMessage> SendClaudeWithRetryAsync(
        string apiKey, string bodyJson, int maxRetries, CancellationToken ct)
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

            try
            {
                last = await _claudeHttp.SendAsync(req, ct);
            }
            catch (Exception ex) when (attempt < maxRetries)
            {
                _log.LogWarning(ex, "[ERP] Claude attempt {A}/{T} threw — retrying", attempt + 1, maxRetries + 1);
                await DelayAsync(attempt, ct);
                continue;
            }

            if (last.IsSuccessStatusCode) return last;

            var status = (int)last.StatusCode;
            if ((status == 429 || status >= 500) && attempt < maxRetries)
            {
                _log.LogWarning("[ERP] Claude {Status} attempt {A}/{T} — retrying", status, attempt + 1, maxRetries + 1);
                await DelayAsync(attempt, ct);
                continue;
            }

            return last;
        }

        return last!;
    }

    // ── Gemini fallback ───────────────────────────────────────────────────────

    private async Task<ErpExtraction?> TryGeminiAsync(string base64Pdf, string fileName, CancellationToken ct)
    {
        var apiKey = _config["Google:ApiKey"];
        if (string.IsNullOrWhiteSpace(apiKey))
        {
            _log.LogDebug("[ERP] Google:ApiKey not configured — Gemini fallback unavailable");
            return null;
        }

        var model = _config["Gemini:Model"] ?? "gemini-2.0-flash";
        var url   = string.Format(GeminiApiUrlTemplate, model, apiKey);

        var body = new
        {
            systemInstruction = new
            {
                parts = new[] { new { text = SystemPromptText } }
            },
            contents = new[]
            {
                new
                {
                    role  = "user",
                    parts = new object[]
                    {
                        new { inlineData = new { mimeType = "application/pdf", data = base64Pdf } },
                        new { text = $"Classify and extract data from the attached PDF: {fileName}" }
                    }
                }
            },
            generationConfig = new
            {
                responseMimeType = "application/json",
                responseSchema   = _geminiSchema
            }
        };

        var bodyJson = JsonSerializer.Serialize(body);
        var maxRetries = int.TryParse(_config["Gemini:MaxRetries"], out var mr) ? mr : 3;

        HttpResponseMessage? response = null;
        for (int attempt = 0; attempt <= maxRetries; attempt++)
        {
            if (attempt > 0) await DelayAsync(attempt - 1, ct);
            try
            {
                var req = new HttpRequestMessage(HttpMethod.Post, url)
                {
                    Content = new StringContent(bodyJson, Encoding.UTF8, "application/json")
                };
                response = await _geminiHttp.SendAsync(req, ct);
            }
            catch (Exception ex) when (attempt < maxRetries)
            {
                _log.LogWarning(ex, "[ERP] Gemini attempt {A}/{T} threw — retrying", attempt + 1, maxRetries + 1);
                continue;
            }

            if (response is null) continue;

            var status = (int)response.StatusCode;
            if (response.IsSuccessStatusCode) break;
            if ((status == 429 || status >= 500) && attempt < maxRetries)
            {
                _log.LogWarning("[ERP] Gemini {Status} attempt {A}/{T} — retrying", status, attempt + 1, maxRetries + 1);
                response = null;
                continue;
            }
            break;
        }

        if (response is null) return null;

        var raw = await response.Content.ReadAsStringAsync(ct);
        if (!response.IsSuccessStatusCode)
        {
            _log.LogError("[ERP] Gemini API error {Status} for {File}: {Body}", response.StatusCode, fileName, raw);
            return null;
        }

        try
        {
            using var doc = JsonDocument.Parse(raw);
            var root = doc.RootElement;

            // Gemini wraps the JSON text in candidates[0].content.parts[0].text
            var jsonText = root
                .GetProperty("candidates")[0]
                .GetProperty("content")
                .GetProperty("parts")[0]
                .GetProperty("text")
                .GetString();

            if (string.IsNullOrWhiteSpace(jsonText)) return null;

            var extraction = JsonSerializer.Deserialize<ErpExtraction>(jsonText, _jsonOpts);
            if (extraction is not null)
                _log.LogInformation("[ERP] Gemini: IsErp={IsErp} Type={Type} Number={Number} File={File}",
                    extraction.IsErpDocument, extraction.DocumentType, extraction.DocumentNumber, fileName);
            return extraction;
        }
        catch (Exception ex)
        {
            _log.LogError(ex, "[ERP] Failed to parse Gemini response for {File}: {Body}", fileName, raw);
            return null;
        }
    }

    // ── Shared ────────────────────────────────────────────────────────────────

    private async Task DelayAsync(int attempt, CancellationToken ct)
    {
        var (min, max) = RetryDelays[Math.Min(attempt, RetryDelays.Length - 1)];
        await Task.Delay(Random.Shared.Next(min, max), ct);
    }
}
