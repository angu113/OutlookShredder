using System.Text;
using System.Text.Json;
using System.Text.Json.Serialization;
using System.Text.RegularExpressions;
using OutlookShredder.Proxy.Models;

namespace OutlookShredder.Proxy.Services;

/// <summary>
/// Classifies an inbound email/document into the fixed mail taxonomy (see wip/mail-classification.md)
/// and produces keyword tags + extracted refs for the workbench tree and (later) full-text search.
///
/// Mirrors ErpAiService's provider pattern: Claude Haiku (tool_use) primary → Gemini Flash (REST
/// JSON mode) fallback, via the shared AiRateLimitHandler with retry/jitter. Text-only — subject,
/// sender, recipients, truncated body, attachment filenames. Cheap-tier models by default since
/// every inbound message is classified.
/// </summary>
public class MailClassifierService
{
    private readonly IConfiguration _config;
    private readonly ILogger<MailClassifierService> _log;
    private readonly AiRateLimitTracker _rateLimits;
    private readonly MailTaxonomyService _taxonomy;
    private readonly HttpClient _claudeHttp;
    private readonly HttpClient _geminiHttp;

    private const string ClaudeApiUrl = "https://api.anthropic.com/v1/messages";
    private const string GeminiApiUrlTemplate =
        "https://generativelanguage.googleapis.com/v1beta/models/{0}:generateContent?key={1}";

    private const int BodyCharCap = 6000;

    // The taxonomy block is injected at runtime (static base + SP-backed learned hints) so confirmed
    // leaves take effect with no deploy. Prefix/suffix bracket the dynamic block.
    private const string SystemPromptPrefix =
        """
        You are an email classifier for Metal Supermarkets Hackensack, a metals distribution company.

        Classify the email into EXACTLY ONE category from this taxonomy. Pick the single best-fitting
        leaf. If nothing fits, choose "Other" and propose a concise free-text otherLabel naming the
        emergent category.

        TAXONOMY (category path -> meaning):

        """;

    private const string SystemPromptSuffix =
        """

        Strong signals: a purchase-order number like HSK-PO0001234 indicates "Supplier/Order
        Confirmations"; a sales-order reference like HSK-SO0001234 indicates "Customer/Orders".

        Disambiguation rules (apply BEFORE defaulting to a Supplier/* category):
        - Supplier/* categories are ONLY for mail a supplier sends US. If the sender is asking US to
          quote, to supply material, or to send THEM their account statement, it is a Customer/* item:
          Customer/Inquiries for a request to quote/supply, Customer/Statements for a statement request
          — NOT Supplier/RFQ Responses or Supplier/Statements.
        - Mail from another Metal Supermarkets franchise location (an @metalsupermarkets.com store/branch
          address, not corporate/franchisor marketing) placing or following up on an order is treated like
          any customer: Customer/Orders. Franchise/* is reserved for franchisor corporate communications
          (Franchise/Newsletters) and prospective-franchisee inquiries (Franchise/Inquiries).
        - A support ticket or notification from ecommhelp@metalsupermarkets.com, or any mail whose subject
          carries an EC###### reference, relates to a customer web order: Customer/Web Orders.
        - A payment confirmation/receipt for money WE pay a supplier or service provider is Supplier/Receipts
          (even when sent via QuickBooks/Intuit or another billing processor), not Corporate/Receipts.
        - PAID vs UNPAID is decisive for supplier financial mail: a "Purchase Receipt", a credit-card
          authorization/approval, "transaction approved/charged", or "payment received / thank you for
          your payment" means the charge ALREADY happened — that is Supplier/Receipts, NEVER
          Supplier/Invoices and Bills, even if an amount is shown. "Supplier/Invoices and Bills" is ONLY
          an UNPAID request to pay (amount due / please remit / pay-now link).

        For "Supplier/Invoices and Bills" and "Supplier/Receipts" (supplier bills, invoices, and
        payment receipts — including those sent via a billing processor such as Enmark,
        QuickBooks/Intuit, Bill.com or Melio): ALWAYS set supplierName to the REAL supplier (read it
        from the subject, body, or sender domain even when the sender is the processor), and set
        supplierReference to the supplier's OWN reference printed on the document — their invoice
        number, sales-order, or quote/order reference (e.g. "Invoice# 3060256", "Order #2808273") —
        NOT our HSK-PO/HSK-SO. Set poNumber only when OUR HSK-PO number is printed.

        For "Supplier/Order Confirmations" (a supplier acknowledging a PO we placed, usually carrying
        our HSK-PO number): when the email body states an explicit promised SHIP or DELIVERY / due date
        for the order, set expectedDate to it in ISO YYYY-MM-DD. Only a concrete date — if just a lead
        time is given (e.g. "ships in 2 weeks") leave expectedDate null. Set poNumber to our HSK-PO.

        Also produce 5-15 lowercase search keywords (entities, document type, product/metal, supplier
        or customer name, reference numbers) to support full-text search, and extract poNumber,
        soNumber, and amount when present. Respond ONLY by calling the classify_email tool.
        """;

    private async Task<string> BuildSystemPromptAsync(CancellationToken ct) =>
        SystemPromptPrefix + await _taxonomy.RenderForPromptAsync(ct) + SystemPromptSuffix;

    private static readonly (int MinMs, int MaxMs)[] RetryDelays =
    [
        (2_000,  4_000),
        (5_000, 10_000),
        (15_000, 25_000),
    ];

    private static readonly JsonSerializerOptions _jsonOpts = new() { PropertyNameCaseInsensitive = true };

    public MailClassifierService(
        IConfiguration config,
        ILogger<MailClassifierService> log,
        AiRateLimitTracker rateLimits,
        MailTaxonomyService taxonomy,
        ILogger<AiRateLimitHandler> handlerLog)
    {
        _config     = config;
        _log        = log;
        _rateLimits = rateLimits;
        _taxonomy   = taxonomy;

        var timeoutSeconds = int.TryParse(_config["Claude:TimeoutSeconds"], out var t) ? t : 60;
        var handler = new AiRateLimitHandler("Claude", rateLimits, handlerLog) { InnerHandler = new HttpClientHandler() };
        _claudeHttp = new HttpClient(handler) { Timeout = TimeSpan.FromSeconds(timeoutSeconds) };
        _geminiHttp = new HttpClient { Timeout = TimeSpan.FromSeconds(timeoutSeconds) };
    }

    /// <summary>Claude primary → Gemini fallback. Returns null only if both are unavailable/failed.</summary>
    public async Task<MailClassificationResult?> ClassifyAsync(MailClassifyInput input, CancellationToken ct = default)
    {
        var result = await TryClaudeAsync(input, ct);
        if (result is null)
        {
            _log.LogWarning("[MailClassify] Claude unavailable or failed — trying Gemini fallback");
            result = await TryGeminiAsync(input, ct);
        }
        if (result is not null)
        {
            // Deterministic pay-link capture (no AI) — payment-processor bills carry the URL in the body.
            result.PayLink ??= ExtractPayLink(input.BodyText);

            // Deterministic backstop to the sharpened prompt: a clear auth/receipt (the charge already
            // happened) must never sit under "Supplier/Invoices and Bills" (an unpaid request to pay).
            // Re-route it to Supplier/Receipts so the bill->PO matcher never treats it as a payable bill.
            if (string.Equals(result.Category, "Supplier/Invoices and Bills", StringComparison.OrdinalIgnoreCase)
                && LooksLikeAuthOrReceipt(input.Subject, input.FromAddress))
            {
                _log.LogInformation("[MailClassify] re-routed auth/receipt to Supplier/Receipts: \"{Subj}\"",
                    input.Subject.Length > 80 ? input.Subject[..80] : input.Subject);
                result.Category = "Supplier/Receipts";
            }
        }
        return result;
    }

    // Auth/receipt fingerprints — the charge already happened (Supplier/Receipts, not Invoices and Bills).
    private static readonly Regex AuthReceiptSubjectRx = new(
        @"\b(receipt|payment\s+received|payment\s+confirmation|thank\s+you\s+for\s+your\s+payment|authoriz|auth\s*code|approved|charged|paid\s+in\s+full|transaction\s+approved)\b",
        RegexOptions.IgnoreCase | RegexOptions.Compiled);
    private static readonly Regex AuthReceiptSenderRx = new(
        @"(creditcardauth|slimcd\.com)", RegexOptions.IgnoreCase | RegexOptions.Compiled);

    /// <summary>True when subject/sender clearly indicate a payment already happened (a receipt or a
    /// credit-card authorization), so it belongs under Supplier/Receipts rather than Invoices and Bills.</summary>
    internal static bool LooksLikeAuthOrReceipt(string? subject, string? from)
        => (!string.IsNullOrWhiteSpace(subject) && AuthReceiptSubjectRx.IsMatch(subject))
        || (!string.IsNullOrWhiteSpace(from)    && AuthReceiptSenderRx.IsMatch(from));

    private static readonly Regex UrlRx =
        new(@"https?://[^\s<>""')]+", RegexOptions.IgnoreCase | RegexOptions.Compiled);
    // Known payment-processor / pay-page fingerprints (matched against the URL, decoding any
    // link-protect wrapper that carries the real target URL-encoded in its query string).
    private static readonly Regex PayDomainRx =
        new(@"payments?\.enmarksystems\.com|enmarkpay|bill\.com|meliopayments|melio\.me|intuit\.com|quickbooks|paypal\.com/(invoice|pay)|stripe\.com|squareup\.com|/pay(now)?\b",
            RegexOptions.IgnoreCase | RegexOptions.Compiled);
    private static readonly Regex PayCueRx =
        new(@"pay\s*(now|online|invoice|your\s+\w+)", RegexOptions.IgnoreCase | RegexOptions.Compiled);

    /// <summary>Best-effort extraction of a "pay now" URL from the body. Prefers a URL that targets a
    /// known processor (even when wrapped by a link-protect redirector); falls back to the first URL
    /// following a "pay now/online/invoice" cue. Returns null when nothing looks like a pay link.</summary>
    internal static string? ExtractPayLink(string? body)
    {
        if (string.IsNullOrWhiteSpace(body)) return null;
        static string Clean(string u) => u.TrimEnd('.', ',', ')', '>', '"', '\'', ']');

        var urls = UrlRx.Matches(body).Select(m => Clean(m.Value)).ToList();
        if (urls.Count == 0) return null;

        foreach (var u in urls)
        {
            string decoded; try { decoded = Uri.UnescapeDataString(u); } catch { decoded = u; }
            if (PayDomainRx.IsMatch(decoded)) return u;
        }

        var cue = PayCueRx.Match(body);
        if (cue.Success)
        {
            var after = UrlRx.Match(body, cue.Index);
            if (after.Success) return Clean(after.Value);
        }
        return null;
    }

    private string BuildUserText(MailClassifyInput m)
    {
        var body = m.BodyText ?? "";
        if (body.Length > BodyCharCap) body = body[..BodyCharCap] + "\n…[truncated]";
        var atts = m.AttachmentNames.Count > 0 ? string.Join(", ", m.AttachmentNames) : "(none)";
        var threadLine = string.IsNullOrWhiteSpace(m.ThreadCategoryHint) ? "" :
            $"\nConversation context: a previous message in this same email thread was classified as \"{m.ThreadCategoryHint}\". " +
            "Strongly prefer that same category unless this message clearly belongs elsewhere.";
        return $"""
            From: {m.FromName} <{m.FromAddress}>
            To: {m.ToLine}
            Subject: {m.Subject}
            Attachments: {atts}{threadLine}

            Body:
            {body}
            """;
    }

    private async Task<object> BuildToolDefinitionAsync(CancellationToken ct)
    {
        var categories = (await _taxonomy.GetLeavesAsync(ct)).Select(l => l.Path).ToArray();
        return new
        {
            name = "classify_email",
            description = "Record the classification of an inbound email/document into the fixed taxonomy.",
            cache_control = new { type = "ephemeral" },
            input_schema = new
            {
                type = "object",
                required = new[] { "category", "confidence" },
                properties = new Dictionary<string, object>
                {
                    ["category"]   = new { type = "string", @enum = categories, description = "The single best taxonomy path." },
                    ["otherLabel"] = new { type = new[] { "string", "null" }, description = "When category is 'Other', a concise proposed sub-label; otherwise null." },
                    ["supplierName"] = new { type = new[] { "string", "null" }, description = "The actual supplier/vendor this email is from or about. If the sender is a payment processor or billing service (see sender guidance), read the REAL supplier from the body (e.g. 'your bill from Hadco Metal Trading' -> Hadco Metal Trading), not the sender. Null if not a supplier email." },
                    ["confidence"] = new { type = "number", description = "Confidence 0.0-1.0 in the chosen category." },
                    ["keywords"]   = new { type = "array", items = new { type = "string" }, description = "5-15 lowercase search keywords/tags." },
                    ["poNumber"]   = new { type = new[] { "string", "null" }, description = "Our PO number if present, e.g. HSK-PO0001234." },
                    ["soNumber"]   = new { type = new[] { "string", "null" }, description = "Our sales-order ref if present, e.g. HSK-SO0001234." },
                    ["amount"]     = new { type = new[] { "string", "null" }, description = "Monetary total as a plain numeric string if financial, else null." },
                    ["supplierReference"] = new { type = new[] { "string", "null" }, description = "On a supplier bill/invoice/receipt: the SUPPLIER's OWN reference printed on it - their invoice number, sales-order, or quote/order reference (e.g. 'Invoice# 3060256', 'Order #2808273', 'Quote ABC123'). This is the supplier's number, NOT our HSK-PO/HSK-SO. Null if not a supplier financial document or none printed." },
                    ["expectedDate"] = new { type = new[] { "string", "null" }, description = "On a Supplier/Order Confirmation: the promised SHIP or DELIVERY/due date for the order, as ISO YYYY-MM-DD. Only set when an explicit date is stated; null if only a lead time (e.g. 'ships in 2 weeks') or not an order confirmation." },
                    ["reasoning"]  = new { type = "string", description = "One or two sentences justifying the category." },
                }
            }
        };
    }

    // ── Claude ──────────────────────────────────────────────────────────────────

    private async Task<MailClassificationResult?> TryClaudeAsync(MailClassifyInput input, CancellationToken ct)
    {
        var apiKey = _config["Anthropic:ApiKey"];
        if (string.IsNullOrEmpty(apiKey)) { _log.LogDebug("[MailClassify] Anthropic:ApiKey not configured"); return null; }

        var maxTokens  = int.TryParse(_config["MailClassifier:MaxTokens"], out var mt) ? mt : 1024;
        var maxRetries = int.TryParse(_config["Claude:MaxRetries"], out var mr) ? mr : 3;
        var model      = _config["MailClassifier:ClaudeModel"] ?? "claude-haiku-4-5-20251001";

        var systemText = await BuildSystemPromptAsync(ct);
        var body = new
        {
            model,
            max_tokens  = maxTokens,
            system      = new object[] { new { type = "text", text = systemText, cache_control = new { type = "ephemeral" } } },
            tools       = new[] { await BuildToolDefinitionAsync(ct) },
            tool_choice = new { type = "tool", name = "classify_email" },
            messages    = new[] { new { role = "user", content = BuildUserText(input) } }
        };

        var bodyJson = JsonSerializer.Serialize(body);
        HttpResponseMessage response;
        try { response = await SendClaudeWithRetryAsync(apiKey, bodyJson, maxRetries, ct); }
        catch (Exception ex) { _log.LogError(ex, "[MailClassify] Claude call failed"); return null; }

        var raw = await response.Content.ReadAsStringAsync(ct);
        if (!response.IsSuccessStatusCode) { _log.LogError("[MailClassify] Claude {Status}: {Body}", response.StatusCode, raw); return null; }

        using var doc = JsonDocument.Parse(raw);
        if (doc.RootElement.TryGetProperty("usage", out var u))
        {
            int U(string n) => u.TryGetProperty(n, out var v) && v.ValueKind == JsonValueKind.Number ? v.GetInt32() : 0;
            _log.LogInformation("[MailClassify] Claude usage in={In} cacheWrite={CW} cacheRead={CR} out={Out}",
                U("input_tokens"), U("cache_creation_input_tokens"), U("cache_read_input_tokens"), U("output_tokens"));
        }
        if (!doc.RootElement.TryGetProperty("content", out var content)) return null;

        foreach (var block in content.EnumerateArray())
        {
            if (!block.TryGetProperty("type", out var typeEl) || typeEl.GetString() != "tool_use") continue;
            if (!block.TryGetProperty("input", out var inputEl)) continue;
            return MapRaw(inputEl.GetRawText(), "claude", model);
        }
        _log.LogWarning("[MailClassify] Claude returned no classify_email block");
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
                _log.LogWarning(ex, "[MailClassify] Claude attempt {A}/{T} threw — retrying", attempt + 1, maxRetries + 1);
                await DelayAsync(attempt, ct); continue;
            }

            if (last.IsSuccessStatusCode) return last;
            var status = (int)last.StatusCode;
            if ((status == 429 || status >= 500) && attempt < maxRetries) { await DelayAsync(attempt, ct); continue; }
            return last;
        }
        return last!;
    }

    // ── Gemini fallback ───────────────────────────────────────────────────────────

    private async Task<MailClassificationResult?> TryGeminiAsync(MailClassifyInput input, CancellationToken ct)
    {
        var apiKey = _config["Google:ApiKey"];
        if (string.IsNullOrWhiteSpace(apiKey)) { _log.LogDebug("[MailClassify] Google:ApiKey not configured"); return null; }

        var model = _config["MailClassifier:GeminiModel"] ?? "gemini-2.5-flash";
        var url   = string.Format(GeminiApiUrlTemplate, model, apiKey);
        var categories = (await _taxonomy.GetLeavesAsync(ct)).Select(l => l.Path).ToArray();

        var schema = new
        {
            type = "object",
            required = new[] { "category", "confidence" },
            properties = new Dictionary<string, object>
            {
                ["category"]   = new { type = "string", @enum = categories },
                ["otherLabel"] = new { type = "string", nullable = true },
                ["supplierName"] = new { type = "string", nullable = true },
                ["confidence"] = new { type = "number" },
                ["keywords"]   = new { type = "array", items = new { type = "string" }, nullable = true },
                ["poNumber"]   = new { type = "string", nullable = true },
                ["soNumber"]   = new { type = "string", nullable = true },
                ["amount"]     = new { type = "string", nullable = true },
                ["supplierReference"] = new { type = "string", nullable = true },
                ["expectedDate"] = new { type = "string", nullable = true },
                ["reasoning"]  = new { type = "string", nullable = true },
            }
        };

        var body = new
        {
            systemInstruction = new { parts = new[] { new { text = await BuildSystemPromptAsync(ct) } } },
            contents = new[] { new { role = "user", parts = new[] { new { text = BuildUserText(input) } } } },
            generationConfig = new { responseMimeType = "application/json", responseSchema = schema }
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
                _log.LogWarning(ex, "[MailClassify] Gemini attempt {A}/{T} threw — retrying", attempt + 1, maxRetries + 1);
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
        if (!response.IsSuccessStatusCode) { _log.LogError("[MailClassify] Gemini {Status}: {Body}", response.StatusCode, raw); return null; }

        try
        {
            using var doc = JsonDocument.Parse(raw);
            var jsonText = doc.RootElement.GetProperty("candidates")[0].GetProperty("content")
                              .GetProperty("parts")[0].GetProperty("text").GetString();
            return string.IsNullOrWhiteSpace(jsonText) ? null : MapRaw(jsonText, "gemini", model);
        }
        catch (Exception ex) { _log.LogError(ex, "[MailClassify] Failed to parse Gemini response: {Body}", raw); return null; }
    }

    // ── Shared ──────────────────────────────────────────────────────────────────

    private MailClassificationResult? MapRaw(string json, string provider, string model)
    {
        var raw = JsonSerializer.Deserialize<RawClassification>(json, _jsonOpts);
        if (raw is null) return null;

        var category = _taxonomy.Coerce(raw.Category);
        var result = new MailClassificationResult
        {
            Category    = category,
            OtherLabel  = category == "Other" ? raw.OtherLabel?.Trim() : null,
            SupplierName = string.IsNullOrWhiteSpace(raw.SupplierName) ? null : raw.SupplierName.Trim(),
            Confidence  = Math.Clamp(raw.Confidence, 0, 1),
            Keywords    = raw.Keywords?.Where(k => !string.IsNullOrWhiteSpace(k)).Select(k => k.Trim().ToLowerInvariant()).Distinct().ToList() ?? [],
            PoNumber    = string.IsNullOrWhiteSpace(raw.PoNumber) ? null : raw.PoNumber.Trim(),
            SoNumber    = string.IsNullOrWhiteSpace(raw.SoNumber) ? null : raw.SoNumber.Trim(),
            Amount      = string.IsNullOrWhiteSpace(raw.Amount) ? null : raw.Amount.Trim(),
            SupplierReference = string.IsNullOrWhiteSpace(raw.SupplierReference) ? null : raw.SupplierReference.Trim(),
            ExpectedDate = string.IsNullOrWhiteSpace(raw.ExpectedDate) ? null : raw.ExpectedDate.Trim(),
            Reasoning   = raw.Reasoning?.Trim(),
            AiProvider  = provider,
            AiModel     = model,
            RawResponse = json,
        };
        _log.LogInformation("[MailClassify] {Provider} -> {Category} ({Conf:P0}) tags={Tags}",
            provider, result.Category, result.Confidence, result.Keywords.Count);
        return result;
    }

    private async Task DelayAsync(int attempt, CancellationToken ct)
    {
        var (min, max) = RetryDelays[Math.Min(attempt, RetryDelays.Length - 1)];
        await Task.Delay(Random.Shared.Next(min, max), ct);
    }

    private sealed class RawClassification
    {
        public string? Category   { get; set; }
        public string? OtherLabel { get; set; }
        public string? SupplierName { get; set; }
        public double  Confidence { get; set; }
        public List<string>? Keywords { get; set; }
        public string? PoNumber   { get; set; }
        public string? SoNumber   { get; set; }
        public string? Amount     { get; set; }
        public string? SupplierReference { get; set; }
        public string? ExpectedDate { get; set; }
        public string? Reasoning  { get; set; }
    }
}
