using System.Text;
using System.Text.Json;
using OutlookShredder.Proxy.Models;

namespace OutlookShredder.Proxy.Services;

/// <summary>
/// Calls the Anthropic API to extract structured RFQ data.
/// The API key never leaves the server — it is read from configuration
/// (User Secrets in development, environment variable / Key Vault in production).
/// </summary>
public class ClaudeService
{
    private readonly IConfiguration _config;
    private readonly ILogger<ClaudeService> _log;
    private readonly HttpClient _http;

    private const string Model = "claude-sonnet-4-6";
    private const string ApiUrl = "https://api.anthropic.com/v1/messages";

    private const string ExtractionPrompt = """
        You are extracting supplier quote data from an RFQ email or attachment.
        Return ONLY a valid JSON object — no commentary, no markdown fences. Use null for absent fields.

        {
          "jobReference":          "6-character alphanumeric code from [XXXXXX] pattern, no brackets",
          "supplierName":          "company or person providing the quote",
          "dateOfQuote":           "YYYY-MM-DD or null",
          "estimatedDeliveryDate": "YYYY-MM-DD or null",
          "products": [
            {
              "productName":              "product name as stated by supplier",
              "unitsRequested":           number or null,
              "unitsQuoted":              number or null,
              "lengthPerUnit":            number or null,
              "lengthUnit":               "ft | m | in | mm | cm | null",
              "weightPerUnit":            number or null,
              "weightUnit":               "lb | kg | oz | g | null",
              "pricePerPound":            number or null,
              "pricePerFoot":             number or null,
              "supplierProductComments":  "supplier notes for this product or null"
            }
          ]
        }

        Rules:
        - Every distinct product or material is a separate entry in products[].
        - Prices are numbers only — no $ or currency symbols.
        - If only one product, still return it inside the products array.
        - supplierProductComments: capture availability, lead times, alternates, spec notes.
        """;

    public ClaudeService(IConfiguration config, ILogger<ClaudeService> log)
    {
        _config = config;
        _log    = log;
        _http   = new HttpClient();
    }

    public async Task<RfqExtraction?> ExtractAsync(ExtractRequest req)
    {
        var apiKey = _config["Anthropic:ApiKey"]
            ?? throw new InvalidOperationException(
                "Anthropic:ApiKey is not configured. " +
                "Add it to User Secrets (right-click project > Manage User Secrets) or appsettings.Development.json.");

        var jobHint = req.JobRefs.Any()
            ? $"Job reference(s) pre-identified by pattern scan: {string.Join(", ", req.JobRefs)}. " +
              "Use these to confirm the jobReference field."
            : string.Empty;

        var systemPrompt = $"You are a precise supplier quote data extraction assistant. " +
                           $"Return ONLY valid JSON. Use null for absent fields. {jobHint}";

        // Determine attachment type and resolve the text content to send.
        // - PDF/DOCX → sent as a native Claude document; no text content decoding needed.
        // - Text-type attachment (txt, csv, rtf, html…) → decode base64 to UTF-8 text.
        // - Body → req.Content is already plain text.
        var ct     = (req.ContentType ?? string.Empty).ToLowerInvariant();
        var isPdf  = ct.Contains("pdf");
        var isDocx = ct.Contains("wordprocessingml") || ct.Contains("msword");

        // Decoded text for non-PDF/DOCX attachments (null means use req.Content or doc API)
        string? decodedAttachmentText = null;
        if (!isPdf && !isDocx && req.SourceType == "attachment" && !string.IsNullOrEmpty(req.Base64Data))
        {
            try
            {
                decodedAttachmentText = Encoding.UTF8.GetString(Convert.FromBase64String(req.Base64Data));
                _log.LogDebug("Decoded text attachment '{File}' ({Chars} chars)", req.FileName, decodedAttachmentText.Length);
            }
            catch (Exception ex)
            {
                _log.LogWarning("Cannot decode attachment '{File}' as UTF-8 text: {Err}", req.FileName, ex.Message);
            }
        }

        // Build content array — PDF/DOCX use Claude's native document understanding
        object userContent;
        if ((isPdf || isDocx) && !string.IsNullOrEmpty(req.Base64Data))
        {
            var mediaType = isPdf
                ? "application/pdf"
                : "application/vnd.openxmlformats-officedocument.wordprocessingml.document";

            userContent = new object[]
            {
                new {
                    type   = "document",
                    source = new { type = "base64", media_type = mediaType, data = req.Base64Data },
                    title  = req.FileName ?? "attachment"
                },
                new {
                    type = "text",
                    text = BuildUserText(req, textContent: null)   // prompt + context only; doc carries the content
                }
            };
        }
        else
        {
            userContent = BuildUserText(req, textContent: decodedAttachmentText);
        }

        var body = new
        {
            model      = Model,
            max_tokens = 2048,
            system     = systemPrompt,
            messages   = new[] { new { role = "user", content = userContent } }
        };

        var request = new HttpRequestMessage(HttpMethod.Post, ApiUrl);
        request.Headers.Add("x-api-key", apiKey);
        request.Headers.Add("anthropic-version", "2023-06-01");
        request.Content = new StringContent(
            JsonSerializer.Serialize(body),
            Encoding.UTF8,
            "application/json");

        var response = await _http.SendAsync(request);
        var raw      = await response.Content.ReadAsStringAsync();

        if (!response.IsSuccessStatusCode)
        {
            _log.LogError("Claude API error {Status}: {Body}", response.StatusCode, raw);
            throw new HttpRequestException($"Claude API returned {response.StatusCode}");
        }

        using var doc  = JsonDocument.Parse(raw);
        var textBlock  = doc.RootElement
            .GetProperty("content")
            .EnumerateArray()
            .FirstOrDefault(b => b.GetProperty("type").GetString() == "text");

        var text = textBlock.TryGetProperty("text", out var t) ? t.GetString() ?? "" : "";

        // Strip any stray markdown fences
        text = System.Text.RegularExpressions.Regex
            .Replace(text.Trim(), @"^```(?:json)?\s*|\s*```$", "", 
                     System.Text.RegularExpressions.RegexOptions.Multiline)
            .Trim();

        return JsonSerializer.Deserialize<RfqExtraction>(text,
            new JsonSerializerOptions { PropertyNameCaseInsensitive = true });
    }

    /// <summary>
    /// Builds the user-turn text for the Claude prompt.
    /// <paramref name="textContent"/> overrides the default body/attachment content selection:
    /// pass the decoded attachment text for non-PDF/DOCX attachments, or null to fall back to
    /// req.Content (body path) or omit the content block (PDF/DOCX document API path).
    /// </summary>
    private static string BuildUserText(ExtractRequest req, string? textContent)
    {
        var sb = new StringBuilder();
        sb.AppendLine(ExtractionPrompt);

        if (!string.IsNullOrEmpty(req.BodyContext))
        {
            sb.AppendLine();
            sb.AppendLine("Email body context:");
            sb.AppendLine(req.BodyContext[..Math.Min(req.BodyContext.Length, 2000)]);
        }

        // Resolve what text to show in the content block:
        //   - Explicit override (decoded text attachment)         → use it
        //   - Body extraction or no base64 (EWS fallback path)   → use req.Content
        //   - PDF/DOCX sent via document API (textContent=null)  → omit block; doc carries content
        var content = textContent
            ?? (req.SourceType == "body" || string.IsNullOrEmpty(req.Base64Data)
                    ? req.Content ?? string.Empty
                    : null);

        if (content is not null)
        {
            sb.AppendLine();
            sb.AppendLine("Content:");
            sb.AppendLine("---");
            sb.AppendLine(content[..Math.Min(content.Length, 12_000)]);
            sb.AppendLine("---");
        }

        if (!string.IsNullOrEmpty(req.FileName))
            sb.AppendLine($"File name: {req.FileName}");

        sb.AppendLine();
        sb.AppendLine("Return ONLY the JSON object.");
        return sb.ToString();
    }
}
