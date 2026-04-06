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

    private const string ApiUrl = "https://api.anthropic.com/v1/messages";

    private const string ExtractionPrompt = """
        You are extracting supplier quote data from a metals distribution RFQ email or attachment.
        Return ONLY a valid JSON object — no commentary, no markdown fences. Use null for absent fields.

        {
          "jobReference":          "6-character alphanumeric code from [XXXXXX] pattern, no brackets",
          "quoteReference":        "supplier's own quote/reference number, e.g. 'QP60600', 'Q-2024-1234' — see QUOTE REFERENCE rules below",
          "supplierName":          "company providing the quote — see SUPPLIER NAME rules below",
          "dateOfQuote":           "YYYY-MM-DD or null",
          "estimatedDeliveryDate": "YYYY-MM-DD or null — derive from lead time if no specific date",
          "freightTerms":          "verbatim freight terms, e.g. FOB Origin / Prepaid & Add / Included / null",
          "products": [
            {
              "productName":              "full spec — see PRODUCT NAME rules below",
              "unitsRequested":           number or null,
              "unitsQuoted":              number or null,
              "lengthPerUnit":            number or null,
              "lengthUnit":               "ft | m | in | mm | cm | null",
              "weightPerUnit":            number or null,
              "weightUnit":               "lb | kg | oz | g | null",
              "pricePerPound":            number or null,
              "pricePerFoot":             number or null,
              "pricePerPiece":            number or null,
              "totalPrice":               number or null,
              "leadTimeText":             "verbatim lead time as stated, e.g. '4-6 weeks ARO' or 'In Stock' or null",
              "certifications":           "e.g. MTR Included / ASTM / AMS / Certified / null",
              "supplierProductComments":  "all remaining notes — see COMMENTS rules below"
            }
          ]
        }

        ── QUOTE REFERENCE ────────────────────────────────────────────────────────
        The supplier's own internal reference number for this quote — distinct from the
        [XXXXXX] job reference which belongs to Metal Supermarkets.
        Look for: "Quote #", "Our Ref:", "Ref:", "Quote No.", "QP", "Q-", or similar
        prefixes in the subject line, document header, or body.
        Strip any prefix label — store only the reference value itself (e.g. "QP60600").
        If no supplier quote reference is present, use null.

        ── SUPPLIER NAME ──────────────────────────────────────────────────────────
        Search in this priority order:
        1. Company letterhead or quote header in the document/attachment.
        2. The "From:" email address display name (company portion, not the person's name).
        3. The email signature block (look for company name after the person's name/title).
        4. If this is a forwarded email, identify the original quoting company, not the forwarder.
        Use the company name only — not a person's first/last name.

        ── PRODUCT NAME ───────────────────────────────────────────────────────────
        Always build the product name to include ALL of the following that are present:
        1. Material grade: 304, 316L, 6061-T6, A36, 1018, Grade 2 Ti, etc.
        2. Product form: Flat Bar, Round Bar, Hex Bar, Angle, Channel, I-Beam, Round Tube,
           Square Tube, Rect Tube, Pipe, Sheet, Plate, Coil, Round Rod, Strip.
        3. ALL dimensions in the standard form for that product:
           - Bar/Strip:  thickness × width (e.g. 3/16" × 2")
           - Sheet/Plate: gauge or thickness × width × length (e.g. 11GA × 48" × 120")
           - Tube/Pipe:  OD × wall thickness (e.g. 2" OD × 0.120" wall)
           - Angle/Channel: leg × leg × thickness (e.g. 1-1/2" × 1-1/2" × 1/8")
        4. Length (e.g. "20' random lengths", "cut to 36\"", "12' mill lengths").
        5. Finish or condition if stated: 2B, #4, HR, CR, DOM, ERW, Annealed, T6.
        Example: "316L Stainless Round Tube 2\" OD × 0.120\" wall × 20' ERW"

        ── PRICING — work through ALL steps; leave a field null only if no path yields a value ──
        Step 1 — Direct: use $/lb, $/ft, $/piece if stated outright.
                  Also capture totalPrice directly if the document states a line total,
                  extended price, amount, subtotal, or extended amount for the line.
        Step 2 — $/cwt: if price is per hundred-weight (cwt), divide by 100 → $/lb.
        Step 3 — $/kg: divide by 2.20462 → $/lb.
        Step 4 — Unit price → $/lb: if pricePerPiece is known AND weightPerUnit is known,
                  YOU MUST compute pricePerPound = pricePerPiece ÷ weightPerUnit (convert weight to lb first).
                  Example: $61.75/pc ÷ 82 lb/pc = $0.7531/lb — always do this arithmetic.
        Step 5 — Unit price → $/ft: if pricePerPiece is known AND lengthPerUnit is known,
                  YOU MUST compute pricePerFoot = pricePerPiece ÷ lengthPerUnit (convert to ft first).
        Step 6 — Total → $/lb: if totalPrice and total weight are derivable,
                  YOU MUST compute pricePerPound = totalPrice ÷ (unitsQuoted × weightPerUnit).
        Step 7 — Total → $/ft: if totalPrice and total length are derivable,
                  YOU MUST compute pricePerFoot = totalPrice ÷ (unitsQuoted × lengthPerUnit in ft).
        Step 8 — Compute line total (forward): after steps 1-7, if totalPrice is still null, derive it:
          a. pricePerPiece × unitsQuoted                                              → totalPrice
          b. pricePerFoot × unitsQuoted × lengthPerUnit (convert to ft first)         → totalPrice
          c. pricePerPound × unitsQuoted × weightPerUnit (convert to lb first)        → totalPrice
          Use the first applicable option in that order (piece price is most direct).
        Prices are bare numbers — no $, commas, or currency symbols.

        ── QUANTITIES ─────────────────────────────────────────────────────────────
        - unitsRequested: pieces/bars/sheets asked for in the original RFQ.
        - unitsQuoted: what the supplier can actually supply (may be less — "can supply 50 of 100").
        - If the supplier gives no separate quantity, assume unitsQuoted = unitsRequested.
        - Linear-feet pricing: when quantity is expressed as total linear footage
          (e.g. "100 LF", "500 linear feet") rather than pieces, set unitsQuoted = that
          footage number, lengthPerUnit = 1, lengthUnit = "ft".
        - ALWAYS extract lengthPerUnit when pricing is per foot ($/ft, $/LF) — it is
          required to compute the line total. Check the product name, the email body, and
          the original RFQ spec. For standard mill bar/tube/structural lengths (e.g.
          "20' random lengths", "24' mill lengths") capture that length here.

        ── DATES ──────────────────────────────────────────────────────────────────
        - dateOfQuote: the date the supplier issued this quote.
        - estimatedDeliveryDate: when the material will arrive / be ready.
          Look for fields labelled: "Delivery Date", "Due Date", "Ship Date",
          "Lead Time", or "ARO" (After Receipt of Order).
          !! DO NOT use "Quote Valid Until", "Expiry", "Valid Through", or any
          quote-validity/expiry date — those describe how long the price is
          guaranteed, not when the goods will be delivered.
          Derivation rules (apply in priority order):
            1. Specific calendar date stated → use it directly (YYYY-MM-DD).
            2. Single lead time expressed in days (e.g. "10 days ARO", "5 days ARO")
               → count that many BUSINESS DAYS (Mon–Fri, skip Sat/Sun) forward from
               dateOfQuote (or today if dateOfQuote is unknown).
               e.g. if today is Wednesday and lead time is 3 days ARO → Monday.
            3. Single lead time expressed in weeks/months (e.g. "3 weeks ARO", "ships in 2 weeks")
               → add that duration in calendar weeks/months to dateOfQuote / today.
            4. Range lead time (e.g. "3-5 days ARO", "2-4 weeks")
               → use the LONGEST end of the range (most conservative estimate):
               "3-5 days ARO" → +5 business days; "2-4 weeks" → +4 calendar weeks.
            5. No delivery information present → null.
          ARO = After Receipt of Order; treat the receipt date as today / dateOfQuote.
        - Always populate leadTimeText with the verbatim lead time string regardless.

        ── MULTIPLE PRODUCTS ──────────────────────────────────────────────────────
        Every distinct grade, form, size, or finish is a SEPARATE entry.
        "1\" and 2\" flat bar" → two entries. "304 and 316 sheet" → two entries.

        ── COMMENTS ───────────────────────────────────────────────────────────────
        Capture in supplierProductComments: partial availability, alternates offered
        ("can supply 316 instead of 304"), spec deviations, cut charges, surcharges,
        minimum order quantities, certification details, freight notes, and any
        dimension detail not already encoded in productName.

        ── NO QUOTE FOUND ─────────────────────────────────────────────────────────
        If the email is an acknowledgement, out-of-office, or clearly not a price quote,
        return one entry with all numeric/date product fields null and explain in
        supplierProductComments. Always return at least one entry in products[].
        """;

    public ClaudeService(IConfiguration config, ILogger<ClaudeService> log)
    {
        _config = config;
        _log    = log;
        var timeoutSeconds = int.TryParse(_config["Claude:TimeoutSeconds"], out var t) ? t : 60;
        _http = new HttpClient { Timeout = TimeSpan.FromSeconds(timeoutSeconds) };
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

        var maxContentChars = int.TryParse(_config["Claude:MaxContentChars"], out var mcc) ? mcc : 12_000;
        var maxContextChars = int.TryParse(_config["Claude:MaxContextChars"], out var mcx) ? mcx : 2_000;

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
                    text = BuildUserText(req, textContent: null, maxContentChars, maxContextChars)
                }
            };
        }
        else
        {
            userContent = BuildUserText(req, textContent: decodedAttachmentText, maxContentChars, maxContextChars);
        }

        var model     = _config["Claude:Model"]     ?? "claude-sonnet-4-6";
        var maxTokens = int.TryParse(_config["Claude:MaxTokens"], out var mt) ? mt : 2048;

        var body = new
        {
            model      = model,
            max_tokens = maxTokens,
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
    private static string BuildUserText(
        ExtractRequest req, string? textContent, int maxContentChars, int maxContextChars)
    {
        var sb = new StringBuilder();
        sb.AppendLine(ExtractionPrompt);

        if (!string.IsNullOrEmpty(req.BodyContext))
        {
            sb.AppendLine();
            sb.AppendLine("Email body context:");
            sb.AppendLine(req.BodyContext[..Math.Min(req.BodyContext.Length, maxContextChars)]);
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
            sb.AppendLine(content[..Math.Min(content.Length, maxContentChars)]);
            sb.AppendLine("---");
        }

        if (!string.IsNullOrEmpty(req.FileName))
            sb.AppendLine($"File name: {req.FileName}");

        sb.AppendLine();
        sb.AppendLine("Return ONLY the JSON object.");
        return sb.ToString();
    }
}
