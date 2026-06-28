using System.Text.Json;
using System.Text.RegularExpressions;
using Microsoft.AspNetCore.Mvc;
using OutlookShredder.Proxy.Services;

namespace OutlookShredder.Proxy.Controllers;

/// <summary>
/// Diagnostic traces for the RFQ extraction pipeline. Read-only.
/// </summary>
[ApiController]
public class DiagController : ControllerBase
{
    private readonly SharePointService        _sp;
    private readonly MailService              _mail;
    private readonly SharePointContractCheckService _contract;
    private readonly IConfiguration           _config;
    private readonly ILogger<DiagController>  _log;

    public DiagController(SharePointService sp, MailService mail, SharePointContractCheckService contract,
        IConfiguration config, ILogger<DiagController> log)
    {
        _sp = sp; _mail = mail; _contract = contract; _config = config; _log = log;
    }

    /// <summary>Runs the SharePoint data-contract self-check on demand (write→read→assert each typed field
    /// round-trips through the real Graph boundary). Use after touching any SP-backed schema.</summary>
    [HttpGet("/api/diag/sp-contract")]
    public async Task<IActionResult> SpContract(CancellationToken ct)
    {
        var r = await _contract.RunAsync(ct);
        return r.AllOk ? Ok(r) : StatusCode(500, r);
    }

    private static readonly Regex BracketRef = new(@"\[([A-Za-z]{2}[A-Za-z0-9]{4})\]", RegexOptions.Compiled);

    /// <summary>
    /// Per-RFQ extraction trace: for each supplier response under the RFQ, lays out every pipeline step
    /// with its data AND its source so a value can be found and verified —
    ///   • the LIVE email (subject / sender / real attachment file names; bytes shown as placeholders)
    ///   • the AI extraction audit (SupplierResponses.ClaudeResponseLog, the raw per-source extractions)
    ///   • the job-reference resolution (email-subject ref vs PDF-extracted ref → final RFQ; mismatch flag)
    ///   • the stored SupplierResponses row + the stored attachment's drive path
    ///   • the resulting SupplierLineItems rows
    /// Every block carries a `source` (MessageId / SP list + item id / drive path).
    /// </summary>
    [HttpGet("/api/diag/extraction-trace")]
    public async Task<IActionResult> ExtractionTrace([FromQuery] string rfqId)
    {
        if (string.IsNullOrWhiteSpace(rfqId)) return BadRequest(new { error = "rfqId is required" });

        try
        {
            var mailbox = _config["Mail:MailboxAddress"] ?? "";
            var sli     = await _sp.ReadSupplierItemsByRfqIdAsync(rfqId);

            static string? S(Dictionary<string, object?> d, string k) => d.TryGetValue(k, out var v) ? v?.ToString() : null;

            var traces = new List<object>();
            foreach (var grp in sli.GroupBy(r => S(r, "SupplierResponseId") ?? "(none)"))
            {
                var srId      = grp.Key;
                var s         = grp.First();
                var messageId = S(s, "MessageId");

                // ── Step 1: the live email (source of truth) ──────────────────────────
                object liveEmail;
                if (!string.IsNullOrEmpty(mailbox) && !string.IsNullOrEmpty(messageId))
                {
                    var msg = await _mail.GetMessageByIdAsync(mailbox, messageId!);
                    if (msg is not null)
                    {
                        var atts = await _mail.GetAttachmentMetaAsync(messageId!);
                        liveEmail = new
                        {
                            source           = new { mailbox, messageId },
                            found            = true,
                            subject          = msg.Subject,
                            from             = msg.From?.EmailAddress?.Address,
                            receivedDateTime = msg.ReceivedDateTime,
                            hasAttachments   = msg.HasAttachments,
                            attachments      = atts.Select(a => new
                            {
                                fileName    = a.Name,
                                bytes       = $"<{a.Size} bytes — not shown>",
                                contentType = a.ContentType,
                            }).ToList(),
                            bodyPreview = Trunc(msg.BodyPreview, 400),
                        };
                    }
                    else liveEmail = new { source = new { mailbox, messageId }, found = false, note = "message no longer in mailbox" };
                }
                else liveEmail = new { source = new { mailbox, messageId }, found = false, note = "no mailbox / messageId on record" };

                // ── Step 2: AI extraction audit (ClaudeResponseLog NDJSON) ────────────
                var log       = (string.IsNullOrEmpty(srId) || srId == "(none)") ? null : await _sp.ReadClaudeResponseLogAsync(srId);
                var aiEntries = ParseNdjson(log);

                // ── Step 3: job-reference resolution ──────────────────────────────────
                var emailSubject = S(s, "EmailSubject") ?? "";
                var emailRef     = BracketRef.Match(emailSubject) is { Success: true } m ? m.Groups[1].Value.ToUpperInvariant() : null;
                // The job ref the PDF itself claims comes from the *attachment-sourced* extraction — not a later
                // body reprocess (which echoes the email subject and would mask a misroute). Track both separately.
                string? pdfRef = null, bodyRef = null;
                foreach (var e in aiEntries)
                {
                    if (!(e.TryGetProperty("Ext", out var ext) && ext.TryGetProperty("JobReference", out var jr)
                          && jr.ValueKind == JsonValueKind.String && !string.IsNullOrWhiteSpace(jr.GetString()))) continue;
                    var entSrc = e.TryGetProperty("Src", out var sp) && sp.ValueKind == JsonValueKind.String ? sp.GetString() : null;
                    if (string.Equals(entSrc, "attachment", StringComparison.OrdinalIgnoreCase)) pdfRef = jr.GetString();
                    else bodyRef = jr.GetString();
                }

                // ── Step 4: stored attachment ─────────────────────────────────────────
                var sourceFile = S(s, "SourceFile");
                object? storedAttachment = string.IsNullOrWhiteSpace(sourceFile) ? null : new
                {
                    fileName = sourceFile,
                    source   = new { drivePath = $"QuoteAttachments/{srId}/{sourceFile}" },
                    bytes    = "<not shown — fetch the drive path to verify>",
                };

                traces.Add(new
                {
                    supplierResponse = new
                    {
                        source           = new { list = "SupplierResponses", itemId = srId },
                        supplierName     = S(s, "SupplierName"),
                        rfqId            = S(s, "RFQ_ID"),
                        processingSource = S(s, "ProcessingSource"),
                        quoteReference   = S(s, "QuoteReference"),
                        receivedAt       = S(s, "ReceivedAt"),
                        processedAt      = S(s, "ProcessedAt"),
                    },
                    liveEmail,
                    aiExtraction = new
                    {
                        source  = new { list = "SupplierResponses", field = "ClaudeResponseLog", itemId = srId },
                        entries = aiEntries,
                    },
                    jobRefResolution = new
                    {
                        emailSubjectRef = emailRef,
                        pdfExtractedRef = pdfRef,
                        finalRfqId      = S(s, "JobReference") ?? rfqId,
                        // A genuine mismatch requires BOTH a bracketed subject ref AND a PDF job ref that
                        // disagree. A PDF ref with no subject bracket (e.g. a supplier's quote-titled email)
                        // is normal PDF-led routing, not a conflict — don't flag it. Mirrors the write-path
                        // jobRefMismatchWarning, which only fires when the subject regex ref is non-empty.
                        mismatch        = !string.IsNullOrWhiteSpace(pdfRef)
                                          && !string.IsNullOrWhiteSpace(emailRef)
                                          && !string.Equals(pdfRef, emailRef, StringComparison.OrdinalIgnoreCase),
                    },
                    storedAttachment,
                    lineItems = grp.Select(r => new
                    {
                        source       = new { list = "SupplierLineItems", itemId = S(r, "SpItemId") },
                        productName  = S(r, "ProductName"),
                        mspc         = S(r, "ProductSearchKey"),
                        sourceFile   = S(r, "SourceFile"),
                        pricePerPound = S(r, "PricePerPound"),
                        totalPrice   = S(r, "TotalPrice"),
                        isRegret     = S(r, "IsRegret"),
                        isSubstitute = S(r, "IsSubstitute"),
                    }).ToList(),
                });
            }

            return Ok(new { rfqId, mailbox, supplierResponseCount = traces.Count, traces });
        }
        catch (Exception ex)
        {
            _log.LogError(ex, "[Diag] extraction-trace failed for {RfqId}", rfqId);
            return StatusCode(500, new { error = ex.Message });
        }
    }

    private static string? Trunc(string? s, int n) => s is null ? null : (s.Length <= n ? s : s[..n] + "…");

    /// <summary>Parses an NDJSON audit log into a list of JSON objects (one per AI extraction), skipping
    /// any unparseable line. Each element is detached from its document so it serializes cleanly.</summary>
    private static List<JsonElement> ParseNdjson(string? ndjson)
    {
        var result = new List<JsonElement>();
        if (string.IsNullOrWhiteSpace(ndjson)) return result;
        foreach (var line in ndjson.Split('\n', StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries))
        {
            try { using var doc = JsonDocument.Parse(line); result.Add(doc.RootElement.Clone()); }
            catch { /* skip unparseable line */ }
        }
        return result;
    }
}
