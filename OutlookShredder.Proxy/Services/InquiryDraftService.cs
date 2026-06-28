using System.Text;
using System.Text.Json;
using OutlookShredder.Proxy.Models;
using OutlookShredder.Proxy.Services.Ai;

namespace OutlookShredder.Proxy.Services;

/// <summary>Input for one AI draft: the triggering inbound text, the prior thread transcript, any linked
/// HSK# / operator notes for context, and any inbound attachments (images/PDFs the customer sent — fed to
/// the model so it can read a sketch or spec sheet when drafting).</summary>
public sealed record InquiryDraftInput(
    string InboundBody,
    string Transcript,
    IReadOnlyList<string> LinkedHsk,
    string? Notes,
    IReadOnlyList<DraftAttachment>? Attachments = null,
    string? CatalogContext = null);

/// <summary>One attachment passed to the model (image or PDF only — CAD/other are not vision-readable).</summary>
public sealed record DraftAttachment(string MimeType, string Base64, string? FileName);

/// <summary>
/// Generates a suggested reply + intent/urgency/needsQuote classification for an inbound customer message.
/// This is the orchestrator: it owns no transport, just a prioritised chain of <see cref="IInquiryDraftProvider"/>
/// adapters (Claude, Gemini, …) injected in registration order. It returns the first provider's result and
/// falls through to the next on null — so switching or adding a provider is "implement the interface + register
/// one line", with no change here or in <see cref="InquiryService"/>. The draft is ALWAYS a suggestion (never
/// auto-sent); on total failure it returns null and the pipeline simply skips the suggestion, so a model outage
/// never blocks ingest.
/// </summary>
public sealed class InquiryDraftService
{
    private readonly IReadOnlyList<IInquiryDraftProvider> _providers;
    private readonly ILogger<InquiryDraftService>         _log;

    public InquiryDraftService(IEnumerable<IInquiryDraftProvider> providers, ILogger<InquiryDraftService> log)
    {
        _providers = providers.ToList();
        _log       = log;
    }

    public async Task<InquiryDraftResult?> DraftAsync(InquiryDraftInput input, CancellationToken ct = default)
    {
        foreach (var provider in _providers)
        {
            if (!provider.IsConfigured) continue;
            try
            {
                var result = await provider.DraftAsync(input, ct);
                if (result is not null) return result;
                _log.LogWarning("[InquiryDraft] provider {Name} produced no draft — trying next", provider.Name);
            }
            catch (Exception ex)
            {
                _log.LogWarning(ex, "[InquiryDraft] provider {Name} threw — trying next", provider.Name);
            }
        }
        _log.LogWarning("[InquiryDraft] no provider produced a draft");
        return null;
    }
}

/// <summary>Pure, unit-testable prompt assembly + result coercion shared by every draft provider (the
/// transport-agnostic glue around the AI call).</summary>
public static class InquiryDraftPrompt
{
    public static readonly string[] Intents   = ["Quote", "OrderStatus", "Question", "Complaint", "Other"];
    public static readonly string[] Urgencies = ["Low", "Normal", "High"];

    public static readonly JsonSerializerOptions JsonOpts = new() { PropertyNameCaseInsensitive = true };

    public static readonly (int MinMs, int MaxMs)[] RetryDelays =
        [(2_000, 4_000), (5_000, 10_000), (15_000, 25_000)];

    public static Task DelayAsync(int attempt, CancellationToken ct)
    {
        var (min, max) = RetryDelays[Math.Min(Math.Max(attempt, 0), RetryDelays.Length - 1)];
        return Task.Delay(Random.Shared.Next(min, max), ct);
    }

    // Stub shop-knowledge prompt — real copy is tracked in wip/customer-experience-sms-inquiry.md
    // ("Content still needed: shop-knowledge system prompt"). Override at runtime via InquiryDraft:SystemPrompt.
    private const string DefaultSystemPrompt =
        """
        You are a sales assistant for Metal Supermarkets Hackensack, a metals shop that cuts and supplies steel,
        aluminium, stainless and other metals to walk-in and trade customers. You draft SMS replies to inbound
        customer texts. Be warm and concise — one or two short sentences fit for a text. Never invent prices,
        stock, or lead times; if a quote or stock check is needed, say the team will confirm.

        On a PRODUCT request your job is to decide whether we have enough to identify the exact product, and if
        not, ask for the SINGLE most important missing detail (echo the customer's own dimensions + quantity so
        they know you understood). You may be given the closest catalog product families — use them.

        Read terse requests with product understanding, not just the literal words:
        - MATERIAL is required. If the text doesn't say Steel, Stainless, or Aluminum, ask which — and list those
          three in the `options` field.
        - "box" / "square tube" with ONE dimension is a SQUARE cross-section: "2 box" = 2x2 square tube. Do NOT
          ask for a second dimension — only the WALL THICKNESS is missing.
        - "angle" with ONE dimension is an EQUAL angle: "2 angle" = 2x2 angle. Only the THICKNESS is missing.
        - Rectangular tube and unequal angle genuinely need both face dimensions.
        - If two details are truly missing (e.g. material AND thickness), ask both in one short message.
        Always move the conversation toward a quote/sale: while details are missing, ask for them; once the
        product is fully specified, confirm we'll price it up / the team will send a quote. Each reply should
        advance the customer's requirements, not just acknowledge.
        Use the `options` field only for a small discrete choice (e.g. material); leave it empty otherwise. If the
        message isn't a product request (order status, general question), a brief helpful acknowledgement is fine.
        """;

    public static string SystemPrompt(IConfiguration config)
        => config["InquiryDraft:SystemPrompt"] is { Length: > 0 } sp ? sp : DefaultSystemPrompt;

    /// <summary>Renders thread messages as a transcript ("Customer:" / "Shop:" lines), oldest-first, keeping
    /// the most recent <paramref name="max"/>. Inbound (Direction="in") is the customer; everything else is us.</summary>
    public static string BuildTranscript(IEnumerable<MessageRecord> messages, int max = 12)
    {
        var ordered = messages.ToList();
        if (ordered.Count > max) ordered = ordered.GetRange(ordered.Count - max, max);
        var sb = new StringBuilder();
        foreach (var m in ordered)
        {
            var who  = string.Equals(m.Direction, "in", StringComparison.OrdinalIgnoreCase) ? "Customer" : "Shop";
            var line = (m.Body ?? "").Replace("\r", " ").Replace("\n", " ").Trim();
            if (line.Length > 0) sb.Append(who).Append(": ").AppendLine(line);
        }
        return sb.ToString().TrimEnd();
    }

    public static string BuildUserText(InquiryDraftInput input)
    {
        var sb = new StringBuilder();
        if (!string.IsNullOrWhiteSpace(input.Transcript))
            sb.AppendLine("Conversation so far:").AppendLine(input.Transcript).AppendLine();
        if (input.LinkedHsk.Count > 0)
            sb.Append("Linked order/quote refs: ").AppendLine(string.Join(", ", input.LinkedHsk));
        if (!string.IsNullOrWhiteSpace(input.Notes))
            sb.Append("Operator notes: ").AppendLine(input.Notes);
        if (!string.IsNullOrWhiteSpace(input.CatalogContext))
            sb.AppendLine("Closest catalog product families (compare the request to these to spot the missing detail):")
              .AppendLine(input.CatalogContext).AppendLine();
        sb.AppendLine().AppendLine("Latest customer message:").Append(input.InboundBody);
        return sb.ToString();
    }

    public static string CoerceIntent(string? v)
        => Intents.FirstOrDefault(i => string.Equals(i, v?.Trim(), StringComparison.OrdinalIgnoreCase)) ?? "Other";

    public static string CoerceUrgency(string? v)
        => Urgencies.FirstOrDefault(u => string.Equals(u, v?.Trim(), StringComparison.OrdinalIgnoreCase)) ?? "Normal";

    public static InquiryDraftResult? MapResult(string json, string model)
    {
        var raw = JsonSerializer.Deserialize<RawDraft>(json, JsonOpts);
        if (raw is null || string.IsNullOrWhiteSpace(raw.Reply)) return null;
        return new InquiryDraftResult
        {
            Reply      = raw.Reply.Trim(),
            Intent     = CoerceIntent(raw.Intent),
            Urgency    = CoerceUrgency(raw.Urgency),
            NeedsQuote = raw.NeedsQuote,
            Options    = (raw.Options ?? []).Where(o => !string.IsNullOrWhiteSpace(o)).Select(o => o.Trim()).Take(4).ToList(),
            AiModel    = model,
        };
    }

    private sealed class RawDraft
    {
        public string?       Reply      { get; set; }
        public string?       Intent     { get; set; }
        public string?       Urgency    { get; set; }
        public bool          NeedsQuote { get; set; }
        public List<string>? Options    { get; set; }
    }
}
