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
    IReadOnlyList<DraftAttachment>? Attachments = null);

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
        You are a sales assistant for Metal Supermarkets Hackensack, a metals distribution shop that cuts and
        supplies steel, aluminium, stainless and other metals to walk-in and trade customers. You are drafting
        a reply to an inbound customer SMS. Be warm, concise, and helpful — one or two short sentences suitable
        for a text message. Never invent prices, stock, or lead times; if a quote or stock check is needed, say
        the team will confirm. If the customer asks about an existing order, acknowledge and say we'll check the
        status. Do not promise anything you cannot know.
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
            AiModel    = model,
        };
    }

    private sealed class RawDraft
    {
        public string? Reply      { get; set; }
        public string? Intent     { get; set; }
        public string? Urgency    { get; set; }
        public bool    NeedsQuote { get; set; }
    }
}
