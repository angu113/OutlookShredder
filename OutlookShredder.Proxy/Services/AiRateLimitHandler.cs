using System.Globalization;

namespace OutlookShredder.Proxy.Services;

/// <summary>
/// DelegatingHandler that reads rate-limit headers from AI provider responses
/// and feeds them into <see cref="AiRateLimitTracker"/>.
///
/// Claude: reads <c>anthropic-ratelimit-requests-remaining</c> and
///   <c>anthropic-ratelimit-requests-reset</c> from every response.
/// Gemini: reads <c>Retry-After</c> from 429 responses and logs any
///   <c>x-ratelimit-*</c> headers present (to discover proactive signals).
/// </summary>
public sealed class AiRateLimitHandler(
    string provider,
    AiRateLimitTracker tracker,
    ILogger<AiRateLimitHandler> log)
    : DelegatingHandler
{
    protected override async Task<HttpResponseMessage> SendAsync(
        HttpRequestMessage request, CancellationToken ct)
    {
        var response = await base.SendAsync(request, ct);

        if (provider == "Claude")
            ReadClaudeHeaders(response);
        else
            ReadGeminiHeaders(response);

        return response;
    }

    private void ReadClaudeHeaders(HttpResponseMessage response)
    {
        if (!TryGetHeaderInt(response, "anthropic-ratelimit-requests-remaining", out var remaining))
            return;

        TryGetHeaderDate(response, "anthropic-ratelimit-requests-reset", out var resetAt);
        tracker.UpdateClaude(remaining, resetAt);

        log.LogDebug("[RateLimit] Claude remaining={Remaining} reset={Reset}",
            remaining, resetAt == default ? "?" : resetAt.ToString("HH:mm:ss"));
    }

    private void ReadGeminiHeaders(HttpResponseMessage response)
    {
        // Log any x-ratelimit-* headers to discover what Gemini exposes.
        foreach (var h in response.Headers.Where(h =>
            h.Key.StartsWith("x-ratelimit", StringComparison.OrdinalIgnoreCase)))
        {
            log.LogDebug("[RateLimit] Gemini header: {Header}={Value}",
                h.Key, string.Join(",", h.Value));
        }

        // Parse proactive remaining/reset if present.
        if (TryGetHeaderInt(response, "x-ratelimit-remaining-requests", out var remaining))
        {
            TryGetHeaderDate(response, "x-ratelimit-reset-requests", out var resetAt);;
            tracker.UpdateGemini(remaining, resetAt);
        }

        // Reactive: parse Retry-After from 429.
        if ((int)response.StatusCode == 429 &&
            response.Headers.TryGetValues("retry-after", out var vals) &&
            int.TryParse(vals.First(), out var seconds))
        {
            tracker.SetGeminiRetryAfter(TimeSpan.FromSeconds(seconds));
        }
    }

    private static bool TryGetHeaderInt(HttpResponseMessage r, string name, out int value)
    {
        value = 0;
        return r.Headers.TryGetValues(name, out var vals)
            && int.TryParse(vals.First(), out value);
    }

    private static bool TryGetHeaderDate(HttpResponseMessage r, string name, out DateTimeOffset value)
    {
        value = default;
        return r.Headers.TryGetValues(name, out var vals)
            && DateTimeOffset.TryParse(vals.First(),
                CultureInfo.InvariantCulture, DateTimeStyles.AssumeUniversal, out value);
    }
}
