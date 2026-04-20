namespace OutlookShredder.Proxy.Services;

/// <summary>
/// Singleton that tracks org-wide rate-limit capacity for each AI provider,
/// populated from response headers after every call.  All five proxy instances
/// running in the same organisation see the same remaining-capacity numbers
/// returned by the API, so each instance independently backs off when the
/// shared quota is nearly exhausted — no cross-process coordination needed.
///
/// Claude: proactive — reads <c>anthropic-ratelimit-requests-remaining</c> /
///   <c>anthropic-ratelimit-requests-reset</c> from every response.
/// Gemini: reactive — reads <c>Retry-After</c> from 429 responses; also logs
///   any <c>x-ratelimit-*</c> headers to identify future proactive signals.
/// </summary>
public sealed class AiRateLimitTracker
{
    private readonly ILogger<AiRateLimitTracker> _log;

    private int  _claudeRemaining = int.MaxValue;
    private long _claudeResetMs   = 0;

    private int  _geminiRemaining = int.MaxValue;
    private long _geminiResetMs   = 0;

    // Pause new calls when remaining drops at or below this value.
    private const int PauseThreshold = 3;

    public AiRateLimitTracker(ILogger<AiRateLimitTracker> log) => _log = log;

    // ── Updaters (called from DelegatingHandler) ──────────────────────────────

    public void UpdateClaude(int remaining, DateTimeOffset? resetAt)
    {
        Volatile.Write(ref _claudeRemaining, remaining);
        if (resetAt.HasValue)
            Interlocked.Exchange(ref _claudeResetMs, resetAt.Value.ToUnixTimeMilliseconds());

        if (remaining <= PauseThreshold)
            _log.LogWarning("[RateLimit] Claude org quota low: {Remaining} requests remaining, resets at {Reset}",
                remaining, resetAt?.ToString("HH:mm:ss") ?? "?");
    }

    public void UpdateGemini(int remaining, DateTimeOffset? resetAt)
    {
        Volatile.Write(ref _geminiRemaining, remaining);
        if (resetAt.HasValue)
            Interlocked.Exchange(ref _geminiResetMs, resetAt.Value.ToUnixTimeMilliseconds());

        if (remaining <= PauseThreshold)
            _log.LogWarning("[RateLimit] Gemini quota low: {Remaining} requests remaining, resets at {Reset}",
                remaining, resetAt?.ToString("HH:mm:ss") ?? "?");
    }

    /// <summary>Called when Gemini returns 429 with a <c>Retry-After</c> header.</summary>
    public void SetGeminiRetryAfter(TimeSpan delay)
    {
        var resetMs = DateTimeOffset.UtcNow.Add(delay).ToUnixTimeMilliseconds();
        Interlocked.Exchange(ref _geminiResetMs, resetMs);
        Volatile.Write(ref _geminiRemaining, 0);
        _log.LogWarning("[RateLimit] Gemini 429 — pausing for {Seconds}s (Retry-After)", (int)delay.TotalSeconds);
    }

    // ── Pre-call throttle check ───────────────────────────────────────────────

    /// <summary>
    /// Awaits until the provider's rate limit window resets when remaining capacity
    /// has dropped to or below <see cref="PauseThreshold"/>.
    /// </summary>
    public async Task ThrottleIfNeededAsync(string provider, CancellationToken ct)
    {
        var remaining = provider == "Claude"
            ? Volatile.Read(ref _claudeRemaining)
            : Volatile.Read(ref _geminiRemaining);

        if (remaining > PauseThreshold) return;

        var resetMs = provider == "Claude"
            ? Interlocked.Read(ref _claudeResetMs)
            : Interlocked.Read(ref _geminiResetMs);

        var waitMs = resetMs - DateTimeOffset.UtcNow.ToUnixTimeMilliseconds();
        if (waitMs <= 0 || waitMs > 120_000) return;  // stale or unreasonably long

        // Spread concurrent waiters across a 2-second jitter window after the reset
        // so 40 slots (8 per instance × 5 instances) don't all fire simultaneously.
        var jitterMs = Random.Shared.Next(0, 2_000);
        _log.LogInformation("[RateLimit] {Provider} pausing {WaitMs}ms + {Jitter}ms jitter — quota at {Remaining}",
            provider, waitMs, jitterMs, remaining);
        await Task.Delay((int)waitMs + jitterMs, ct);

        // Do NOT reset remaining here — let the next API response header update it.
        // This means late-waking waiters re-enter ThrottleIfNeededAsync, see resetMs is
        // now in the past (waitMs <= 0), and pass through immediately rather than stacking
        // up a second wave.
    }

    // ── Status (for /api/mail/status or /api/health) ──────────────────────────

    public (int Remaining, DateTimeOffset? ResetAt) ClaudeStatus =>
        (Volatile.Read(ref _claudeRemaining),
         _claudeResetMs == 0 ? null : DateTimeOffset.FromUnixTimeMilliseconds(Interlocked.Read(ref _claudeResetMs)));

    public (int Remaining, DateTimeOffset? ResetAt) GeminiStatus =>
        (Volatile.Read(ref _geminiRemaining),
         _geminiResetMs == 0 ? null : DateTimeOffset.FromUnixTimeMilliseconds(Interlocked.Read(ref _geminiResetMs)));
}
