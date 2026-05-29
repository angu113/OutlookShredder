using OutlookShredder.Proxy.Models;

namespace OutlookShredder.Proxy.Services;

/// <summary>
/// Alternates each extraction call between two <see cref="IAiExtractionService"/> instances
/// (round-robin) to spread load and exercise both providers. If the selected provider throws
/// after its own internal retries, the call is transparently retried against the other
/// provider — matching <see cref="FallbackAiExtractionService"/> cancellation semantics.
/// </summary>
/// <remarks>
/// Must be held as a singleton so the <see cref="_counter"/> persists across calls. A fresh
/// instance each call would restart from zero and never alternate.
/// </remarks>
public class RoundRobinAiExtractionService : IAiExtractionService
{
    private readonly IAiExtractionService _a;
    private readonly IAiExtractionService _b;
    private readonly ILogger<RoundRobinAiExtractionService> _log;
    private readonly TimeSpan _primaryTimeout;
    private long _counter = -1;

    public RoundRobinAiExtractionService(
        IAiExtractionService a,
        IAiExtractionService b,
        ILogger<RoundRobinAiExtractionService> log,
        IConfiguration? config = null)
    {
        _a = a;
        _b = b;
        _log = log;
        var seconds = int.TryParse(config?["AI:PrimaryTimeoutSeconds"], out var s) ? s : 30;
        _primaryTimeout = TimeSpan.FromSeconds(seconds);
    }

    public string ProviderName => $"round-robin({_a.ProviderName}, {_b.ProviderName})";

    public Task<RfqExtraction?> ExtractRfqAsync(ExtractRequest request, CancellationToken ct = default)
    {
        var (primary, secondary) = Pick();
        return InvokeWithFallback(
            "ExtractRfq",
            primary,
            secondary,
            (svc, innerCt) => svc.ExtractRfqAsync(request, innerCt),
            ct);
    }

    public Task<PoExtraction?> ExtractPurchaseOrderAsync(
        string base64Pdf,
        string fileName,
        string emailBodyContext,
        string emailSubject,
        List<string> jobRefs,
        CancellationToken ct = default)
    {
        var (primary, secondary) = Pick();
        return InvokeWithFallback(
            "ExtractPurchaseOrder",
            primary,
            secondary,
            (svc, innerCt) => svc.ExtractPurchaseOrderAsync(base64Pdf, fileName, emailBodyContext, emailSubject, jobRefs, innerCt),
            ct);
    }

    private (IAiExtractionService primary, IAiExtractionService secondary) Pick()
    {
        var n = Interlocked.Increment(ref _counter);
        return (n & 1) == 0 ? (_a, _b) : (_b, _a);
    }

    /// <summary>
    /// Tries the primary first. If it throws (after its own internal retries) OR doesn't
    /// return within <see cref="_primaryTimeout"/>, falls over to the secondary. The
    /// timeout exists because some SDK paths swallow 429s and retry internally for
    /// minutes — we don't want a quota-throttled provider to block the request when the
    /// other provider could answer in seconds (CLAUDE.md fail-fast rule).
    /// </summary>
    private async Task<T> InvokeWithFallback<T>(
        string op,
        IAiExtractionService primary,
        IAiExtractionService secondary,
        Func<IAiExtractionService, CancellationToken, Task<T>> invoke,
        CancellationToken ct)
    {
        _log.LogInformation("[RoundRobin] Routing {Op} to {Primary}", op, primary.ProviderName);

        using var primaryCts = CancellationTokenSource.CreateLinkedTokenSource(ct);
        var primaryTask = invoke(primary, primaryCts.Token);
        var timeoutTask = Task.Delay(_primaryTimeout, ct);
        var winner      = await Task.WhenAny(primaryTask, timeoutTask);

        if (winner == primaryTask)
        {
            try
            {
                return await primaryTask;
            }
            catch (Exception ex) when (!ct.IsCancellationRequested)
            {
                _log.LogWarning(ex,
                    "[RoundRobin] {Primary} failed for {Op} — retrying on {Secondary}",
                    primary.ProviderName, op, secondary.ProviderName);
                return await invoke(secondary, ct);
            }
        }

        // Timeout fired first — cancel the primary (best-effort; if the SDK ignores
        // cancellation it'll keep running but we'll move on) and fall over.
        _log.LogWarning(
            "[RoundRobin] {Primary} did not respond within {Timeout}s for {Op} — falling over to {Secondary}",
            primary.ProviderName, _primaryTimeout.TotalSeconds, op, secondary.ProviderName);
        primaryCts.Cancel();
        return await invoke(secondary, ct);
    }
}
