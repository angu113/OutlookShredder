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
    private long _counter = -1;

    public RoundRobinAiExtractionService(
        IAiExtractionService a,
        IAiExtractionService b,
        ILogger<RoundRobinAiExtractionService> log)
    {
        _a = a;
        _b = b;
        _log = log;
    }

    public string ProviderName => $"round-robin({_a.ProviderName}, {_b.ProviderName})";

    public Task<RfqExtraction?> ExtractRfqAsync(ExtractRequest request, CancellationToken ct = default)
    {
        var (primary, secondary) = Pick();
        return InvokeWithFallback(
            "ExtractRfq",
            primary,
            secondary,
            svc => svc.ExtractRfqAsync(request, ct),
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
            svc => svc.ExtractPurchaseOrderAsync(base64Pdf, fileName, emailBodyContext, emailSubject, jobRefs, ct),
            ct);
    }

    private (IAiExtractionService primary, IAiExtractionService secondary) Pick()
    {
        var n = Interlocked.Increment(ref _counter);
        return (n & 1) == 0 ? (_a, _b) : (_b, _a);
    }

    private async Task<T> InvokeWithFallback<T>(
        string op,
        IAiExtractionService primary,
        IAiExtractionService secondary,
        Func<IAiExtractionService, Task<T>> invoke,
        CancellationToken ct)
    {
        _log.LogInformation("[RoundRobin] Routing {Op} to {Primary}", op, primary.ProviderName);
        try
        {
            return await invoke(primary);
        }
        catch (Exception ex) when (!ct.IsCancellationRequested)
        {
            _log.LogWarning(ex,
                "[RoundRobin] {Primary} failed for {Op} — retrying on {Secondary}",
                primary.ProviderName, op, secondary.ProviderName);
            return await invoke(secondary);
        }
    }
}
