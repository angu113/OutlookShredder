using OutlookShredder.Proxy.Models;

namespace OutlookShredder.Proxy.Services;

/// <summary>
/// Wraps two <see cref="IAiExtractionService"/> instances. Calls the primary first; if the
/// primary throws after its own internal retries, logs a warning and retries on the secondary
/// once. Caller-initiated cancellation is honoured — the secondary is not tried when the
/// caller has cancelled.
/// </summary>
public class FallbackAiExtractionService : IAiExtractionService
{
    private readonly IAiExtractionService _primary;
    private readonly IAiExtractionService _secondary;
    private readonly ILogger<FallbackAiExtractionService> _log;

    public FallbackAiExtractionService(
        IAiExtractionService primary,
        IAiExtractionService secondary,
        ILogger<FallbackAiExtractionService> log)
    {
        _primary = primary;
        _secondary = secondary;
        _log = log;
    }

    public string ProviderName => $"{_primary.ProviderName} (fallback: {_secondary.ProviderName})";

    public async Task<RfqExtraction?> ExtractRfqAsync(ExtractRequest request, CancellationToken ct = default)
    {
        try
        {
            return await _primary.ExtractRfqAsync(request, ct);
        }
        catch (Exception ex) when (!ct.IsCancellationRequested)
        {
            _log.LogWarning(ex,
                "[Fallback] {Primary} failed for ExtractRfq — retrying on {Secondary}",
                _primary.ProviderName, _secondary.ProviderName);
            return await _secondary.ExtractRfqAsync(request, ct);
        }
    }

    public async Task<PoExtraction?> ExtractPurchaseOrderAsync(
        string base64Pdf,
        string fileName,
        string emailBodyContext,
        string emailSubject,
        List<string> jobRefs,
        CancellationToken ct = default)
    {
        try
        {
            return await _primary.ExtractPurchaseOrderAsync(
                base64Pdf, fileName, emailBodyContext, emailSubject, jobRefs, ct);
        }
        catch (Exception ex) when (!ct.IsCancellationRequested)
        {
            _log.LogWarning(ex,
                "[Fallback] {Primary} failed for ExtractPurchaseOrder — retrying on {Secondary}",
                _primary.ProviderName, _secondary.ProviderName);
            return await _secondary.ExtractPurchaseOrderAsync(
                base64Pdf, fileName, emailBodyContext, emailSubject, jobRefs, ct);
        }
    }
}
