using OutlookShredder.Proxy.Models;

namespace OutlookShredder.Proxy.Services;

/// <summary>
/// Wraps two <see cref="IAiExtractionService"/> instances with smart overload-aware routing:
///
///   1. Try primary.
///   2. If primary throws <see cref="AiServiceOverloadedException"/> → switch to secondary immediately.
///   3. If primary throws any other error → fall back to secondary (existing behaviour).
///   4. If secondary also throws <see cref="AiServiceOverloadedException"/> → retry primary once more
///      (it may have recovered in the time the secondary attempt took).
///   5. Any other secondary failure propagates to the caller.
///
/// Caller-initiated cancellation is honoured — the secondary is never tried after cancellation.
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
        _primary   = primary;
        _secondary = secondary;
        _log       = log;
    }

    public string ProviderName => $"{_primary.ProviderName} (fallback: {_secondary.ProviderName})";

    public async Task<RfqExtraction?> ExtractRfqAsync(ExtractRequest request, CancellationToken ct = default)
    {
        var primaryWasOverloaded = false;
        try
        {
            return await _primary.ExtractRfqAsync(request, ct);
        }
        catch (AiServiceOverloadedException) when (!ct.IsCancellationRequested)
        {
            primaryWasOverloaded = true;
            _log.LogInformation("[Fallback] {Primary} overloaded — switching immediately to {Secondary}",
                _primary.ProviderName, _secondary.ProviderName);
        }
        catch (Exception ex) when (!ct.IsCancellationRequested)
        {
            _log.LogWarning(ex, "[Fallback] {Primary} failed for ExtractRfq — retrying on {Secondary}",
                _primary.ProviderName, _secondary.ProviderName);
        }

        try
        {
            return await _secondary.ExtractRfqAsync(request, ct);
        }
        catch (AiServiceOverloadedException) when (!ct.IsCancellationRequested && primaryWasOverloaded)
        {
            _log.LogWarning("[Fallback] Both {Primary} and {Secondary} overloaded — retrying {Primary}",
                _primary.ProviderName, _secondary.ProviderName);
            return await _primary.ExtractRfqAsync(request, ct);
        }
        catch (Exception ex) when (!ct.IsCancellationRequested)
        {
            _log.LogWarning(ex, "[Fallback] {Secondary} also failed for ExtractRfq",
                _secondary.ProviderName);
            throw;
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
        var primaryWasOverloaded = false;
        try
        {
            return await _primary.ExtractPurchaseOrderAsync(
                base64Pdf, fileName, emailBodyContext, emailSubject, jobRefs, ct);
        }
        catch (AiServiceOverloadedException) when (!ct.IsCancellationRequested)
        {
            primaryWasOverloaded = true;
            _log.LogInformation("[Fallback] {Primary} overloaded — switching immediately to {Secondary}",
                _primary.ProviderName, _secondary.ProviderName);
        }
        catch (Exception ex) when (!ct.IsCancellationRequested)
        {
            _log.LogWarning(ex, "[Fallback] {Primary} failed for ExtractPurchaseOrder — retrying on {Secondary}",
                _primary.ProviderName, _secondary.ProviderName);
        }

        try
        {
            return await _secondary.ExtractPurchaseOrderAsync(
                base64Pdf, fileName, emailBodyContext, emailSubject, jobRefs, ct);
        }
        catch (AiServiceOverloadedException) when (!ct.IsCancellationRequested && primaryWasOverloaded)
        {
            _log.LogWarning("[Fallback] Both {Primary} and {Secondary} overloaded — retrying {Primary}",
                _primary.ProviderName, _secondary.ProviderName);
            return await _primary.ExtractPurchaseOrderAsync(
                base64Pdf, fileName, emailBodyContext, emailSubject, jobRefs, ct);
        }
        catch (Exception ex) when (!ct.IsCancellationRequested)
        {
            _log.LogWarning(ex, "[Fallback] {Secondary} also failed for ExtractPurchaseOrder",
                _secondary.ProviderName);
            throw;
        }
    }
}
