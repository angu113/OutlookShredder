using OutlookShredder.Proxy.Models;

namespace OutlookShredder.Proxy.Services;

/// <summary>
/// Abstraction for AI-powered extraction services.
/// Implementations handle provider-specific API calls, retry logic, and response parsing.
/// </summary>
public interface IAiExtractionService
{
    /// <summary>Provider name for logging and diagnostics.</summary>
    string ProviderName { get; }

    /// <summary>
    /// Extracts supplier quote data from email content or attachments.
    /// </summary>
    /// <param name="request">Email content, attachments, RLI context, and metadata.</param>
    /// <param name="ct">Cancellation token.</param>
    /// <returns>Extracted RFQ data, or null if extraction failed.</returns>
    Task<RfqExtraction?> ExtractRfqAsync(ExtractRequest request, CancellationToken ct = default);

    /// <summary>
    /// Extracts purchase order data from a PDF attachment.
    /// </summary>
    /// <param name="base64Pdf">Base64-encoded PDF bytes.</param>
    /// <param name="fileName">Original PDF filename.</param>
    /// <param name="emailBodyContext">Short snippet from the email body for context.</param>
    /// <param name="emailSubject">Email subject line.</param>
    /// <param name="jobRefs">Pre-identified job references from regex scan.</param>
    /// <param name="ct">Cancellation token.</param>
    /// <returns>Extracted PO data, or null if extraction failed.</returns>
    Task<PoExtraction?> ExtractPurchaseOrderAsync(
        string base64Pdf,
        string fileName,
        string emailBodyContext,
        string emailSubject,
        List<string> jobRefs,
        CancellationToken ct = default);
}
