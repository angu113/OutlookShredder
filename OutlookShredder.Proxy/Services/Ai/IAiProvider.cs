using OutlookShredder.Proxy.Models;

namespace OutlookShredder.Proxy.Services.Ai;

/// <summary>
/// Abstraction for AI providers that can extract RFQ data from emails and attachments.
/// Implementations should handle document types (PDF, DOCX), text attachments, and retry logic.
/// </summary>
public interface IAiProvider
{
    /// <summary>
    /// Gets the name of this provider (e.g., "claude", "gpt4", "llama").
    /// </summary>
    string Name { get; }

    /// <summary>
    /// Extracts structured RFQ data from an email or attachment using this AI provider.
    /// </summary>
    /// <param name="request">The extraction request containing email content, attachments, and RLI context.</param>
    /// <param name="cancellationToken">Cancellation token for long-running operations.</param>
    /// <returns>Extracted RFQ data, or null if extraction failed.</returns>
    Task<RfqExtraction?> ExtractAsync(ExtractRequest request, CancellationToken cancellationToken = default);
}
