using OutlookShredder.Proxy.Models;

namespace OutlookShredder.Proxy.Services.Ai;

/// <summary>
/// Adapter for one AI provider that can produce an inquiry reply draft + classification. Register one per
/// provider (Claude, Gemini, …); <see cref="InquiryDraftService"/> tries them in registration order and falls
/// through on null. Switching/adding a provider = implement this + add a DI line, with no change to the
/// orchestrator or the inquiry pipeline. Mirrors the storage adapter seam (IInquiryStore/IMessageStore).
/// </summary>
public interface IInquiryDraftProvider
{
    /// <summary>Short provider name for logs (e.g. "Claude", "Gemini").</summary>
    string Name { get; }

    /// <summary>True when the provider has the config it needs (e.g. an API key). Unconfigured providers are
    /// skipped by the orchestrator so a deployment can run with any subset wired.</summary>
    bool IsConfigured { get; }

    /// <summary>Produces a draft, or null on failure / no usable output (the orchestrator then tries the next).</summary>
    Task<InquiryDraftResult?> DraftAsync(InquiryDraftInput input, CancellationToken ct = default);
}
