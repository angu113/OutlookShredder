using System.Text.RegularExpressions;
using OutlookShredder.Proxy.Models;

namespace OutlookShredder.Proxy.Services;

/// <summary>
/// Shared SHR-token ([SHR:{rfqId}]) routing step. Invoked by both the mail
/// poller and the add-in extract endpoint before AI extraction runs, so any
/// path that ingests a supplier reply can short-circuit it into
/// SupplierConversations instead of producing a duplicate SLI.
///
/// Resolution precedence (unchanged from the original inline block in
/// MailPollerService): sender-domain match against SupplierCacheService.DomainMap
/// (Suppliers SP list's ContactEmail column), with a substring fallback; then
/// the SR-cache fallback for historical RFQs whose supplier was established
/// during a prior reply.
/// </summary>
public sealed class ShrConvInRouter
{
    // Format: [SHR:HQABC123] (new 8-char) or [SHR:ABC123] (legacy 6-char).
    private static readonly Regex ShrTokenRegex =
        new(@"\[SHR:(HQ[A-Z0-9]{6}|[A-Z0-9]{6})\]",
            RegexOptions.Compiled | RegexOptions.IgnoreCase);

    private readonly SupplierCacheService      _suppliers;
    private readonly SharePointService         _sp;
    private readonly ILogger<ShrConvInRouter>  _log;

    public ShrConvInRouter(
        SupplierCacheService      suppliers,
        SharePointService         sp,
        ILogger<ShrConvInRouter>  log)
    {
        _suppliers = suppliers;
        _sp        = sp;
        _log       = log;
    }

    public sealed record Result(bool Routed, string? ShrRfqId, string? ResolvedSupplier)
    {
        public static readonly Result NotTagged = new(false, null, null);
        public static Result Unresolvable(string rfqId)                  => new(false, rfqId, null);
        public static Result WrittenToConv(string rfqId, string supplier) => new(true,  rfqId, supplier);
    }

    /// <summary>
    /// Inspects <paramref name="searchText"/> for an [SHR:{rfqId}] token. If present
    /// and the supplier resolves, writes a conv-in row to SupplierConversations and
    /// returns <see cref="Result.Routed"/>==true — callers should skip AI extraction.
    /// If the token is present but the supplier can't be resolved, returns the rfqId
    /// via <see cref="Result.ShrRfqId"/> so the caller can seed it into its own job
    /// references and fall through to extraction.
    /// </summary>
    public async Task<Result> TryRouteAsync(
        string         searchText,
        string         fromAddr,
        string         subject,
        string         body,
        string?        messageId,
        bool           hasAttachments,
        DateTimeOffset receivedAt,
        string?        contactEmail = null)
    {
        var m = ShrTokenRegex.Match(searchText ?? string.Empty);
        if (!m.Success) return Result.NotTagged;

        var rfqId    = m.Groups[1].Value.ToUpperInvariant();
        var supplier = ResolveSupplierFromEmail(fromAddr)
                       ?? await _sp.ResolveSupplierNameFromSrAsync(rfqId, fromAddr);

        if (supplier is null)
        {
            _log.LogInformation(
                "[SHR] Token [{RfqId}] from {From} — supplier unresolvable; caller will extract",
                rfqId, fromAddr);
            return Result.Unresolvable(rfqId);
        }

        _log.LogInformation(
            "[SHR] Token [{RfqId}] from {From} — routing to SupplierConversations (supplier={Supplier})",
            rfqId, fromAddr, supplier);

        await _sp.WriteConversationMessageAsync(new ConversationMessage
        {
            RfqId            = rfqId,
            SupplierName     = supplier,
            Direction        = "in",
            MessageId        = messageId,
            SentAt           = receivedAt,
            Subject          = subject,
            BodyText         = (body ?? string.Empty)[..Math.Min((body ?? string.Empty).Length, 4_000)],
            HasAttachments   = hasAttachments,
            ExtractedPricing = false,
            ContactEmail     = contactEmail ?? fromAddr,
        });

        return Result.WrittenToConv(rfqId, supplier);
    }

    /// <summary>
    /// Post-extraction hook: logs the inbound email as a conv-in row after the
    /// mail poller or add-in extract endpoint has finished writing SR+SLI.
    /// Lets the conversation viewer treat SC as the single source of thread truth
    /// — SR+SLI become pricing-only records, never read for thread display.
    /// WriteConversationMessageAsync dedupes on MessageId, so this is safe to call
    /// even when the SHR bypass already wrote a row for the same message.
    /// </summary>
    public async Task WriteConvInFromExtractionAsync(
        string?        rfqId,
        string?        supplierName,
        string?        messageId,
        string?        subject,
        string?        body,
        DateTimeOffset receivedAt,
        bool           hasAttachments,
        string?        fromAddr)
    {
        if (string.IsNullOrEmpty(rfqId)       ||
            string.IsNullOrEmpty(supplierName) ||
            string.IsNullOrEmpty(messageId))
        {
            // Need all three to produce a useful thread entry. Drop silently so
            // "orphan" / unknown-supplier extractions don't clutter the conv list.
            return;
        }

        var b = body ?? string.Empty;
        await _sp.WriteConversationMessageAsync(new ConversationMessage
        {
            RfqId            = rfqId,
            SupplierName     = supplierName,
            Direction        = "in",
            MessageId        = messageId,
            SentAt           = receivedAt,
            Subject          = subject,
            BodyText         = b[..Math.Min(b.Length, 4_000)],
            HasAttachments   = hasAttachments,
            ExtractedPricing = true,
            ContactEmail     = fromAddr,
        });
    }

    // Sender-domain resolver — identical semantics to the prior
    // MailPollerService.ResolveSupplierFromEmail helper. Skips our own
    // mithrilmetals.com domain so we never record ourselves as a supplier.
    private string? ResolveSupplierFromEmail(string fromAddr)
    {
        if (string.IsNullOrWhiteSpace(fromAddr)) return null;
        var at = fromAddr.IndexOf('@');
        if (at < 0) return null;
        var domain = fromAddr[(at + 1)..].ToLowerInvariant();

        if (domain.Equals("mithrilmetals.com", StringComparison.OrdinalIgnoreCase))
            return null;

        if (_suppliers.DomainMap.TryGetValue(domain, out var name)) return name;
        return _suppliers.ResolveByDomainSubstring(domain);
    }
}
