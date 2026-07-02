using OutlookShredder.Proxy.Models;

namespace OutlookShredder.Proxy.Services.Storage;

/// <summary>
/// DAO contract for the SMS customer-inquiry pipeline's portable entities — <see cref="Inquiry"/> and
/// <see cref="MessagingContact"/>. Deliberately storage-agnostic: no Graph / SharePoint types appear in the
/// surface, so the backing store can be swapped (e.g. SharePoint → Azure SQL) by providing another
/// implementation and changing one DI registration. <see cref="InquiryService"/> and the controllers depend
/// only on this interface, never on the SharePoint connection. The SharePoint implementation is
/// <see cref="SharePointInquiryStore"/>; an Azure SQL port adds an <c>AzureSqlInquiryStore</c> alongside it.
/// </summary>
public interface IInquiryStore
{
    /// <summary>Creates the backing tables/lists + indexes if absent. Called once at startup so the first
    /// inbound message isn't slowed by provisioning, and so index-at-construction holds (a brand-new
    /// filtered column can't be observed first).</summary>
    Task EnsureProvisionedAsync(CancellationToken ct = default);

    Task<MessagingContact?> GetContactAsync(string phone, CancellationToken ct = default);
    /// <summary>Inserts (when <see cref="MessagingContact.SpItemId"/> is null) or updates the contact.
    /// Returns the persisted row id.</summary>
    Task<int> UpsertContactAsync(MessagingContact contact, CancellationToken ct = default);

    /// <summary>The phone's inquiries, most-recent (LastMessageAt) first.</summary>
    Task<IReadOnlyList<Inquiry>> GetInquiriesByPhoneAsync(string phone, CancellationToken ct = default);
    /// <summary>All inquiries (most-recent first), optionally filtered by <paramref name="status"/> and a
    /// free-text <paramref name="query"/> over phone/id.</summary>
    Task<IReadOnlyList<Inquiry>> GetInquiriesAsync(string? status = null, string? query = null, CancellationToken ct = default);
    Task<Inquiry?> GetInquiryByIdAsync(string cinqId, CancellationToken ct = default);
    Task<int> CreateInquiryAsync(Inquiry inquiry, CancellationToken ct = default);
    Task UpdateInquiryAsync(Inquiry inquiry, CancellationToken ct = default);

    // Drafts — AI-suggested replies (Phase 2). Never auto-sent; an operator accepts/dismisses.
    Task<int> CreateDraftAsync(InquiryDraft draft, CancellationToken ct = default);
    Task<IReadOnlyList<InquiryDraft>> GetDraftsByInquiryAsync(string inquiryId, CancellationToken ct = default);
    Task UpdateDraftStatusAsync(int spItemId, string status, CancellationToken ct = default);

    // Notes (append-only) + linked quotations (HSK# references).
    Task<int> CreateNoteAsync(InquiryNote note, CancellationToken ct = default);
    Task<IReadOnlyList<InquiryNote>> GetNotesByInquiryAsync(string inquiryId, CancellationToken ct = default);
    Task<int> CreateQuotationAsync(InquiryQuotation quotation, CancellationToken ct = default);
    Task<IReadOnlyList<InquiryQuotation>> GetQuotationsByInquiryAsync(string inquiryId, CancellationToken ct = default);

    /// <summary>One-time identity backfill: rewrites the operator identity stored as <paramref name="fromName"/>
    /// to <paramref name="toName"/> across <c>Inquiries.AssignedTo</c>, <c>InquiryNotes.NoteAuthor</c> and
    /// <c>InquiryQuotations.LinkedBy</c> (case-insensitive match; rows already equal to the target are skipped).
    /// When <paramref name="apply"/> is false it only counts the matches (dry run — no writes). Returns the
    /// per-list match/patch counts plus the affected inquiry ids so the caller can refresh the cache.</summary>
    Task<IdentityBackfillResult> BackfillIdentityAsync(string fromName, string toName, bool apply, CancellationToken ct = default);

    /// <summary>One-time migration: populate native *Dt dateTime columns from the legacy text date columns
    /// across the inquiry lists. Idempotent. Returns (scanned, patched, failed).</summary>
    Task<(int Scanned, int Patched, int Failed)> BackfillDateTimeColumnsAsync(CancellationToken ct = default);
}

/// <summary>Result of <see cref="IInquiryStore.BackfillIdentityAsync"/> — matches found and (when applied) rows
/// patched per list, plus the distinct inquiry ids touched.</summary>
public sealed record IdentityBackfillResult(
    int AssignedMatched, int AssignedPatched,
    int NotesMatched,    int NotesPatched,
    int QuotesMatched,   int QuotesPatched,
    IReadOnlyList<string> AffectedInquiryIds);

/// <summary>
/// DAO contract for thread messages. This is the existing Messages store (shared with internal/email
/// messaging), surfaced as an abstraction so the inquiry pipeline depends on a seam rather than the
/// SharePoint connection directly. The SharePoint implementation is <see cref="SharePointMessageStore"/>.
/// </summary>
public interface IMessageStore
{
    Task EnsureProvisionedAsync(CancellationToken ct = default);
    /// <summary>Appends a message row; sets <see cref="MessageRecord.SpItemId"/> on success.</summary>
    Task AppendAsync(MessageRecord message, CancellationToken ct = default);
    /// <summary>Updates an outbound message's delivery status by its provider SID. Returns the owning
    /// InquiryId on success (so the caller can also refresh its in-memory cache — see
    /// InquiryService.UpdateMessageStatusAsync), or null if no row matched.</summary>
    Task<string?> UpdateStatusBySidAsync(string sid, string status, CancellationToken ct = default);
    /// <summary>The inquiry's messages, oldest-first, capped at <paramref name="take"/> — for AI thread context.</summary>
    Task<IReadOnlyList<MessageRecord>> GetByInquiryAsync(string inquiryId, int take = 20, CancellationToken ct = default);

    /// <summary>Stores one attachment's bytes durably under the inquiry (so previews/AI survive the carrier's
    /// short media-retention window). <paramref name="fileName"/> is an opaque, path-safe key.</summary>
    Task SaveMediaAsync(string inquiryId, string fileName, byte[] bytes, CancellationToken ct = default);

    /// <summary>Reads back a stored attachment (content-type + bytes), or null if absent.</summary>
    Task<(string ContentType, byte[] Bytes)?> GetMediaAsync(string inquiryId, string fileName, CancellationToken ct = default);

    /// <summary>Patches an existing message's body + media by provider SID (media backfill/recovery). False if
    /// no row matched.</summary>
    Task<bool> PatchBodyMediaBySidAsync(string sid, string body, string? mediaJson, CancellationToken ct = default);

    /// <summary>Sets one message's read flag by id (Phase 7 per-message toggle). False on miss.</summary>
    Task<bool> SetMessageReadAsync(int spItemId, bool read, CancellationToken ct = default);
    /// <summary>Sets the read flag on every message in an inquiry (Phase 7 mark-all). Returns the count patched.</summary>
    Task<int> SetAllReadByInquiryAsync(string inquiryId, bool read, CancellationToken ct = default);
}
