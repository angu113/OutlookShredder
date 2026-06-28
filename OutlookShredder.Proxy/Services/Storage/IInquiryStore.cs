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
    Task<Inquiry?> GetInquiryByIdAsync(string cinqId, CancellationToken ct = default);
    Task<int> CreateInquiryAsync(Inquiry inquiry, CancellationToken ct = default);
    Task UpdateInquiryAsync(Inquiry inquiry, CancellationToken ct = default);
}

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
    /// <summary>Updates an outbound message's delivery status by its provider SID. False if no row matched.</summary>
    Task<bool> UpdateStatusBySidAsync(string sid, string status, CancellationToken ct = default);
}
