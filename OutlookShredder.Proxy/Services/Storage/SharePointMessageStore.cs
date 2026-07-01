using OutlookShredder.Proxy.Models;

namespace OutlookShredder.Proxy.Services.Storage;

/// <summary>
/// SharePoint-backed <see cref="IMessageStore"/>. A thin adapter over the existing Messages-list methods on
/// <see cref="SharePointService"/> (which owns the SharePoint connection + the shared Messages list), so the
/// inquiry pipeline talks to the <see cref="IMessageStore"/> seam instead of the connection class. An Azure
/// SQL port replaces this with a SQL-backed implementation and re-points the DI registration.
/// </summary>
public sealed class SharePointMessageStore : IMessageStore
{
    private readonly SharePointService _sp;
    public SharePointMessageStore(SharePointService sp) => _sp = sp;

    public Task EnsureProvisionedAsync(CancellationToken ct = default) => _sp.EnsureMessagesListAsync(ct);

    public Task AppendAsync(MessageRecord message, CancellationToken ct = default) => _sp.WriteMessageAsync(message, ct);

    public Task<string?> UpdateStatusBySidAsync(string sid, string status, CancellationToken ct = default)
        => _sp.UpdateMessageStatusBySidAsync(sid, status, ct);

    public async Task<IReadOnlyList<MessageRecord>> GetByInquiryAsync(string inquiryId, int take = 20, CancellationToken ct = default)
        => await _sp.ReadMessagesByInquiryAsync(inquiryId, take, ct);

    public Task SaveMediaAsync(string inquiryId, string fileName, byte[] bytes, CancellationToken ct = default)
        => _sp.UpsertInquiryMediaAsync(inquiryId, fileName, bytes, ct);

    public Task<(string ContentType, byte[] Bytes)?> GetMediaAsync(string inquiryId, string fileName, CancellationToken ct = default)
        => _sp.GetInquiryMediaAsync(inquiryId, fileName, ct);

    public Task<bool> PatchBodyMediaBySidAsync(string sid, string body, string? mediaJson, CancellationToken ct = default)
        => _sp.PatchMessageBodyMediaBySidAsync(sid, body, mediaJson, ct);

    public Task<bool> SetMessageReadAsync(int spItemId, bool read, CancellationToken ct = default)
        => _sp.SetMessageReadAsync(spItemId, read, ct);

    public Task<int> SetAllReadByInquiryAsync(string inquiryId, bool read, CancellationToken ct = default)
        => _sp.SetMessagesReadByInquiryAsync(inquiryId, read, ct);
}
