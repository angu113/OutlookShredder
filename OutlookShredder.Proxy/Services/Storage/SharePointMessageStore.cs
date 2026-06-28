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

    public Task<bool> UpdateStatusBySidAsync(string sid, string status, CancellationToken ct = default)
        => _sp.UpdateMessageStatusBySidAsync(sid, status, ct);
}
