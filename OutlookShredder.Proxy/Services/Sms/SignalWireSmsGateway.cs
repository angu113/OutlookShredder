namespace OutlookShredder.Proxy.Services.Sms;

/// <summary>SignalWire-backed <see cref="ISmsGateway"/> — a thin adapter over <see cref="SignalWireService"/>
/// (the carrier connection). Swapping carriers = another ISmsGateway impl + a DI swap.</summary>
public sealed class SignalWireSmsGateway : ISmsGateway
{
    private readonly SignalWireService _sw;
    public SignalWireSmsGateway(SignalWireService sw) => _sw = sw;

    public string Name => "SignalWire";
    public bool IsConfigured => _sw.IsConfigured;
    public string? FromNumber => _sw.FromNumber;

    public Task<string?> SendAsync(string to, string body, string? statusCallback = null,
        IReadOnlyList<string>? mediaUrls = null, CancellationToken ct = default)
        => _sw.SendSmsAsync(to, body, statusCallback, mediaUrls, ct);

    public Task<(string ContentType, byte[] Bytes)?> DownloadMediaAsync(string mediaUrl, CancellationToken ct = default)
        => _sw.DownloadMediaAsync(mediaUrl, ct);

    public Task<string?> GetStatusAsync(string sid, CancellationToken ct = default)
        => _sw.GetMessageStatusAsync(sid, ct);
}
