namespace OutlookShredder.Proxy.Services.Sms;

/// <summary>
/// Adapter for an outbound SMS carrier. Lets the inquiry pipeline send texts without binding to a specific
/// vendor SDK — swap carriers (the Twilio → SignalWire move was exactly this) or add a per-region carrier by
/// implementing this + changing one DI line. Mirrors the storage/AI adapter seams.
/// </summary>
public interface ISmsGateway
{
    string Name { get; }
    bool IsConfigured { get; }
    /// <summary>The sender number, when configured.</summary>
    string? FromNumber { get; }
    /// <summary>Sends a text; when <paramref name="statusCallback"/> is set the carrier POSTs delivery-status
    /// updates there. Supplying <paramref name="mediaUrls"/> (publicly-fetchable URLs the carrier downloads)
    /// sends an MMS. Returns the provider message SID, or null on failure.</summary>
    Task<string?> SendAsync(string to, string body, string? statusCallback = null,
        IReadOnlyList<string>? mediaUrls = null, CancellationToken ct = default);

    /// <summary>Fetches an inbound media part (MMS attachment) from the carrier's auth-protected media store.
    /// Returns (contentType, bytes), or null if unavailable. Used at ingest to store media durably.</summary>
    Task<(string ContentType, byte[] Bytes)?> DownloadMediaAsync(string mediaUrl, CancellationToken ct = default);

    /// <summary>Looks up a sent message's CURRENT delivery status by provider SID — for backfilling rows whose
    /// status callback never landed. Returns null on failure/not-found.</summary>
    Task<string?> GetStatusAsync(string sid, CancellationToken ct = default);
}
