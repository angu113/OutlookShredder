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
    /// updates there. Returns the provider message SID, or null on failure.</summary>
    Task<string?> SendAsync(string to, string body, string? statusCallback = null, CancellationToken ct = default);
}
