namespace OutlookShredder.Proxy.Services;

/// <summary>
/// Thrown by an AI extraction service when the provider is temporarily overloaded
/// (HTTP 503/529 or SDK-level timeout due to high demand). The fallback coordinator
/// catches this to switch providers immediately rather than burning retries on a
/// busy service.
/// </summary>
public sealed class AiServiceOverloadedException(string provider, Exception? inner = null)
    : Exception($"AI provider '{provider}' is temporarily overloaded", inner)
{
    public string Provider { get; } = provider;
}
