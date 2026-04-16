namespace OutlookShredder.Proxy.Services.Ai;

/// <summary>
/// Factory for creating AI provider instances.
/// Supports registration and resolution of multiple provider implementations.
/// </summary>
public interface IAiProviderFactory
{
    /// <summary>
    /// Gets an AI provider by name.
    /// </summary>
    /// <param name="name">Provider name (e.g., "claude", "gpt4").</param>
    /// <returns>The requested provider, or null if not registered.</returns>
    IAiProvider? GetProvider(string name);

    /// <summary>
    /// Gets the default AI provider.
    /// </summary>
    IAiProvider GetDefaultProvider();

    /// <summary>
    /// Gets all registered provider names.
    /// </summary>
    IEnumerable<string> GetAvailableProviders();
}

/// <summary>
/// Default implementation that uses dependency injection to resolve providers.
/// </summary>
public class AiProviderFactory : IAiProviderFactory
{
    private readonly IServiceProvider _serviceProvider;
    private readonly Dictionary<string, Type> _providers;
    private readonly string _defaultProviderName;

    public AiProviderFactory(IServiceProvider serviceProvider, Dictionary<string, Type> providers, string defaultProviderName)
    {
        _serviceProvider = serviceProvider;
        _providers = providers;
        _defaultProviderName = defaultProviderName;
    }

    public IAiProvider? GetProvider(string name)
    {
        if (!_providers.TryGetValue(name.ToLowerInvariant(), out var providerType))
            return null;

        return (IAiProvider?)_serviceProvider.GetService(providerType);
    }

    public IAiProvider GetDefaultProvider()
    {
        var provider = GetProvider(_defaultProviderName);
        if (provider == null)
            throw new InvalidOperationException(
                $"Default AI provider '{_defaultProviderName}' is not registered. " +
                $"Available providers: {string.Join(", ", _providers.Keys)}");
        return provider;
    }

    public IEnumerable<string> GetAvailableProviders() => _providers.Keys;
}
