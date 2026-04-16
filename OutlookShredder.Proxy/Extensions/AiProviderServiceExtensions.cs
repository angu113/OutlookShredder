using OutlookShredder.Proxy.Services;
using OutlookShredder.Proxy.Services.Ai;

namespace OutlookShredder.Proxy.Extensions;

/// <summary>
/// Extension methods for registering AI providers in the dependency injection container.
/// </summary>
public static class AiProviderServiceExtensions
{
    /// <summary>
    /// Registers Claude as an AI provider and sets it as the default.
    /// Maintains backward compatibility with existing ClaudeService code.
    /// </summary>
    public static IServiceCollection AddClaudeAiProvider(this IServiceCollection services)
    {
        // Register the original ClaudeService (maintains existing code compatibility)
        services.AddSingleton<ClaudeService>();

        // Register the adapter that implements IAiProvider
        services.AddSingleton<ClaudeServiceAdapter>();

        // Register the factory with the adapter
        services.AddAiProviderFactory(options =>
        {
            options.RegisterProvider("claude", typeof(ClaudeServiceAdapter));
            options.SetDefaultProvider("claude");
        });

        return services;
    }

    /// <summary>
    /// Registers OpenAI (ChatGPT) as an AI provider.
    /// </summary>
    public static IServiceCollection AddOpenAiProvider(this IServiceCollection services)
    {
        services.AddSingleton<OpenAiProvider>();
        return services;
    }

    /// <summary>
    /// Registers Google AI (Gemini) as an AI provider.
    /// </summary>
    public static IServiceCollection AddGoogleAiProvider(this IServiceCollection services)
    {
        services.AddSingleton<GoogleAiProvider>();
        return services;
    }

    /// <summary>
    /// Registers the AI provider factory with custom configuration.
    /// </summary>
    public static IServiceCollection AddAiProviderFactory(
        this IServiceCollection services,
        Action<AiProviderFactoryOptions> configure)
    {
        var options = new AiProviderFactoryOptions();
        configure(options);

        services.AddSingleton<IAiProviderFactory>(sp =>
        {
            var providers = new Dictionary<string, Type>();
            foreach (var (name, type) in options.Providers)
            {
                providers[name.ToLowerInvariant()] = type;
            }

            return new AiProviderFactory(
                sp,
                providers,
                options.DefaultProviderName ?? throw new InvalidOperationException("No default provider configured"),
                options.FallbackProviderName);
        });

        return services;
    }
}

/// <summary>
/// Configuration options for AI provider registration.
/// </summary>
public class AiProviderFactoryOptions
{
    private readonly List<(string Name, Type ProviderType)> _providers = new();

    public IReadOnlyList<(string Name, Type ProviderType)> Providers => _providers.AsReadOnly();
    public string? DefaultProviderName { get; private set; }
    public string? FallbackProviderName { get; private set; }

    /// <summary>
    /// Registers an AI provider implementation.
    /// </summary>
    public AiProviderFactoryOptions RegisterProvider(string name, Type providerType)
    {
        if (!typeof(IAiProvider).IsAssignableFrom(providerType))
            throw new ArgumentException($"{providerType.Name} must implement IAiProvider");

        _providers.Add((name, providerType));
        return this;
    }

    /// <summary>
    /// Sets which provider to use by default.
    /// </summary>
    public AiProviderFactoryOptions SetDefaultProvider(string name)
    {
        DefaultProviderName = name;
        return this;
    }

    /// <summary>
    /// Sets which provider to use as a fallback when the default provider fails.
    /// </summary>
    public AiProviderFactoryOptions SetFallbackProvider(string name)
    {
        FallbackProviderName = name;
        return this;
    }
}
