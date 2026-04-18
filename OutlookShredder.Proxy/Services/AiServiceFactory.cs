namespace OutlookShredder.Proxy.Services;

/// <summary>
/// Factory for creating AI extraction service instances based on configuration.
/// Supports multiple providers with fallback logic.
/// </summary>
public class AiServiceFactory
{
    private readonly IServiceProvider _serviceProvider;
    private readonly IConfiguration _config;
    private readonly ILogger<AiServiceFactory> _log;

    public AiServiceFactory(
        IServiceProvider serviceProvider,
        IConfiguration config,
        ILogger<AiServiceFactory> log)
    {
        _serviceProvider = serviceProvider;
        _config = config;
        _log = log;
    }

    /// <summary>
    /// Gets the configured AI extraction service, or the first available if config is invalid.
    /// Provider is selected via AI:Provider in appsettings.
    /// Falls back to the default provider if not specified.
    /// </summary>
    public IAiExtractionService GetService()
    {
        var provider = _config["AI:Provider"]?.ToLowerInvariant() ?? "claude";

        _log.LogInformation("AI provider configured: {Provider}", provider);

        return provider switch
        {
            "gemini" => GetService<GeminiExtractionService>(),
            "claude" => GetService<ClaudeExtractionService>(),
            // Future: "openai" => GetService<OpenAiExtractionService>(),
            _ => GetService<ClaudeExtractionService>() // default fallback
        };
    }

    private IAiExtractionService GetService<T>() where T : IAiExtractionService
    {
        var service = _serviceProvider.GetRequiredService<T>();
        _log.LogInformation("Using AI provider: {Provider}", service.ProviderName);
        return service;
    }
}
