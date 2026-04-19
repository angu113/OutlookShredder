namespace OutlookShredder.Proxy.Services;

/// <summary>
/// Resolves the AI extraction service based on <c>AI:Provider</c> in configuration.
/// When the other provider's API key is configured, wraps the chosen primary in a
/// <see cref="FallbackAiExtractionService"/> so a runtime failure of the primary
/// transparently retries against the secondary.
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

    public IAiExtractionService GetService()
    {
        var providerName = (_config["AI:Provider"] ?? "claude").ToLowerInvariant();
        var primary = ResolveByName(providerName);
        var secondary = ResolveByName(providerName == "gemini" ? "claude" : "gemini");

        if (!IsProviderKeyConfigured(secondary))
        {
            _log.LogInformation("AI provider: {Primary} (no fallback — {Secondary} API key not configured)",
                primary.ProviderName, secondary.ProviderName);
            return primary;
        }

        _log.LogInformation("AI provider: {Primary} with fallback to {Secondary}",
            primary.ProviderName, secondary.ProviderName);

        var fallbackLog = _serviceProvider.GetRequiredService<ILogger<FallbackAiExtractionService>>();
        return new FallbackAiExtractionService(primary, secondary, fallbackLog);
    }

    private IAiExtractionService ResolveByName(string name) => name switch
    {
        "gemini" => _serviceProvider.GetRequiredService<GeminiExtractionService>(),
        // Future: "openai" => _serviceProvider.GetRequiredService<OpenAiExtractionService>(),
        _        => _serviceProvider.GetRequiredService<ClaudeExtractionService>(),
    };

    private bool IsProviderKeyConfigured(IAiExtractionService svc) => svc switch
    {
        ClaudeExtractionService => !string.IsNullOrWhiteSpace(_config["Anthropic:ApiKey"]),
        GeminiExtractionService => !string.IsNullOrWhiteSpace(_config["Google:ApiKey"]),
        _ => false,
    };
}
