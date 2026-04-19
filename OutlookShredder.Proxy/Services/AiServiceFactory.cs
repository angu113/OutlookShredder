namespace OutlookShredder.Proxy.Services;

/// <summary>
/// Resolves the AI extraction service based on <c>AI:Provider</c> in configuration.
///
/// <para>Supported values (case-insensitive):</para>
/// <list type="bullet">
///   <item><c>claude</c> (default) — Claude primary; Gemini fallback if its API key is configured.</item>
///   <item><c>gemini</c> — Gemini primary; Claude fallback if its API key is configured.</item>
///   <item><c>roundrobin</c> / <c>round-robin</c> — alternate Claude and Gemini per call with cross-fallback on failure.
///         Requires both API keys; degrades to single-provider mode if only one is configured.</item>
/// </list>
///
/// <para>The resolved service is cached for the lifetime of this factory (which is a singleton),
/// so any per-instance state — notably <see cref="RoundRobinAiExtractionService"/>'s alternation
/// counter — persists across extraction calls.</para>
/// </summary>
public class AiServiceFactory
{
    private readonly IServiceProvider _serviceProvider;
    private readonly IConfiguration _config;
    private readonly ILogger<AiServiceFactory> _log;
    private readonly Lazy<IAiExtractionService> _cached;

    public AiServiceFactory(
        IServiceProvider serviceProvider,
        IConfiguration config,
        ILogger<AiServiceFactory> log)
    {
        _serviceProvider = serviceProvider;
        _config = config;
        _log = log;
        _cached = new Lazy<IAiExtractionService>(Build, LazyThreadSafetyMode.ExecutionAndPublication);
    }

    public IAiExtractionService GetService() => _cached.Value;

    private IAiExtractionService Build()
    {
        var providerName = (_config["AI:Provider"] ?? "claude").ToLowerInvariant();

        if (providerName is "roundrobin" or "round-robin")
        {
            return BuildRoundRobin();
        }

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

    private IAiExtractionService BuildRoundRobin()
    {
        var claude = ResolveByName("claude");
        var gemini = ResolveByName("gemini");
        var claudeOk = IsProviderKeyConfigured(claude);
        var geminiOk = IsProviderKeyConfigured(gemini);

        if (!claudeOk && !geminiOk)
        {
            _log.LogWarning("AI provider: roundrobin requested but neither Anthropic:ApiKey nor Google:ApiKey is configured — returning Claude (will fail on first call)");
            return claude;
        }

        if (!geminiOk)
        {
            _log.LogWarning("AI provider: roundrobin requested but Google:ApiKey missing — falling back to Claude alone");
            return claude;
        }

        if (!claudeOk)
        {
            _log.LogWarning("AI provider: roundrobin requested but Anthropic:ApiKey missing — falling back to Gemini alone");
            return gemini;
        }

        _log.LogInformation("AI provider: round-robin between {A} and {B} with cross-fallback on failure",
            claude.ProviderName, gemini.ProviderName);

        var rrLog = _serviceProvider.GetRequiredService<ILogger<RoundRobinAiExtractionService>>();
        return new RoundRobinAiExtractionService(claude, gemini, rrLog);
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
