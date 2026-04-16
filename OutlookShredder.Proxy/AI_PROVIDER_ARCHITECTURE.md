# AI Provider Architecture

This document explains how to add support for new AI models (GPT-4, Llama, etc.) to the OutlookShredder proxy.

## Overview

The proxy now uses a pluggable AI provider architecture. The extraction logic is no longer tightly coupled to Claude—any AI model can be plugged in by implementing the `IAiProvider` interface.

## Adding a New AI Provider

### 1. Create a Provider Implementation

Create a new file implementing `IAiProvider`:

```csharp
using OutlookShredder.Proxy.Models;
using OutlookShredder.Proxy.Services.Ai;

namespace OutlookShredder.Proxy.Services.Ai;

public class OpenAiProvider : IAiProvider
{
    private readonly IConfiguration _config;
    private readonly ILogger<OpenAiProvider> _log;
    private readonly HttpClient _http;

    public string Name => "gpt4";

    private const string ApiUrl = "https://api.openai.com/v1/chat/completions";

    public OpenAiProvider(IConfiguration config, ILogger<OpenAiProvider> log)
    {
        _config = config;
        _log = log;
        var timeoutSeconds = int.TryParse(_config["OpenAi:TimeoutSeconds"], out var t) ? t : 60;
        _http = new HttpClient { Timeout = TimeSpan.FromSeconds(timeoutSeconds) };
    }

    public async Task<RfqExtraction?> ExtractAsync(ExtractRequest request, CancellationToken cancellationToken = default)
    {
        // TODO: Implement extraction using OpenAI API
        // See ClaudeAiProvider for reference implementation
        
        var apiKey = _config["OpenAi:ApiKey"]
            ?? throw new InvalidOperationException("OpenAi:ApiKey is not configured.");

        // Build request, call API, parse response
        // Return RfqExtraction or null on failure
        
        throw new NotImplementedException();
    }
}
```

### 2. Register the Provider

Update `Program.cs` to register your provider:

```csharp
// Add the provider to DI
builder.Services.AddSingleton<OpenAiProvider>();

// Register it with the factory
builder.Services.AddAiProviderFactory(options =>
{
    options.RegisterProvider("claude", typeof(ClaudeAiProvider));
    options.RegisterProvider("gpt4", typeof(OpenAiProvider));
    options.SetDefaultProvider("claude");  // Set which one is used by default
});

// Keep backward compatibility
builder.Services.AddSingleton<ClaudeService>();
```

Or, if you want to make it the default:

```csharp
options.SetDefaultProvider("gpt4");
```

### 3. Add Configuration

Add API keys and settings to `appsettings.json` or `appsettings.secrets.json`:

```json
{
  "OpenAi": {
    "ApiKey": "sk-...",
    "Model": "gpt-4-turbo",
    "TimeoutSeconds": 60,
    "MaxTokens": 4096,
    "MaxRetries": 3
  }
}
```

### 4. Using the Provider in Controllers

Controllers can use the provider in two ways:

#### Option A: Use the Factory (Flexible)
```csharp
public class ExtractController : ControllerBase
{
    private readonly IAiProviderFactory _aiFactory;

    public ExtractController(IAiProviderFactory aiFactory)
    {
        _aiFactory = aiFactory;
    }

    [HttpPost("extract")]
    public async Task<ActionResult> Extract([FromBody] ExtractRequest req)
    {
        // Get a specific provider
        var provider = _aiFactory.GetProvider("gpt4");
        
        // Or use the default
        var defaultProvider = _aiFactory.GetDefaultProvider();
        
        var extraction = await defaultProvider.ExtractAsync(req);
        // ...
    }
}
```

#### Option B: Use ClaudeService (Backward Compatible)
Existing code using `ClaudeService` continues to work without changes:

```csharp
public class ExtractController : ControllerBase
{
    private readonly ClaudeService _claude;

    public ExtractController(ClaudeService claude)
    {
        _claude = claude;
    }

    [HttpPost("extract")]
    public async Task<ActionResult> Extract([FromBody] ExtractRequest req)
    {
        var extraction = await _claude.ExtractAsync(req);
        // ...
    }
}
```

The `ClaudeService` wrapper automatically delegates to the default provider, so existing code works unchanged.

## Provider Responsibilities

Each provider must:

1. **Extract RFQ Data**: Parse supplier quotes from emails/attachments
2. **Return Consistent Schema**: Always return `RfqExtraction` with the standard `ProductLine[]`
3. **Handle Media Types**: Support:
   - PDF attachments (base64-encoded)
   - DOCX attachments (base64-encoded)
   - Plain text attachments
   - Email body text
4. **Handle RLI Context**: Support the "Requested items for this RFQ" section for anchor matching
5. **Implement Retry Logic**: Handle rate limits and transient failures
6. **Log Operations**: Use `ILogger<T>` for debugging and monitoring
7. **Support Cancellation**: Respect `CancellationToken` for timeouts

## System Prompt Strategy

Each provider receives the same high-level instructions encoded in the `ExtractRequest`:

- **Job references** (pre-identified patterns like `[XXXXXX]`)
- **RLI items** (requested products with MSPC codes)
- **Email context** (body snippet for matching)
- **Content** (email body or attachment text)

Providers are expected to translate these into their own API format (e.g., Claude uses tool-use, GPT-4 might use structured output, etc.).

### Claude's Approach (Reference)
- Uses **tool-use** for guaranteed JSON structure
- Sends a detailed system prompt explaining the extraction task
- Uses **prompt caching** for efficiency

### Recommended for Other Providers
- **GPT-4/4-turbo**: Use structured output mode (JSON schema mode)
- **Llama (local)**: Use function calling or constrain with grammar
- **Generic LLM API**: Use JSON in-context examples + few-shot prompting

## Testing a New Provider

1. **Unit tests**: Mock `IConfiguration` and `ILogger`
2. **Integration tests**: Call a real API endpoint with test data
3. **Dry-run endpoint**: Add a `GET /api/test-extract?provider=gpt4` to test without writing to SharePoint

Example dry-run controller:

```csharp
[HttpGet("test-extract")]
public async Task<ActionResult> TestExtract(
    [FromQuery] string provider = "claude",
    [FromBody] ExtractRequest req)
{
    var aiProvider = _aiFactory.GetProvider(provider);
    if (aiProvider == null)
        return BadRequest($"Provider '{provider}' not found. Available: {string.Join(", ", _aiFactory.GetAvailableProviders())}");

    var result = await aiProvider.ExtractAsync(req);
    return Ok(new { success = result != null, extraction = result });
}
```

## Performance Considerations

- **Prompt Caching** (Claude): Reduces cost and latency for repeated system prompts
- **Token Limits**: Different models have different limits; adjust `MaxTokens` per model
- **Retry Delays**: Some models are more prone to rate limits; adjust `MaxRetries` and delays
- **Concurrency**: Use connection pooling (`HttpClient` singleton per provider)

## Migration Path

### Current State
- `ClaudeService` directly calls Anthropic API
- Controllers depend on `ClaudeService`

### After This Refactoring
- `ClaudeService` becomes a wrapper around `IAiProvider`
- New code can target `IAiProvider` or `IAiProviderFactory`
- Old code targeting `ClaudeService` continues to work unchanged

### Adding a Second Provider
1. Create `OpenAiProvider : IAiProvider`
2. Register it in `Program.cs`
3. Existing `ClaudeService` clients continue working
4. New code can choose: `_aiFactory.GetProvider("gpt4")`

## Example: Full Flow with Multiple Providers

```csharp
// In Program.cs
builder.Services.AddSingleton<ClaudeAiProvider>();
builder.Services.AddSingleton<OpenAiProvider>();
builder.Services.AddAiProviderFactory(options =>
{
    options.RegisterProvider("claude", typeof(ClaudeAiProvider));
    options.RegisterProvider("gpt4", typeof(OpenAiProvider));
    options.SetDefaultProvider("claude");
});
builder.Services.AddSingleton<ClaudeService>();

// In a controller or service
public class RfqProcessingService
{
    private readonly IAiProviderFactory _aiFactory;

    public RfqProcessingService(IAiProviderFactory aiFactory)
    {
        _aiFactory = aiFactory;
    }

    public async Task ProcessWithBothModels(ExtractRequest req)
    {
        var claude = _aiFactory.GetProvider("claude");
        var gpt4 = _aiFactory.GetProvider("gpt4");

        var claudeResult = await claude.ExtractAsync(req);
        var gpt4Result = await gpt4.ExtractAsync(req);

        // Compare results, use consensus, or store both for analysis
    }
}
```

## Troubleshooting

### "No default provider is configured"
- Check `Program.cs`—did you call `SetDefaultProvider()`?
- Verify the provider type is registered with `AddSingleton<>`

### Provider not found
- Call `_aiFactory.GetAvailableProviders()` to list registered providers
- Check provider name matches registration (case-insensitive)

### Configuration values not reading
- Verify `appsettings.json` structure (nested under provider name, e.g., `OpenAi:ApiKey`)
- Check `appsettings.secrets.json` is deployed and readable
- Use `ILogger` to confirm configuration values are loaded

## References

- `IAiProvider` interface: `Services/Ai/IAiProvider.cs`
- `ClaudeAiProvider` implementation: `Services/Ai/ClaudeAiProvider.cs`
- Factory registration: `Extensions/AiProviderServiceExtensions.cs`
- Service registration: `Program.cs` (search "AI Provider Registration")
