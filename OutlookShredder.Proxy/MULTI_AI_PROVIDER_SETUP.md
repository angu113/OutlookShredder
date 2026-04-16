# Multi-AI Provider Configuration Guide

This guide explains how to configure and use OpenAI and Google AI providers alongside the existing Claude provider.

## Configuration

Add the following to your `appsettings.json` or `appsettings.secrets.json`:

### Claude (Anthropic)
```json
{
  "Claude": {
    "ApiKey": "sk-ant-...",
    "Model": "claude-3-5-sonnet-20241022",
    "MaxRetries": 3,
    "TimeoutSeconds": 30
  }
}
```

### OpenAI (ChatGPT)
```json
{
  "OpenAi": {
    "ApiKey": "sk-...",
    "Model": "gpt-4-turbo",
    "MaxRetries": 3,
    "TimeoutSeconds": 30
  }
}
```

Supported models:
- `gpt-4-turbo` - Latest GPT-4 Turbo model (recommended)
- `gpt-4` - Stable GPT-4
- `gpt-4o` - GPT-4 Optimized
- `gpt-3.5-turbo` - Faster, cheaper alternative

Get your API key from: https://platform.openai.com/api-keys

### Google AI (Gemini)
```json
{
  "Google": {
    "ApiKey": "AIzaSy...",
    "Model": "gemini-1.5-pro",
    "MaxRetries": 3,
    "TimeoutSeconds": 30
  }
}
```

Supported models:
- `gemini-1.5-pro` - Latest Gemini Pro (recommended)
- `gemini-1.5-flash` - Faster, cheaper alternative
- `gemini-1.0-pro` - Stable Gemini 1.0

Get your API key from: https://makersuite.google.com/app/apikey

## Using Providers in Code

### Getting the Default Provider
The default provider is Claude. To use it:

```csharp
public class ExtractController(IAiProviderFactory factory)
{
    [HttpPost("extract")]
    public async Task<IActionResult> Extract([FromBody] ExtractRequest req)
    {
        var provider = factory.GetDefaultProvider();
        var extraction = await provider.ExtractAsync(req);
        return Ok(extraction);
    }
}
```

### Switching to a Specific Provider
To use a different provider:

```csharp
public class ExtractController(IAiProviderFactory factory)
{
    [HttpPost("extract")]
    public async Task<IActionResult> Extract([FromBody] ExtractRequest req, [FromQuery] string? provider)
    {
        var aiProvider = factory.GetProvider(provider ?? "claude");
        if (aiProvider == null)
            return BadRequest($"Unknown provider: {provider}");
        
        var extraction = await aiProvider.ExtractAsync(req);
        return Ok(extraction);
    }
}
```

### Listing Available Providers
```csharp
var providers = factory.GetAvailableProviders();
// Returns: ["claude", "openai", "gpt4", "gemini", "google"]
```

## Changing the Default Provider

In `Program.cs`, modify the provider registration:

```csharp
builder.Services.AddClaudeAiProvider();
builder.Services.AddOpenAiProvider();
builder.Services.AddGoogleAiProvider();

builder.Services.AddAiProviderFactory(options =>
{
    options.RegisterProvider("claude", typeof(ClaudeServiceAdapter));
    options.RegisterProvider("gpt4", typeof(OpenAiProvider));
    options.RegisterProvider("openai", typeof(OpenAiProvider));
    options.RegisterProvider("gemini", typeof(GoogleAiProvider));
    options.RegisterProvider("google", typeof(GoogleAiProvider));
    options.SetDefaultProvider("gpt4");  // ← Change default here
});
```

## Running with Limited Providers

To reduce startup time or dependencies, register only the providers you need:

```csharp
// Only Claude
builder.Services.AddClaudeAiProvider();
builder.Services.AddAiProviderFactory(options =>
{
    options.RegisterProvider("claude", typeof(ClaudeServiceAdapter));
    options.SetDefaultProvider("claude");
});
```

```csharp
// Only OpenAI
builder.Services.AddOpenAiProvider();
builder.Services.AddAiProviderFactory(options =>
{
    options.RegisterProvider("gpt4", typeof(OpenAiProvider));
    options.RegisterProvider("openai", typeof(OpenAiProvider));
    options.SetDefaultProvider("gpt4");
});
```

## Provider Capabilities

All providers implement the same `IAiProvider` interface:

```csharp
public interface IAiProvider
{
    string Name { get; }
    Task<RfqExtraction?> ExtractAsync(ExtractRequest request, CancellationToken cancellationToken = default);
}
```

### Extraction Features
- **Input**: Emails, PDFs, DOCX, text attachments
- **Output**: Structured `RfqExtraction` with:
  - Supplier name
  - Product descriptions with dimensions, grades, forms
  - Quote references
  - Pricing (per-lb, per-ft, per-piece, total)
  - Quantities and units
  - Comments and notes

### Anchoring (RLI Context)
All providers support optional RFQ Line Item (RLI) context for anchoring products:

```csharp
var request = new ExtractRequest
{
    Content = "...",
    RliItems = new List<RliContextItem>
    {
        new() { ProductName = "304 SS Round Bar 1\"", Mspc = "SS-RB-1" },
        new() { ProductName = "316L SS Flat Bar 1/2\" x 2\"", Mspc = "SS-FB-1-2" }
    }
};

var extraction = await provider.ExtractAsync(request);
```

## Performance Notes

**Latency** (typical):
- Claude: 2-5 seconds
- OpenAI: 3-8 seconds
- Google Gemini: 1-4 seconds

**Cost** (per extraction, approximate):
- Claude Sonnet 3.5: $0.003
- GPT-4 Turbo: $0.01
- Gemini Pro: $0.001

**Throughput**:
- All providers: ~100 concurrent requests
- With retries: configured exponential backoff (1s → 30s max)

## Troubleshooting

### "Unknown provider: xxx"
```csharp
// Ensure provider is registered in Program.cs
builder.Services.AddAiProviderFactory(options =>
{
    options.RegisterProvider("xxx", typeof(MyProvider));
});
```

### API Key not found
```csharp
// Ensure key is in appsettings.json or appsettings.secrets.json
// For development, use:
// dotnet user-secrets set "OpenAi:ApiKey" "sk-..."
```

### Timeout errors
Increase `TimeoutSeconds` in config for slow networks:
```json
{
  "OpenAi": {
    "ApiKey": "...",
    "TimeoutSeconds": 60  // Default is 30
  }
}
```

### Rate limiting (429 errors)
Increase `MaxRetries` to handle backoff better:
```json
{
  "OpenAi": {
    "ApiKey": "...",
    "MaxRetries": 5  // Default is 3
  }
}
```

## Adding a New Provider

1. **Create the provider class** implementing `IAiProvider`:
```csharp
public class MyAiProvider : IAiProvider
{
    public string Name => "myai";
    
    public async Task<RfqExtraction?> ExtractAsync(
        ExtractRequest request, 
        CancellationToken cancellationToken = default)
    {
        // Your implementation
    }
}
```

2. **Register in Program.cs**:
```csharp
builder.Services.AddSingleton<MyAiProvider>();
builder.Services.AddAiProviderFactory(options =>
{
    // ... existing providers ...
    options.RegisterProvider("myai", typeof(MyAiProvider));
});
```

3. **Use in code**:
```csharp
var provider = factory.GetProvider("myai");
```

## API Key Security

### Development
Store keys in user secrets:
```bash
dotnet user-secrets set "Claude:ApiKey" "sk-ant-..."
dotnet user-secrets set "OpenAi:ApiKey" "sk-..."
dotnet user-secrets set "Google:ApiKey" "AIzaSy..."
```

### Production
- Store in Azure Key Vault
- Or environment variables:
```bash
set "OPENAI_API_KEY=sk-..."
```

### Never
- Commit keys to git
- Log keys in error messages
- Expose keys in API responses

## Testing Multiple Providers

Create an endpoint to compare extraction results:

```csharp
[HttpPost("test/compare")]
public async Task<IActionResult> CompareProviders([FromBody] ExtractRequest req)
{
    var providers = new[] { "claude", "gpt4", "gemini" };
    var results = new Dictionary<string, RfqExtraction?>();
    
    foreach (var name in providers)
    {
        var provider = factory.GetProvider(name);
        results[name] = await provider?.ExtractAsync(req);
    }
    
    return Ok(results);
}
```

## Migration from Single Provider

If you're currently using only Claude:

1. Add new providers to `appsettings.json`
2. Register them in `Program.cs`
3. Keep `SetDefaultProvider("claude")` to maintain existing behavior
4. Existing code continues working without changes
5. Gradually migrate to new providers as needed

No breaking changes! Old code targeting `ClaudeService` directly still works.
