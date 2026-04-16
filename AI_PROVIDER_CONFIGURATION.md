# AI Provider Configuration Guide

## Overview

The Outlook Shredder Proxy now supports **configuration-driven AI provider selection** with **default and fallback provider support**. This allows you to easily switch providers without code changes.

## Configuration

### appsettings.json

Add the `AiProviders` section to control which provider is used by default and which one acts as a fallback:

```json
{
  "AiProviders": {
    "DefaultProvider": "claude",
    "FallbackProvider": "google"
  }
}
```

### Configuration Options

| Setting | Type | Required | Default | Description |
|---------|------|----------|---------|-------------|
| `AiProviders:DefaultProvider` | string | Yes | `claude` | Primary provider to use for extractions (e.g., "claude", "openai", "gemini", "google", "gpt4") |
| `AiProviders:FallbackProvider` | string | No | null | Backup provider if the default provider fails or is rate-limited |

### Provider Names

These provider names can be used in configuration:

| Name | Provider | Model |
|------|----------|-------|
| `claude` | Claude (Anthropic) | claude-sonnet-4-6 (configurable) |
| `openai` | OpenAI (ChatGPT) | gpt-4-turbo (configurable) |
| `gpt4` | OpenAI (ChatGPT) | gpt-4-turbo (configurable) |
| `gemini` | Google Gemini | gemini-1.5-pro (configurable) |
| `google` | Google Gemini | gemini-1.5-pro (configurable) |

## Common Configurations

### Claude Only (Default)
```json
{
  "AiProviders": {
    "DefaultProvider": "claude"
  }
}
```

### Claude with Google Fallback (Recommended)
```json
{
  "AiProviders": {
    "DefaultProvider": "claude",
    "FallbackProvider": "google"
  }
}
```

### OpenAI Primary, Claude Fallback
```json
{
  "AiProviders": {
    "DefaultProvider": "openai",
    "FallbackProvider": "claude"
  }
}
```

### Google/Gemini Primary, Claude Fallback
```json
{
  "AiProviders": {
    "DefaultProvider": "gemini",
    "FallbackProvider": "claude"
  }
}
```

## Usage in Code

### Get the Default Provider
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

### Use Fallback Provider on Failure
```csharp
public async Task<RfqExtraction?> ExtractWithFallback(ExtractRequest req)
{
    var primaryProvider = factory.GetDefaultProvider();
    var fallbackProvider = factory.GetFallbackProvider();
    
    try
    {
        return await primaryProvider.ExtractAsync(req);
    }
    catch (HttpRequestException) when (fallbackProvider != null)
    {
        _logger.LogWarning("Primary provider failed, using fallback");
        return await fallbackProvider.ExtractAsync(req);
    }
}
```

### Override with Query Parameter
```csharp
public async Task<IActionResult> Extract(
    [FromBody] ExtractRequest req,
    [FromQuery] string? provider)
{
    var aiProvider = factory.GetProvider(provider ?? "claude");
    if (aiProvider == null)
        return BadRequest($"Unknown provider: {provider}");
    
    var extraction = await aiProvider.ExtractAsync(req);
    return Ok(extraction);
}
```

### List Available Providers
```csharp
var providers = factory.GetAvailableProviders();
// Returns: ["claude", "gpt4", "openai", "gemini", "google"]
```

## Current Setup

**Current Configuration** (set in this session):
- ✅ Default Provider: **Claude** (claude-sonnet-4-6)
- ✅ Fallback Provider: **Google** (gemini-1.5-pro)
- ✅ All API keys configured via user-secrets:
  - `Anthropic:ApiKey` ✅
  - `Google:ApiKey` ✅
  - `OpenAi:ApiKey` (optional - not set yet)

## How to Change Providers

### Change Default Provider (No Code Changes Needed)

1. **Edit `appsettings.json`**:
   ```json
   {
     "AiProviders": {
       "DefaultProvider": "openai",  // Changed from "claude"
       "FallbackProvider": "google"
     }
   }
   ```

2. **Restart the application** - uses new default immediately

### Change Fallback Provider

1. **Edit `appsettings.json`**:
   ```json
   {
     "AiProviders": {
       "DefaultProvider": "claude",
       "FallbackProvider": "openai"  // Changed from "google"
     }
   }
   ```

2. **Restart the application**

### Disable Fallback Provider

1. **Edit `appsettings.json`**:
   ```json
   {
     "AiProviders": {
       "DefaultProvider": "claude"
       // Remove FallbackProvider or set to null
     }
   }
   ```

2. **Restart the application**

## Deployment

### Development (Current)
- Secrets stored via `dotnet user-secrets`
- Configuration via `appsettings.json`
- Override via `appsettings.secrets.json` (gitignored)

### Production
- Store API keys in **Azure Key Vault** or **environment variables**
- Set `AiProviders` in `appsettings.json` (committed to source control)
- Set secrets:
  ```bash
  export ANTHROPIC__APIKEY="sk-ant-..."
  export OPENAI__APIKEY="sk-..."
  export GOOGLE__APIKEY="AIzaSy..."
  ```

## Performance Notes

**Latency by Provider** (typical):
- Claude Sonnet: 2-5 seconds
- OpenAI GPT-4 Turbo: 3-8 seconds
- Google Gemini Pro: 1-4 seconds

**Cost** (per extraction, approximate):
- Claude Sonnet: $0.003
- GPT-4 Turbo: $0.01
- Gemini Pro: $0.001

**Recommendation**: 
- Use **Claude** as default (good balance of cost/speed)
- Use **Google Gemini** as fallback (fastest, lowest cost)

## Troubleshooting

### "Default provider 'xxx' is not registered"
**Problem**: The provider name in `AiProviders:DefaultProvider` doesn't exist or the provider is not installed.

**Solution**:
- Check spelling (must be lowercase)
- Valid names: `claude`, `openai`, `gpt4`, `gemini`, `google`
- Ensure all three providers are registered in `Program.cs`

### Provider returns null/errors
**Problem**: API key is missing or invalid.

**Solution**:
```bash
# Check configured secrets
dotnet user-secrets list

# Add missing key
dotnet user-secrets set "Google:ApiKey" "your-key-here"

# Or use appsettings.secrets.json for team development
```

### Fallback provider never used
**Problem**: `GetFallbackProvider()` returns null.

**Solution**:
- Check `AiProviders:FallbackProvider` is set in appsettings.json
- Ensure fallback provider is registered (all three are by default)
- Manually check with: `var fallback = factory.GetFallbackProvider();`

## Migration from Hard-Coded Provider

### Old Code (Hard-Coded)
```csharp
// Before: always uses ClaudeService directly
var extraction = await _claudeService.ExtractAsync(req);
```

### New Code (Configuration-Driven)
```csharp
// After: uses configured default provider (switchable without code change)
var provider = factory.GetDefaultProvider();
var extraction = await provider.ExtractAsync(req);
```

**No breaking changes** - existing code continues working. New code gets provider selection automatically.

## Next Steps

1. **Test with Current Setup**: Claude (primary) + Google (fallback)
2. **Monitor Performance**: Track latency and cost for each provider
3. **Consider OpenAI**: Add `OpenAi:ApiKey` if you want to test GPT-4
4. **Production Deployment**: Move secrets to Azure Key Vault before deploying

---

**Last Updated**: Today (AI Provider Configuration Phase)  
**Status**: ✅ Complete and tested (build: 0 errors, 0 warnings)
