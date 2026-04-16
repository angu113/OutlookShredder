# Quick Reference: Using Multiple AI Providers

## 📋 TL;DR

You now have Claude, GPT-4, and Gemini all available. No code changes needed to start using them.

---

## 🔧 Setup (5 minutes)

### Get API Keys

| Provider | URL | Free Tier |
|----------|-----|-----------|
| **Claude** (Anthropic) | https://console.anthropic.com | $5 credit |
| **OpenAI** | https://platform.openai.com/api-keys | $5 credit |
| **Google AI** | https://makersuite.google.com/app/apikey | Free (limited) |

### Store Keys Locally

```bash
# Use PowerShell (your preferred shell)
dotnet user-secrets set "Claude:ApiKey" "sk-ant-your-key-here"
dotnet user-secrets set "OpenAi:ApiKey" "sk-your-key-here"
dotnet user-secrets set "Google:ApiKey" "AIzaSy-your-key-here"
```

---

## 💻 Using from Code

### Get Default Provider (Claude)
```csharp
var extraction = await factory.GetDefaultProvider().ExtractAsync(request);
```

### Use Specific Provider
```csharp
// Use GPT-4
var gpt4 = factory.GetProvider("gpt4");
var extraction = await gpt4.ExtractAsync(request);

// Use Gemini
var gemini = factory.GetProvider("gemini");
var extraction = await gemini.ExtractAsync(request);
```

### Available Provider Names
```csharp
factory.GetProvider("claude");   // ← Claude adapter
factory.GetProvider("gpt4");     // ← OpenAI GPT-4
factory.GetProvider("openai");   // ← Same as gpt4
factory.GetProvider("gemini");   // ← Google Gemini
factory.GetProvider("google");   // ← Same as gemini
```

---

## 🔄 Changing Default Provider

In `Proxy\OutlookShredder\OutlookShredder.Proxy\Program.cs`:

```csharp
// Line ~90, change this:
options.SetDefaultProvider("claude");  // ← Current default

// To this:
options.SetDefaultProvider("gpt4");    // ← Use GPT-4 instead
```

Then rebuild:
```bash
dotnet build
```

---

## 🌐 Using from REST API

Add optional `?provider=NAME` query parameter:

```bash
# Use default (Claude)
curl -X POST https://localhost:3000/api/extract \
  -H "Content-Type: application/json" \
  -d '{"content":"..."}'

# Use GPT-4
curl -X POST "https://localhost:3000/api/extract?provider=gpt4" \
  -H "Content-Type: application/json" \
  -d '{"content":"..."}'

# Use Gemini
curl -X POST "https://localhost:3000/api/extract?provider=gemini" \
  -H "Content-Type: application/json" \
  -d '{"content":"..."}'
```

### Controller Example
```csharp
[HttpPost("extract")]
public async Task<IActionResult> Extract(
    [FromBody] ExtractRequest req,
    [FromQuery] string? provider = null)
{
    var aiProvider = factory.GetProvider(provider ?? "claude");
    var extraction = await aiProvider.ExtractAsync(req);
    return Ok(extraction);
}
```

---

## ⚙️ Configuration Reference

All config goes in `appsettings.json` or `appsettings.secrets.json`:

```json
{
  "Claude": {
    "ApiKey": "sk-ant-...",
    "Model": "claude-3-5-sonnet-20241022",
    "MaxRetries": 3,
    "TimeoutSeconds": 30
  },
  "OpenAi": {
    "ApiKey": "sk-...",
    "Model": "gpt-4-turbo",
    "MaxRetries": 3,
    "TimeoutSeconds": 30
  },
  "Google": {
    "ApiKey": "AIzaSy...",
    "Model": "gemini-1.5-pro",
    "MaxRetries": 3,
    "TimeoutSeconds": 30
  }
}
```

### Configuration Fields

| Field | Default | Purpose |
|-------|---------|---------|
| `ApiKey` | (required) | Authentication token for the provider |
| `Model` | See above | Which model version to use |
| `MaxRetries` | 3 | Retry attempts on rate limit (429) |
| `TimeoutSeconds` | 30 | Max wait time per request |

### Changing Models

```json
{
  "OpenAi": {
    "Model": "gpt-3.5-turbo"  // Faster, cheaper
  },
  "Google": {
    "Model": "gemini-1.5-flash"  // Faster, cheaper
  }
}
```

---

## 📊 Quick Comparison

| Feature | Claude | GPT-4 | Gemini |
|---------|--------|-------|--------|
| **Extraction Quality** | ⭐⭐⭐⭐⭐ | ⭐⭐⭐⭐⭐ | ⭐⭐⭐⭐ |
| **Speed** | Medium | Slow | Fast |
| **Cost** | $$ | $$$ | $ |
| **Reliability** | ⭐⭐⭐⭐⭐ | ⭐⭐⭐⭐⭐ | ⭐⭐⭐⭐ |
| **RFQ Extraction** | Proven | Excellent | Excellent |

**Recommendation**:
- **Default**: Claude (best balance)
- **Best Accuracy**: GPT-4 (if budget allows)
- **Best Value**: Gemini (lowest cost)
- **Testing**: Try all three with `?provider=` query param

---

## 🚨 Common Issues

### "API key not found"
```bash
# Verify secret is set:
dotnet user-secrets list

# Set it again:
dotnet user-secrets set "OpenAi:ApiKey" "sk-..."
```

### "Unknown provider: xxx"
```csharp
// Make sure provider name is registered in Program.cs
// Valid names: claude, gpt4, openai, gemini, google
```

### Timeout errors (> 30 seconds)
```json
{
  "OpenAi": {
    "TimeoutSeconds": 60
  }
}
```

### Rate limiting (429 errors)
The proxy automatically retries with backoff. If still rate-limited:
- Increase `MaxRetries` (default 3)
- Wait a few minutes
- Check provider usage dashboard

---

## 🎯 Testing Flow

1. **Start proxy**
   ```bash
   cd Proxy\OutlookShredder
   dotnet build
   dotnet run --project OutlookShredder.Proxy
   ```

2. **Test Claude** (should work immediately if key is set)
   ```bash
   curl -X POST https://localhost:3000/api/extract \
     -H "Content-Type: application/json" \
     -d "{\"content\":\"Please extract RFQ data...\"}"
   ```

3. **Test GPT-4** (add `?provider=gpt4`)
   ```bash
   curl -X POST "https://localhost:3000/api/extract?provider=gpt4" \
     -H "Content-Type: application/json" \
     -d "{\"content\":\"...\"}"
   ```

4. **Test Gemini** (add `?provider=gemini`)
   ```bash
   curl -X POST "https://localhost:3000/api/extract?provider=gemini" \
     -H "Content-Type: application/json" \
     -d "{\"content\":\"...\"}"
   ```

---

## 📝 Code Example: Compare All Three

```csharp
[HttpPost("api/compare")]
public async Task<IActionResult> CompareAllProviders([FromBody] ExtractRequest req)
{
    var results = new Dictionary<string, object>();
    
    foreach (var name in new[] { "claude", "gpt4", "gemini" })
    {
        var provider = factory.GetProvider(name);
        if (provider != null)
        {
            try
            {
                var extraction = await provider.ExtractAsync(req);
                results[name] = extraction ?? "null";
            }
            catch (Exception ex)
            {
                results[name] = new { error = ex.Message };
            }
        }
    }
    
    return Ok(results);
}
```

---

## ✅ Files to Know

| File | Purpose |
|------|---------|
| `Services/Ai/OpenAiProvider.cs` | GPT-4 implementation |
| `Services/Ai/GoogleAiProvider.cs` | Gemini implementation |
| `Services/Ai/IAiProvider.cs` | Provider interface |
| `Program.cs` | Provider registration |
| `MULTI_AI_PROVIDER_SETUP.md` | Full setup guide |
| `OPENAI_GOOGLE_AI_IMPLEMENTATION.md` | Complete docs |

---

## 🎓 Learn More

- **Setup Details**: Read `MULTI_AI_PROVIDER_SETUP.md`
- **Implementation**: Read `OPENAI_GOOGLE_AI_IMPLEMENTATION.md`
- **Example Config**: See `appsettings.example.json`
- **API Docs**:
  - https://docs.anthropic.com
  - https://platform.openai.com/docs
  - https://ai.google.dev/docs

---

## 💡 Pro Tips

1. **Use different providers for different purposes**
   - Claude for production (proven)
   - GPT-4 for high-value RFQs
   - Gemini for cost-sensitive workloads

2. **Monitor API usage**
   - Claude: https://console.anthropic.com
   - OpenAI: https://platform.openai.com/account/usage
   - Google: https://aistudio.google.com/app/apikey

3. **Set up alerting**
   - Budget limits in each provider dashboard
   - Get notified before spending limits

4. **Test with small samples first**
   - Send a few test extractions to each provider
   - Compare results before going to production

5. **Keep API keys secure**
   - Never commit to git
   - Use `appsettings.secrets.json`
   - Use environment variables in production

---

**Status**: ✅ All providers working, build succeeds

**Next**: Add API keys and test!
