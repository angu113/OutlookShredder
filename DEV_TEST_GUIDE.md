# Development Testing Guide

## Quick Start

### 1. Start the Proxy Service

Choose one:

**Batch Script (Windows Command Prompt)**:
```cmd
start-dev-proxy.bat
```

**PowerShell**:
```powershell
.\start-dev-proxy.ps1
```

**Direct**:
```powershell
cd "C:\Users\angus\Shredder\Proxy\OutlookShredder\OutlookShredder.Proxy\bin\Release\net8.0"
.\OutlookShredder.Proxy.exe
```

### 2. Verify Proxy is Running

You should see output like:
```
ShredderProxy starting — logs: C:\Users\angus\Shredder\Proxy\OutlookShredder\Logs\proxy-2024-01-15.log
Kestrel listening on http://localhost:7000
```

---

## Testing Extraction API

### Test 1: Basic Extraction (Claude)

```powershell
$body = @{
    Content = "Quote for 1000 lbs 304 Stainless Steel Round Bar, 1 inch diameter, $2.50 per pound"
} | ConvertTo-Json

Invoke-WebRequest -Uri "http://localhost:7000/api/extract" `
  -Method POST `
  -Headers @{"Content-Type"="application/json"} `
  -Body $body
```

**Expected Result**: Structured extraction with supplier name, products, quantities, pricing

### Test 2: With RLI Context (Anchoring)

```powershell
$body = @{
    Content = "We can supply 304 SS Round Bar 1 inch"
    RliItems = @(
        @{ ProductName = "304 SS Round Bar 1 inch"; Mspc = "SS-RB-1" }
    )
} | ConvertTo-Json

Invoke-WebRequest -Uri "http://localhost:7000/api/extract" `
  -Method POST `
  -Headers @{"Content-Type"="application/json"} `
  -Body $body
```

### Test 3: Specify Provider

```powershell
# Use Google Gemini instead of default Claude
Invoke-WebRequest -Uri "http://localhost:7000/api/extract?provider=google" `
  -Method POST `
  -Headers @{"Content-Type"="application/json"} `
  -Body $body
```

### Test 4: Test Fallback (Kill Primary Provider)

This would require modifying configuration or API keys, but the mechanism is:
1. Default provider (Claude) is attempted
2. If it fails (rate limit, API error, etc.), fallback (Google) is used
3. Check logs to see provider switching

---

## Testing Email Features

### Test: Send RFQ Email

**API Endpoint**: `POST /api/send-rfq-email`

```powershell
$body = @{
    To = "supplier@example.com"
    Subject = "RFQ Request - Stainless Steel"
    Body = "We need a quote for 500 lbs of 304 SS"
} | ConvertTo-Json

Invoke-WebRequest -Uri "http://localhost:7000/api/send-rfq-email" `
  -Method POST `
  -Headers @{"Content-Type"="application/json"} `
  -Body $body
```

**Verify**:
1. Email is sent from `store@mithrilmetals.com`
2. **No Reply-To header** (main requirement)
3. Check supplier's reply goes to `store@mithrilmetals.com`

---

## Viewing Logs

### Real-Time Log Monitoring

```powershell
# Watch logs as they appear (like 'tail -f' in Unix)
Get-Content "C:\Users\angus\Shredder\Proxy\OutlookShredder\Logs\proxy-*.log" -Tail 50 -Wait
```

### View Last 100 Lines

```powershell
Get-Content "C:\Users\angus\Shredder\Proxy\OutlookShredder\Logs\proxy-*.log" -Tail 100
```

### Search for Specific Message

```powershell
# Find all "ERROR" messages
Select-String "ERROR" "C:\Users\angus\Shredder\Proxy\OutlookShredder\Logs\proxy-*.log"

# Find provider-related messages
Select-String "provider\|Provider" "C:\Users\angus\Shredder\Proxy\OutlookShredder\Logs\proxy-*.log"
```

---

## Testing Configuration Changes

### Test: Switch Default Provider

1. Edit `appsettings.json`:
```json
{
  "AiProviders": {
    "DefaultProvider": "google",  // Change from "claude" to "google"
    "FallbackProvider": "claude"  // Change from "google" to "claude"
  }
}
```

2. **Stop the proxy** (Ctrl+C)

3. **Start the proxy** again

4. Test extraction - should now use Google first

### Test: Disable Fallback

1. Edit `appsettings.json`:
```json
{
  "AiProviders": {
    "DefaultProvider": "claude"
    // Remove or comment out FallbackProvider
  }
}
```

2. Restart proxy

3. Check logs - should show no fallback available

---

## Performance Testing

### Measure Response Time

```powershell
$timer = [System.Diagnostics.Stopwatch]::StartNew()

$response = Invoke-WebRequest -Uri "http://localhost:7000/api/extract" `
  -Method POST `
  -Headers @{"Content-Type"="application/json"} `
  -Body $body

$timer.Stop()

Write-Host "Response time: $($timer.ElapsedMilliseconds)ms"
Write-Host "Response: $($response.Content)"
```

### Provider Performance Comparison

Test each provider and compare:

```powershell
# Claude
Measure-Command {
    Invoke-WebRequest -Uri "http://localhost:7000/api/extract?provider=claude" `
      -Method POST `
      -Headers @{"Content-Type"="application/json"} `
      -Body $body
} | Select-Object TotalMilliseconds

# Google
Measure-Command {
    Invoke-WebRequest -Uri "http://localhost:7000/api/extract?provider=google" `
      -Method POST `
      -Headers @{"Content-Type"="application/json"} `
      -Body $body
} | Select-Object TotalMilliseconds

# OpenAI (if configured)
Measure-Command {
    Invoke-WebRequest -Uri "http://localhost:7000/api/extract?provider=openai" `
      -Method POST `
      -Headers @{"Content-Type"="application/json"} `
      -Body $body
} | Select-Object TotalMilliseconds
```

---

## Troubleshooting Tests

### Test 1: Port 7000 Already in Use

```powershell
netstat -ano | findstr :7000
```

**Solution**: Stop the process using port 7000, or change port in appsettings.json:
```json
{
  "Kestrel": {
    "Endpoints": {
      "Http": { "Url": "http://localhost:7001" }  // Changed from 7000
    }
  }
}
```

### Test 2: API Returns 401 Unauthorized

**Cause**: API key not configured

```powershell
# Check secrets
dotnet user-secrets list

# Add missing key
dotnet user-secrets set "Anthropic:ApiKey" "your-key-here"
```

### Test 3: Provider Not Found

```powershell
Invoke-WebRequest "http://localhost:7000/api/extract?provider=invalid" ...
```

**Response**: 400 Bad Request - "Unknown provider: invalid"

**Valid providers**: claude, openai, gpt4, gemini, google

### Test 4: Timeout (>30 seconds)

**Cause**: Provider is slow or blocked

**Check logs** for:
- Rate limiting (429 errors)
- API errors
- Network issues

**Solution**:
- Increase `TimeoutSeconds` in appsettings.json
- Check API quota/limits
- Verify internet connection

---

## Expected Log Output

### Successful Extraction

```
[14:23:45 INF] ShredderProxy starting — logs: C:\...\Logs\proxy-2024-01-15.log
[14:23:45 INF] Kestrel listening on http://localhost:7000
[14:23:50 INF] Extraction request received from API
[14:23:50 INF] Using provider: claude
[14:23:52 INF] Extraction completed successfully
[14:23:52 INF] Response sent: RfqExtraction with 3 products
```

### With Fallback

```
[14:23:50 INF] Using provider: claude
[14:23:55 WRN] Claude API error: Rate limited (429)
[14:23:55 INF] Switching to fallback provider: google
[14:23:57 INF] Fallback extraction succeeded
```

### Provider Not Available

```
[14:23:50 INF] Provider requested: openai
[14:23:50 ERR] Provider not registered: openai
[14:23:50 WRN] Fallback provider: google
```

---

## Checklist: Complete Dev Test

- [ ] Proxy starts without errors
- [ ] Can call extraction API with Claude
- [ ] Can specify provider with query parameter
- [ ] Fallback works (if primary fails)
- [ ] Email sends with no Reply-To header
- [ ] Logs are created and readable
- [ ] Performance is acceptable (<10s per request)
- [ ] Configuration changes work without rebuild
- [ ] Error messages are helpful
- [ ] All 3 providers work (if keys configured)

---

## Next Steps After Testing

1. **Verify Email**: Check that replies go to store@mithrilmetals.com (not old address)
2. **Test Fallback**: Intentionally fail primary provider, verify fallback works
3. **Monitor Costs**: Track API usage and costs by provider
4. **Performance**: Compare latency and quality across providers
5. **Plan Production**: Set up Azure Key Vault for secrets

---

**Date**: Development Build
**Status**: Ready for testing
**Build Version**: Release (net8.0)
