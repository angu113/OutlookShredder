# Development Build - Ship Ready!

**Date**: 2024
**Status**: [READY] Proxy built and ready for dev testing
**Build Version**: Release (net8.0)

---

## QUICK START

### Start Proxy Immediately

**Option 1: PowerShell (Recommended)**
```powershell
C:\Users\angus\Shredder\Proxy\OutlookShredder\start-dev-proxy.ps1
```

**Option 2: Batch**
```cmd
C:\Users\angus\Shredder\Proxy\OutlookShredder\start-dev-proxy.bat
```

**Option 3: Direct**
```powershell
cd "C:\Users\angus\Shredder\Proxy\OutlookShredder\OutlookShredder.Proxy\bin\Release\net8.0"
.\OutlookShredder.Proxy.exe
```

### Expected Output
```
ShredderProxy starting — logs: C:\...\Logs\proxy-2024-01-15.log
Kestrel listening on http://localhost:7000
```

---

## Build Details

### What's Built

[OK] **Outlook Shredder Proxy** (Release)
- ✓ Errors: 0
- ✓ Warnings: 0
- ✓ Size: ~50MB
- ✓ Ready to test

**Location**:
```
C:\Users\angus\Shredder\Proxy\OutlookShredder\OutlookShredder.Proxy\bin\Release\net8.0\
```

**Includes**:
- OutlookShredder.Proxy.exe (main executable)
- appsettings.json (configuration)
- All .NET 8.0 dependencies
- Logging infrastructure

### Configuration Ready

**AI Providers**:
- Default: Claude (claude-sonnet-4-6)
- Fallback: Google (gemini-1.5-pro)

**Email**:
- From: store@mithrilmetals.com
- Reply-To: REMOVED (no longer sent)

**Secrets**:
- All 7 secrets configured (user-secrets)
- SharePoint credentials ready
- Service Bus connection ready
- Claude API key ready
- Google Gemini key ready

---

## Testing Resources

### 1. Dev Test Guide
**File**: `DEV_TEST_GUIDE.md`
**Contents**:
- API testing examples
- Performance benchmarking
- Troubleshooting steps
- Log monitoring
- Complete test checklist

### 2. Quick Start Scripts
**Files**:
- `start-dev-proxy.ps1` - PowerShell (colored output)
- `start-dev-proxy.bat` - Windows batch

### 3. Build Summary
**File**: `DEV_BUILD_SUMMARY.md`
**Contents**:
- Build locations
- Deployment artifacts
- Configuration details
- Troubleshooting guide

---

## What to Test

### Priority 1: Core Functionality
- [ ] Proxy starts and listens on :7000
- [ ] Extraction API works
- [ ] Claude provider responds
- [ ] Google fallback provider works
- [ ] Logs are created

### Priority 2: New Features
- [ ] Reply-To header NOT sent on emails
- [ ] Email from store@mithrilmetals.com
- [ ] Configuration-driven provider selection
- [ ] Fallback provider activated on failures

### Priority 3: Performance
- [ ] Response time acceptable (<10s)
- [ ] Compare Claude vs Google latency
- [ ] Monitor memory usage

---

## Testing Flow

### Quick Test (5 minutes)

1. Start proxy: `.\start-dev-proxy.ps1`
2. Open PowerShell
3. Test extraction:
```powershell
$body = @{
    Content = "Please quote 1000 lbs 304 SS Round Bar 1 inch diameter"
} | ConvertTo-Json

Invoke-WebRequest -Uri "http://localhost:7000/api/extract" `
  -Method POST `
  -Headers @{"Content-Type"="application/json"} `
  -Body $body
```
4. Should get back structured extraction
5. Check logs in `Logs/` directory

### Full Test (30 minutes)

Follow all steps in `DEV_TEST_GUIDE.md`:
- Test each provider
- Test fallback
- Test email configuration
- Test performance
- Verify logs

---

## Files Ready in Git

**Commits**:
```
- fb8f877: dev quick-start scripts and test guide
- 6b21d35: build summary and deployment guide
- 92d06e9: provider configuration summary
- bb47413: configuration-driven provider selection
```

**New Documentation** (in remote):
- DEV_BUILD_SUMMARY.md
- DEV_TEST_GUIDE.md
- AI_PROVIDER_CONFIGURATION.md
- PROVIDER_CONFIGURATION_SUMMARY.md
- start-dev-proxy.bat
- start-dev-proxy.ps1

**All pushed to**: https://github.com/angu113/OutlookShredder

---

## Notes

### Proxy Service
- [OK] Built in Release mode (optimized)
- [OK] Can run as console app (for dev)
- [OK] Can register as Windows Service (for production)
- [OK] All secrets configured
- [OK] All integrations ready (SharePoint, Service Bus, Email, AI)

### WPF Shredder
- [INFO] Build separately to avoid solution conflicts
- [TODO] Build with: `dotnet build Shredder.csproj -c Release --no-dependencies`
- [TODO] Connect to localhost:7000 after proxy is running

### Next Session
1. Run dev proxy with `start-dev-proxy.ps1`
2. Follow DEV_TEST_GUIDE.md for comprehensive testing
3. Build Shredder WPF separately
4. Test end-to-end workflow
5. Verify email Reply-To behavior
6. Compare provider performance

---

## Deployment

### For Testing
```powershell
# Just run the script
.\start-dev-proxy.ps1
```

### For Production
```powershell
# Register as Windows Service
sc create ShredderProxy binPath="C:\path\to\OutlookShredder.Proxy.exe"
net start ShredderProxy
```

---

**Status**: READY FOR DEV TESTING
**Build Date**: Today
**Next**: Run start-dev-proxy.ps1 to begin testing
**Documentation**: Complete and in Git
**Configuration**: Production-ready defaults
