# 🚀 QUICK START: Switch to Gemini NOW

## ⚡ Immediate Action Required

Your Anthropic API is down. Here's how to switch to Google Gemini in **3 steps**:

---

## Step 1: Get Gemini API Key (2 minutes)

1. Open: https://aistudio.google.com/app/apikey
2. Sign in with Google account
3. Click **"Create API Key"**
4. Copy the key (starts with `AIza...`)

---

## Step 2: Update Configuration (1 minute)

Edit: `C:\Users\angus\source\repos\angu113\OutlookShredder\OutlookShredder.Proxy\appsettings.secrets.json`

Add these lines:

```json
{
  "AI": {
    "Provider": "gemini"
  },
  "Google": {
    "ApiKey": "AIza_YOUR_KEY_HERE"
  }
}
```

**If the file doesn't exist**, create it with:

```json
{
  "AI": {
    "Provider": "gemini"
  },
  "Google": {
    "ApiKey": "AIza_YOUR_KEY_HERE"
  },
  "Anthropic": {
    "ApiKey": "your_old_claude_key_if_you_have_it"
  }
}
```

---

## Step 3: Restart Proxy (30 seconds)

```powershell
# Stop current proxy
Stop-Process -Name "OutlookShredder.Proxy" -ErrorAction SilentlyContinue

# Run the proxy (or just launch Shredder - it auto-starts the proxy)
cd C:\Users\angus\source\repos\angu113\OutlookShredder\OutlookShredder.Proxy\bin\Debug\net8.0
.\OutlookShredder.Proxy.exe
```

**OR** just close and reopen **Shredder** - it will auto-start the proxy with the new config.

---

## ✅ Verify It Worked

Check the proxy console/logs for:

```
[INFO] AI provider configured: gemini
[INFO] Using AI provider: Gemini
```

---

## ⚠️ Important Note

The Gemini implementation is a **placeholder** right now. To make it fully work, we need to:

1. **Test** that it throws the NotImplementedException (expected)
2. **Implement** the actual Gemini API calls using the proper SDK

**For now**, this gets the architecture in place. When the Anthropic API comes back, you can switch back by changing:

```json
{
  "AI": {
    "Provider": "claude"
  }
}
```

---

## 🆘 Need Help?

**If Gemini doesn't work yet:**
1. It's expected - the implementation is incomplete
2. Wait for Anthropic API to come back, OR
3. I can complete the Gemini implementation now (will take ~30 minutes)

**Want me to finish Gemini implementation?** Just say:
> "Complete the Gemini implementation"

And I'll:
- Research the correct Mscc.GenerativeAI API usage
- Implement proper `ExtractRfqAsync` logic
- Test with your API key
- Handle PDF extraction for POs

---

**Next Steps**: See `AI_PLUGGABLE_IMPLEMENTATION.md` for full details.
