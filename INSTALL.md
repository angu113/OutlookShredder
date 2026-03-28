# OutlookShredder — Installation Guide

Two components need to run on the workstation at all times:

| Component | Port | Purpose |
|---|---|---|
| **ShredderProxy** | 7000 (HTTP) / 7001 (HTTPS) | Calls Claude API, polls mailbox, writes to SharePoint |
| **ShredderAddinHost** | 3000 (HTTPS) | Serves the Outlook task pane and dashboard |

Choose **one** of the two install methods below.

---

## Prerequisites (both methods)

1. **.NET 8 Runtime** — download from [https://dotnet.microsoft.com/download/dotnet/8.0](https://dotnet.microsoft.com/download/dotnet/8.0)
   - Choose **Windows x64 — Runtime (not SDK)**
2. Copy the deployment package from the network drive to a local folder, e.g. `C:\ShredderDeploy\`

---

## Method A — PowerShell (requires Administrator)

Open **PowerShell as Administrator** and run:

```powershell
Set-ExecutionPolicy -Scope Process -ExecutionPolicy Bypass
C:\ShredderDeploy\deploy.ps1
```

The script will:
- Install **ShredderProxy** and **ShredderAddinHost** as Windows Services (start automatically at boot)
- Create and trust a localhost HTTPS certificate
- Register the Outlook add-in catalog in the registry

After the script completes, skip to **[Configure Secrets](#configure-secrets)** below.

---

## Method B — Manual install (no Administrator required)

### Step 1 — Copy files

Create two folders anywhere under your user profile, for example:

```
C:\Users\<you>\AppData\Local\ShredderProxy\
C:\Users\<you>\AppData\Local\ShredderAddinHost\
```

Copy the contents of `ShredderDeploy\Proxy\` into the first folder.
Copy the contents of `ShredderDeploy\AddinHost\` into the second folder.
Copy `ShredderDeploy\manifest.xml` into the AddinHost folder.

### Step 2 — Trust the HTTPS certificate

Open **Command Prompt** (no elevation needed) and run:

```cmd
dotnet dev-certs https --trust
```

If prompted by a Windows security dialog, click **Yes**. This installs a localhost certificate for the current user only — it does not require admin rights.

### Step 3 — Configure secrets

See **[Configure Secrets](#configure-secrets)** below, then return here.

### Step 4 — Auto-start both services at login

You need both applications to start when you log in. The easiest way is to add shortcuts to your Startup folder.

1. Press **Win + R**, type `shell:startup`, press Enter — this opens your personal Startup folder.
2. Right-click inside the folder → **New → Shortcut**.
3. For the Proxy shortcut:
   - Location: `C:\Users\<you>\AppData\Local\ShredderProxy\OutlookShredder.Proxy.exe`
   - Name: `ShredderProxy`
4. Repeat for the AddinHost:
   - Location: `C:\Users\<you>\AppData\Local\ShredderAddinHost\OutlookShredder.AddinHost.exe`
   - Name: `ShredderAddinHost`

Both will now start automatically each time you log in. To start them immediately without rebooting, double-click each shortcut once.

**Tip — hide the console windows:**
Right-click each shortcut → Properties → **Run: Minimized**.

### Step 5 — Register the Outlook add-in

Open **PowerShell** (no elevation needed) and run:

```powershell
Set-ExecutionPolicy -Scope Process -ExecutionPolicy Bypass
C:\Users\<you>\AppData\Local\ShredderAddinHost\install-addin.ps1
```

This writes a single registry key under `HKCU` (your user only — no admin required).

---

## Configure Secrets

Both install methods require the secrets file to be filled in before the proxy will start correctly.

1. Open `appsettings.secrets.json` inside the **Proxy** install folder.
   (If it does not exist, copy `appsettings.secrets.template.json` and rename it.)

2. Fill in all values:

```json
{
  "Anthropic": {
    "ApiKey": "sk-ant-..."
  },
  "SharePoint": {
    "TenantId":     "xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx",
    "ClientId":     "xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx",
    "ClientSecret": "your-client-secret"
  },
  "Mail": {
    "MailboxAddress": "rfq@metalsupermarkets.com"
  }
}
```

3. Save the file. The proxy reads it on startup — restart the proxy after any change.

---

## Register the Outlook Add-in (first time only)

If the install script did not do this automatically:

1. Open **PowerShell** (no elevation needed).
2. Navigate to the AddinHost install folder.
3. Run:
   ```powershell
   Set-ExecutionPolicy -Scope Process -ExecutionPolicy Bypass
   .\install-addin.ps1
   ```
4. Close and reopen **Outlook**.
5. Open any email → **Home ribbon → Get Add-ins → My Add-ins**.
6. Scroll to the **OutlookShredder.AddinHost** section and click **Add**.

---

## Verify the installation

With both services running, open a browser and go to:

- `http://localhost:7000/api/health` — should return `{"status":"ok",...}`
- `https://localhost:3000/dashboard.html` — should show the RFQ dashboard

If the proxy health check fails, check the log file `proxy.log` in the Proxy install folder for error details.

---

## Troubleshooting

| Symptom | Likely cause | Fix |
|---|---|---|
| Proxy won't start | Missing or incomplete `appsettings.secrets.json` | Fill in all required values (see Configure Secrets above) |
| HTTPS certificate error in browser | Dev cert not trusted | Re-run `dotnet dev-certs https --trust` |
| Add-in not visible in Outlook ribbon | Catalog not registered or Outlook not restarted | Re-run `install-addin.ps1`, then fully close and reopen Outlook |
| Dashboard shows "Error loading data" | Proxy not running | Start `OutlookShredder.Proxy.exe` and check `http://localhost:7000/api/health` |
| Port already in use | Another process on 7000/3000 | Change the port in `appsettings.json` under `Kestrel → Endpoints` |
