# OutlookShredder

Office.js Web Add-in for Outlook that detects [XXXXXX] job references in incoming emails and extracts structured RFQ supplier quote data into:

**SharePoint:** `https://metalsupermarkets.sharepoint.com/sites/hackensack/Lists/RFQLineItems`

---

## Solution structure

| Project | Type | Purpose |
|---|---|---|
| `OutlookShredder.AddinHost` | ASP.NET Core Empty | Serves add-in HTML/JS/CSS over HTTPS via IIS Express (port 3000) |
| `OutlookShredder.Proxy` | ASP.NET Core Web API | Holds Anthropic API key; calls Claude; writes to SharePoint (port 5001) |

## First-time setup

### 1. Configure secrets

Right-click `OutlookShredder.Proxy` > **Manage User Secrets**, then add:

```json
{
  "Anthropic": { "ApiKey": "sk-ant-..." },
  "SharePoint": {
    "TenantId":     "your-tenant-id",
    "ClientId":     "your-client-id",
    "ClientSecret": "your-client-secret"
  }
}
```

### 2. Set multiple startup projects

Right-click the solution > **Set Startup Projects** > select **Multiple startup projects**, set both to **Start**.

### 3. Provision SharePoint columns (run once)

Start the solution (F5), then call:
```
POST https://localhost:5001/api/setup-columns
```
Use the **HTTP files** in the Proxy project or a browser/Postman.

### 4. Sideload the manifest in Outlook

Outlook desktop → File → Manage Add-ins → Add custom add-in → From file → `OutlookShredder.AddinHost/manifest.xml`

---

## How it works

1. Email arrives → Outlook shows the **Process RFQ** button in the ribbon
2. Add-in scans subject + body + attachment names for `[XXXXXX]` pattern
3. If matched → calls `POST https://localhost:5001/api/extract` with email content
4. C# proxy calls Claude API, extracts JSON, writes one SharePoint row per product
5. Task pane shows extracted products; Dashboard at `/dashboard.html` shows all records

## SharePoint columns captured per product row

SupplierName · DateOfQuote · EstimatedDeliveryDate · ProductName ·
UnitsRequested · UnitsQuoted · LengthPerUnit · LengthUnit ·
WeightPerUnit · WeightUnit · PricePerPound · PricePerFoot · SupplierProductComments
