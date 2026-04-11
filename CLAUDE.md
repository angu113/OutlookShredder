# OutlookShredder — Developer Hints

Two projects in one repo:
- **OutlookShredder.Proxy** — ASP.NET Core 8 Windows Service; all business logic, Graph API, Claude API
- **OutlookShredder.AddinHost** — Static file host for the Office.js Outlook taskpane add-in

## Proxy (`OutlookShredder.Proxy/`)

### Bootstrap (`Program.cs`)
Registers all DI, runs as Windows Service (`ShredderProxy`) or console. Key config sections: `SharePoint`, `Mail`, `Anthropic`, `ServiceBus`, `Suppliers`, `Proxy`.

### Controllers

| Controller | Prefix | Purpose |
|-----------|--------|---------|
| `ExtractController` | `/api` | Office.js calls `POST /api/extract` to extract one email/attachment via Claude |
| `QcController` | `/api/qc` | Read QC list, check last-modified, trigger LQ update, patch single row |
| `ServiceBusController` | `/api/service-bus` | Serves Service Bus connection string + topic to Shredder clients |
| `CatalogController` | `/api/catalog` | Read/refresh product catalog, backfill, fuzzy-resolve vendor names |
| `RfqNewController` | `/api/rfq-new` | Send RFQ emails via Graph; served product/supplier data for RfqNew tab |
| `UpdateController` | `/api/update` | Version check + download publish package ZIP |

**ExtractController endpoints:**
- `POST /api/extract` — body: `ExtractRequest` → `ExtractResponse`; calls Claude, writes SharePoint, stamps "RFQ-Processed" on message, publishes notification
- `POST /api/setup-supplier-lists` — idempotent: creates SupplierResponses + SupplierLineItems SP lists

**QcController endpoints:**
- `GET /api/qc` → `{ columns[], rows[][], itemIds[], lastModified }`
- `GET /api/qc/last-modified` → `{ lastModified }`
- `POST /api/qc/update-lq` → `{ updated[], misses[] }`
- `PATCH /api/qc/update-row` body: `{ itemId, qc, qcCut }` → patches QC and QC Cut columns on a single SharePoint item

**ServiceBusController endpoints:**
- `GET /api/service-bus/config` → `{ configured: bool, connectionString: string|null, topicName: string }` — Shredder fetches this on startup instead of maintaining its own copy of the connection string

**CatalogController endpoints:**
- `GET /api/catalog` → catalog items + cache status
- `POST /api/catalog/refresh` — force cache reload
- `POST /api/catalog/backfill` — patch `CatalogProductName` on all SupplierLineItems rows
- `GET /api/catalog/resolve?name=` → fuzzy match result

**RfqNewController endpoints:**
- `GET /api/rfq-new/catalog` → `ProductCatalogItem[]` (`Mspc`, `Name`, `Category`, `Shape`)
- `GET /api/rfq-new/categories` → string[]
- `GET /api/rfq-new/shapes` → string[]
- `GET /api/rfq-new/supplier-relationships` → `(SupplierName, Email, Metal/Category, Shape)[]`
- `GET /api/rfq-new/existing-ids` → string[]
- `POST /api/rfq-new/create` — creates RFQ Reference + SupplierLineItems rows
- `POST /api/rfq-new/send-email` body: `{ subject, body, bccAddresses[] }`

**ShredderConfig endpoints (on ExtractController or separate):**
- `GET /api/shredder-config/{name}` → `{ value }`
- `PUT /api/shredder-config/{name}` body: `{ value }`

**RfqImport endpoints (on ExtractController or separate):**
- `GET /api/rfq-import/scan?mailbox=&folder=` → `RfqEmailCandidate[]`
- `GET /api/rfq-import/existing-ids` → string[]
- `POST /api/rfq-import/import` body: `RfqEmailCandidate`
- `GET /api/rfq-import/processed?top=N` → `ProcessedEmailItem[]`
- `POST /api/rfq-import/reprocess` body: `{ messageIds[] }`
- `POST /api/rfq-import/dedupe-supplier-responses?dryRun=true` — dedup SR-level and within-SR SLI duplicates using name normalisation + Jaccard ≥ 0.5
- `DELETE /api/rfq-import/clean` — wipe SupplierResponses + SupplierLineItems

**Supplier data endpoints:**
- `GET /api/items?top=N&raw=true` — all SLI items merged with SR fields; `raw=true` skips the in-memory Jaccard dedup (exposes all SP rows with `SpItemId` populated — useful for admin cleanup)
- `GET /api/items/by-rfq/{rfqId}` — SLI items for one RFQ (always includes `SpItemId`)
- `DELETE /api/sr/{srId}` — delete all SLIs under an SR then the SR itself
- `DELETE /api/sli/{itemId}` — delete a single SLI by its SP item ID

### Services

**`ClaudeService`** (singleton)
- `ExtractAsync(ExtractRequest)` → `RfqExtraction`
- Uses **tool_use** (`extract_rfq` tool) for schema-enforced output — no free-text JSON parsing
- Static system prompt + tool definition sent with **prompt caching** (`anthropic-beta: prompt-caching-2024-07-31`)
- Retries 429/5xx/network errors up to `Claude:MaxRetries` with randomised jitter
- Logs warning when `stop_reason == "max_tokens"` or content is truncated at `Claude:MaxContentChars`
- Extraction fields: jobReference, quoteReference, supplierName, freightTerms, products[]
  (dateOfQuote/estimatedDeliveryDate removed — dates come from the RFQ Reference record, not extraction)

**`SharePointService`** (singleton)
- All Graph API calls. Uses `ClientSecretCredential` (app-only, `Sites.FullControl.All`).
- `WriteProductRowAsync(extraction, productLine, request, source, sourceFile, index)` → `SpWriteResult`
  - Deduplicates by email+product; prefers attachment source over body
  - OOF detection; resolves supplier via `SupplierCacheService`
- `ReadQcListAsync()` → `{ columns, rows, itemIds, lastModified }` — itemIds are SharePoint item IDs, parallel-indexed with rows[]
- `GetQcLastModifiedAsync()` → `DateTime?`
- `UpdateQcLqAsync()` → `(updated count, misses list)` — derives $/lb from quote rows, updates QC 'LQ' column
- `UpdateQcRowAsync(itemId, qc, qcCut)` — patches QC and QC Cut fields on a single SP item; resolves internal column names automatically
- `GetPublishVersionAsync()` → version string
- `EnsureSupplierListsAsync()` — idempotent list creation (provisions SupplierResponses, SupplierLineItems, PurchaseOrders)
- `WritePurchaseOrderAsync(rfqId, supplierName, poNumber, receivedAt, messageId, lineItemsJson)` → deduped by MessageId
- `ReadPurchaseOrdersAsync()` → `List<PurchaseOrderRecord>` — all PO rows
- `GET /api/purchase-orders` controller endpoint — Shredder loads this on startup

**`MailService`** (singleton)
- Graph API for mailbox (app-only, `Mail.ReadWrite` + `Mail.Send`)
- `SendRfqEmailAsync(subject, body, bccAddresses)` — sends via Graph
- `GetMessageByIdAsync(mailbox, messageId)` → message metadata + body
- `MarkProcessedAsync(mailbox, itemId, extra)` — stamps "RFQ-Processed" category
- Strips RE:/FW: prefixes, [EXTERNAL] tags, converts HTML → plain text
- Extracts job references via regex `RFQ\s+\[([A-Za-z0-9]+)\]`

**`MailPollerService`** (hosted service — background)
- Polls inbox every `Mail:PollIntervalSeconds` (default 30s) for messages without "RFQ-Processed"
- Per message: strips FW:/RE:/[EXTERNAL] prefixes, then routes:
  - Subject starts with `"Purchase Order #HSK-PO"` → `ProcessPurchaseOrderAsync` (extracts PDF via Claude, writes to `PurchaseOrders` SP list, publishes `EventType="PO"` to Service Bus, stamps "RFQ-Processed"+"PO-Processed")
  - Everything else → normal RFQ pipeline (Claude extract → SharePoint → notify)
- `ReprocessMessagesAsync(messageIds[])` — manual re-extraction (routes POs correctly too)
- Config: `Mail:MailboxAddress`, `Mail:LookbackHours` (default 24), `Mail:MaxEmailsPerMinute`, `Mail:BodyContextChars`

**`RfqNotificationService`** (singleton pub/sub)
- `Subscribe()` → `(Guid, ChannelReader<string>)` — SSE subscriber registration
- `NotifyRfqProcessed(RfqProcessedNotification)` — broadcasts to:
  - SSE: `"rfq-processed\n{json}"` to all connected clients
  - Azure Service Bus topic (`ServiceBus:TopicName`, default `rfq-updates`)

**`SupplierCacheService`** (hosted service)
- In-memory known-supplier list from SharePoint (`Suppliers:SourcingList`)
- Refreshes every `Suppliers:RefreshIntervalMinutes`

**`ProductCatalogService`** (hosted service)
- In-memory product catalog from SharePoint Catalog list
- `RefreshAsync()`, `ResolveProduct(vendorName)` → `(Name, SearchKey)?`
- `CachedNames`, `LastRefreshAt`, `LastError`, `LastDiag`

**`ProductDeduplicator`**
- Fuzzy-match vendor product descriptions against catalog names

**`LqUpdateService`** (hosted service)
- Periodically calls `SharePointService.UpdateQcLqAsync()`

### Key Models (`Models/`)

**`ExtractRequest`** — from Office.js add-in or MailPollerService
```
Content, SourceType ("body"|"attachment"), FileName, Base64Data, ContentType,
JobRefs[], BodyContext, EmailBody, EmailSubject, EmailFrom, ReceivedAt,
HasAttachment, EwsToken, EwsUrl, ItemId, AttachId
```

**`ExtractResponse`**
```
Success, Extracted: RfqExtraction, Rows: SpWriteResult[]
```

**`RfqExtraction`** — Claude output
```
JobReference, QuoteReference, SupplierName, DateOfQuote, EstimatedDeliveryDate,
FreightTerms, Products: ProductLine[]
```

**`ProductLine`**
```
ProductName, UnitsRequested, UnitsQuoted, LengthPerUnit, LengthUnit,
WeightPerUnit, WeightUnit, PricePerPound, PricePerFoot, PricePerPiece,
TotalPrice, LeadTimeText, Certifications, SupplierProductComments
```

**`RfqProcessedNotification`** / **`RfqBusMessage`** (Service Bus wire format)
```
EventType ("SR"|"RFQ"), RfqId, SupplierName, MessageId, Products[]: { Name, TotalPrice }
```
`MessageId` = Graph message ID of the source email. Shredder uses it as the dedup key so two distinct emails from the same supplier each trigger their own toast (while SSE + Service Bus delivering the same event are still collapsed to one).

### SharePoint Lists

| List | Key columns |
|------|-------------|
| `SupplierResponses` | EmailFrom, Subject, ReceivedAt, SourceType, FileName, SupplierName, IsOutOfOffice, IsSupplierUnknown |
| `SupplierLineItems` | JobReference, QuoteReference, ProductName, SupplierProductName, CatalogProductName, CatalogProductSearchKey, Mspc, pricing fields (PricePerPound/Foot/Piece, TotalPrice), IsRegret |
| `RFQ References` | RfqId, Requester, DateSent, EmailRecipients, ProductLineCount, Notes |
| `QC` | Dynamic columns — Metal, Shape, LQ (Last Quote $/lb), and product name columns |
| `Catalog` | Mspc, Name, SearchKey, Category, Shape |
| `SourcingList` | Supplier names + emails (source for SupplierCacheService) |

### Configuration

Secrets go in `appsettings.secrets.json` (gitignored) — copy from `appsettings.secrets.template.json` and fill in values. Can also be set via environment variables.

**Required secrets** (app will not work without these):

| Key | Purpose |
|-----|---------|
| `SharePoint:TenantId` | Azure AD tenant ID |
| `SharePoint:ClientId` | Azure AD app registration client ID |
| `SharePoint:ClientSecret` | Azure AD app registration client secret (`Sites.FullControl.All`, `Mail.ReadWrite`, `Mail.Send`) |
| `Anthropic:ApiKey` | Claude API key (from console.anthropic.com) |
| `Mail:MailboxAddress` | UPN of the mailbox to monitor (e.g. `store@mithrilmetals.com`) |
| `ServiceBus:ConnectionString` | Azure Service Bus namespace connection string (send+listen policy on the `rfq-updates` topic) |

**Optional / tuning** (have defaults in `appsettings.json`):

| Key | Default | Purpose |
|-----|---------|---------|
| `ServiceBus:TopicName` | `rfq-updates` | Topic name |
| `Claude:Model` | `claude-sonnet-4-6` | Model ID |
| `Claude:MaxTokens` | `4096` | Max output tokens; raise if truncation warnings appear in logs |
| `Claude:MaxRetries` | `3` | Retry count on 429/5xx/network errors |
| `Claude:TimeoutSeconds` | `60` | HTTP timeout per call |
| `Claude:MaxContentChars` | `12000` | Text truncation limit sent to Claude |
| `Mail:FromAddress` | `store@mithrilmetals.com` | Sender address on RFQ emails |
| `Mail:ReplyToAddress` | `hackensack@metalsupermarkets.com` | Reply-To on RFQ emails |
| `Mail:PollIntervalSeconds` | `30` | How often the poller checks for new mail |
| `Mail:LookbackHours` | `24` | Rolling window of messages considered per poll |
| `Proxy:AllowedOrigin` | `https://localhost:3000` | CORS origin for AddinHost |

### Logging
Serilog — console + rolling daily file at `Logs/proxy-.log`.

---

## AddinHost (`OutlookShredder.AddinHost/`)

Minimal ASP.NET Core 8 static file server (HTTPS on port 3000). Serves the Office.js taskpane.

CORS allows: `https://localhost`, `https://outlook.office.com`, `https://outlook.office365.com`, `https://*.office365.com`, `https://*.microsoft.com`.

### wwwroot Files

| File | Purpose |
|------|---------|
| `manifest.xml` | Office add-in manifest (declares taskpane for Outlook) |
| `taskpane.html` | Add-in UI shell |
| `taskpane.js` | Main add-in logic; calls `POST /api/extract` on proxy |
| `attachmentReader.js` | Extracts attachment as base64 via Office.js API; falls back to EWS SOAP |

**User workflow in Outlook:**
1. Select email → open Shredder taskpane
2. Review body / select attachment
3. Click Extract → `POST /api/extract` → Claude → SharePoint → notification
4. Shredder desktop receives Service Bus event → refreshes RFQ grid

---

## Cross-Repo Communication Summary

```
Shredder desktop  ──HTTP──►  Proxy (/api/*)
Office.js addin   ──HTTP──►  Proxy (/api/extract)
Proxy             ──HTTPS──► Anthropic API (claude-3-* models)
Proxy             ──HTTPS──► Microsoft Graph (mail + SharePoint)
Proxy             ──AMQP───► Azure Service Bus (topic: rfq-updates)
Shredder desktop  ◄─AMQP───  Azure Service Bus (RfqServiceBusListener)
```

## Rules

- `@mithrilmetals.com` is never a valid supplier — never appear in extraction results or email targets
- All SharePoint writes go through `SharePointService` — no direct Graph calls from controllers
- Claude extraction prompt and JSON schema live in `ClaudeService.ExtractAsync` — edit there to change extraction behaviour
