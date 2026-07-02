# OutlookShredder — Developer Hints

## Repo Coordination

This proxy is accessed as a **git submodule** inside the Shredder repo at `Shredder/Proxy/OutlookShredder/`. There is no separate primary clone — edit in-place inside Shredder, then commit + push from the submodule working directory, and bump the submodule pointer in the parent Shredder commit.

**Breaking API changes must land with the matching Shredder-side change in the same commit that bumps the submodule pointer.** A pointer bump without the companion client change (or vice versa) will fail at runtime.

The submodule clone has no local git identity by default; pass `-c user.name=Angus -c user.email=angus@mithrilmetals.com` on `git commit` or set it locally once via `git config`.

Two projects in one repo:
- **OutlookShredder.Proxy** — ASP.NET Core 10 Windows Service; all business logic, Graph API, Claude API
- **OutlookShredder.AddinHost** — Static file host for the Office.js Outlook taskpane add-in

## Proxy (`OutlookShredder.Proxy/`)

### Bootstrap (`Program.cs`)
Registers all DI, runs as Windows Service (`ShredderProxy`) or console. Key config sections: `SharePoint`, `Mail`, `Anthropic`, `ServiceBus`, `Suppliers`, `Proxy`.

Also installs:
- **Per-request HTTP log middleware** for `/api/*` — logs `METHOD PATH -> STATUS in Nms`; slow requests (≥1000ms) or 5xx responses are logged at Warning level to make perf regressions visible in the proxy log without drowning it in info noise.
- **SharePoint pre-warm** fired from `ApplicationStarted` — calls `SharePointService.PrewarmAsync` which resolves site ID, common list IDs (SupplierResponses, SupplierConversations), and issues a throwaway Graph query to cache the OAuth token + establish the HTTP/2 connection. Cuts first-user-request latency by ~500ms. Non-fatal — logged at Warning on failure.

### Controllers

> **This table is a curated subset, not exhaustive.** The proxy has ~44 controllers; the routes below are
> the ones you'll touch most. The source of truth for any route is the `[Route(...)]` + `[Http*]` attributes
> on the controller class — `grep -rn "\[Http" Controllers/` to enumerate. Newer surfaces (Pulse/CX SMS,
> Workflow, Forge, ShadowCat/Reconciliation, MailEval/MailRules/MailClassify, CutOptimizer, Statements,
> Steve/OpenBravo relay, Archive, Drawing, Todos, Cache, etc.) each have their own wip/implementation doc.

| Controller | Prefix | Purpose |
|-----------|--------|---------|
| `ExtractController` | `/api` | Office.js calls `POST /api/extract` to extract one email/attachment via Claude |
| `QcController` | `/api/qc` | Read QC list, check last-modified, trigger LQ update, patch single row |
| `ServiceBusController` | `/api/service-bus` | Serves Service Bus connection string + topic to Shredder clients |
| `CatalogController` | `/api/catalog` | Read/refresh product catalog, backfill, fuzzy-resolve vendor names |
| `RfqNewController` | `/api/rfq-new` | Send RFQ emails via Graph; served product/supplier data for RfqNew tab |
| `UpdateController` | `/api/update` | Version check + download publish package ZIP |
| `HealthController` | `/api/health` | Aggregated service health for Shredder's Home dashboard |
| `MailStatusController` | `/api/mail` | Live snapshot of poller, reprocess batch, rate limiter, and in-flight messages |
| `SupplierConversationsController` | `/api` | Read supplier conversation threads + send follow-up inquiries (WIP) |
| `RfqSummaryController` | `/api/rfq` | `POST /api/rfq/summarize` — turns a client-assembled RFQ text input into ≤3 AI bullet points (Claude); empty bullets on failure so the client falls back to its deterministic summary |
| `CustomersController` | `/api/customers` | CRM — lookup by phone, import partners/contacts/customer-info (enrichment), list contacts, per-customer payment `terms` |
| `ImportController` | `/api/import` | Drop-and-run CSV import for BP/contact bulk loads (see Import directory below) |
| `ErpController` | `/api/erp` | Proxy SP PDFs (with FAB-drawing append), build a slip's combined DXF, ERP document records + stamp annotations |
| `DiagController` | `/api/diag` | Read-only extraction-pipeline traces (live email vs extracted/stored) + `GET /api/diag/sp-contract` (SP data-contract round-trip self-check) |
| `InquiriesController` | `/api/inquiries` | **Pulse / CX SMS customer inquiries** (Phases 0–7) — list/detail, send SMS + outbound MMS, notes, quotations, read-state, AI draft accept/dismiss/regenerate, identity/media backfill |
| `MessagingController` | `/api/messages` | Generic messaging gateway — conversation summaries/threads, `POST /api/messages/send` (sms/email/internal), mark-read, known users (predates Pulse; powers the older Messages surface) |
| `SmsWebhookController` | `/api/sms` | SignalWire inbound webhook + delivery-status callback (HMAC-validated, enqueues to `sms-inbound-jobs`) + dev-seed. Auth-exempt + loopback-bound. **Public ingress is now the `OutlookShredder.SmsWebhook` Azure Function** (see below); this controller is the in-proxy receiver/status sink. |

**ExtractController endpoints:**
- `POST /api/extract` — body: `ExtractRequest` → `ExtractResponse`; calls Claude, writes SharePoint, stamps "RFQ-Processed" on message, publishes notification
- `POST /api/setup-supplier-lists` — idempotent: creates SupplierResponses + SupplierLineItems SP lists
- `PATCH /api/sr/{srId}/rfq-id` body: `{ rfqId }` — reparents a SupplierResponse and all its child SupplierLineItems to a new RFQ ID; rfqId must be 6 chars: 2 letters + 4 alphanumeric (e.g. `AW0001`)
- `PATCH /api/sli/{sliItemId}/rfq-id` body: `{ rfqId }` — reparents a single SupplierLineItem to a new RFQ ID; if the source SR has no remaining SLI children it is deleted
- `POST /api/sr/{srId}/detach-file` body: `{ fileName }` — removes a wrongly-attached quote PDF (job-ref-mismatch cleanup): deletes the drive file at `QuoteAttachments/{srId}/{fileName}` and clears the stale `SourceFile` pointer on the SR (resetting `ProcessingSource` to `body`) and all its child SLI rows. Line-item data is left intact. Returns `{ fileDeleted, sliCleared }` (`fileDeleted=false` is normal when the PDF was filed under its own RFQ and only a stale text pointer leaked onto this SR).
- `GET /api/version` → `{ version }` — returns the running assembly's `InformationalVersion` (distinct from `/api/publish/version`, which reads SharePoint's version.txt)

**ErpController endpoints (`/api/erp`):**
- `GET /api/erp/pdf?url=&appendFabs=` — downloads an SP PDF with app-only creds (clients can't fetch SP WebUrls directly). `appendFabs=true` runs `PickingSlipFabAppender.AppendFabDrawings` — one rendered drawing page per **deduped** `FAB:` note (dedup is by part slug, so OpenBravo's double-printed echoes collapse to one). Append is per-request, not persisted.
- `GET /api/erp/fab-dxf?url=` → `{ ok, partCount, parts[], dxfBase64 }` — downloads the slip, develops its **deduped** `FAB:` notes (`PickingSlipFabAppender.GetFabDescs`), and lays every flat pattern into ONE DXF via `FabDxfBuilder` (parts left-to-right **1" apart**, bottom-aligned, shared cut/bend layers). `ok=false` (200) when the slip has no developable FAB notes; 502 on download/build error. The Shredder client (`ErpView.EnsureCadDrawingStampAsync`) calls this when no `.ncex`/`.dxf` for the HSK# exists in the OneDrive CAD folder, saves the bytes as `{HSK#}.dxf` (overwrite — OneDrive versions older copies), stamps `DIBUJO: {file}`, and shows a header notice. Un-developable notes are skipped (logged), never fatal.
- `GET /api/erp/documents[/{spItemId}]`, `POST /api/erp/setup`, `PATCH /api/erp/documents/{id}/annotations`, `DELETE /api/erp/clean-by-type?types=` — ERP document records + stamp-annotation persistence (full-list replace).

**HealthController endpoints:**
- `GET /api/health` → `{ services: [ { id, label, status, detail } ] }` — lightweight snapshot of SharePoint (catalog cache state), Mailbox (config), Service Bus (config), Claude (key configured), Gemini (key configured), and AI Routing (current `IAiExtractionService.ProviderName`). `status` is one of `ok` | `degraded` | `fail` | `disabled`. Reads cached state only — does not hit Graph live.

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

**MailStatusController endpoints:**
- `GET /api/mail/status` → `MailStatus` — live snapshot: poller cycle state, reprocess batch progress, rate limiter, and list of in-flight messages

**SupplierConversationsController endpoints (WIP — supplier follow-up feature):**
- `GET /api/supplier-conversations?rfqId=&supplierName=&outboundOnly=false` → `{ rfqId, supplierName, messages: ConversationMessage[] }` — merged thread (inbound from `SupplierResponses` + outbound from `SupplierConversations`); set `outboundOnly=true` to skip the SR scan when the caller already has inbound data from the RFQ grid
- `POST /api/supplier-inquiry/send` body: `SupplierInquiryRequest { to, subject, body, rfqId, supplierName, supplierResponseId?, inReplyTo?, attachmentName?, attachmentContentBase64?, attachmentContentType? }` → `{ success, spItemId }` — sends via `MailService.SendSupplierInquiryAsync`, appends outbound row to `SupplierConversations` list. Rejects `@mithrilmetals.com` recipients.
- SP list `SupplierConversations` is provisioned by `setup-supplier-lists` with columns `RFQ_ID, SupplierName, SupplierResponseId, Direction, MessageId, InReplyTo, SentAt, EmailSubject, BodyText, HasAttachments, ExtractedPricing`. `RFQ_ID + SupplierName` are indexed. Also indexes `SupplierResponses`/`SupplierLineItems`/`PurchaseOrders` hot columns so thread queries don't trigger unindexed-scan warnings.

**Mail maintenance endpoints:**
- `POST /api/mail/processed-emails?top=N` — alias for rfq-import/processed; returns `ProcessedEmailItem[]`
- `POST /api/mail/reprocess-selected` body: `{ messageIds[] }` — manually re-extracts specific messages by Graph message ID (routes POs correctly too); blocks until batch completes
- `POST /api/mail/backfill-message-ids?days=7` — scans SR rows within the window that are missing MessageId; matches them to Graph messages by sender+time (±5 min), patches MessageId on matched rows and their child SLIs
- `POST /api/mail/deduplicate?days=7` — deletes SR rows with no MessageId AND collapses duplicate rows sharing the same MessageId (keeps highest-scoring: attachment source > priced SLI > has QuoteReference); also deletes child SLI rows for deleted SRs
- Run order: `setup-supplier-lists` first (creates the column), then `backfill-message-ids`, then `deduplicate`

**MessageId maintenance endpoints (legacy — prefer Mail maintenance endpoints above):**
- `POST /api/mail/backfill-message-ids?days=7` — scans SR rows within the window that are missing MessageId; matches them to Graph messages by sender+time (±5 min), patches MessageId on matched rows and their child SLIs
- `POST /api/mail/deduplicate?days=7` — deletes SR rows with no MessageId AND collapses duplicate rows sharing the same MessageId (keeps highest-scoring: attachment source > priced SLI > has QuoteReference); also deletes child SLI rows for deleted SRs
- Run order: `setup-supplier-lists` first (creates the column), then `backfill-message-ids`, then `deduplicate`

**Supplier data endpoints:**
- `GET /api/items?top=N&raw=true` — all SLI items merged with SR fields; `raw=true` skips the in-memory Jaccard dedup (exposes all SP rows with `SpItemId` populated — useful for admin cleanup)
- `GET /api/items/by-rfq/{rfqId}` — SLI items for one RFQ (always includes `SpItemId`)
- `DELETE /api/sr/{srId}` — delete all SLIs under an SR then the SR itself
- `DELETE /api/sli/{itemId}` — delete a single SLI by its SP item ID

**InquiriesController endpoints (`/api/inquiries`) — Pulse / CX SMS inquiries.** Thin over `InquiryService`; all
threading / opt-out / draft / notification logic lives in the service. Inquiry ids are `CINQ-…`.
- `GET /api/inquiries?status=&q=` — list inquiries (status tab + text filter)
- `GET /api/inquiries/unread-total` — total unread inbound across active inquiries (the taskbar + Pulse-icon badge)
- `GET /api/inquiries/find?phone=` — does a thread already exist for this number? → `{ found, inquiryId }` (one-thread-per-customer)
- `POST /api/inquiries/start` body `{ phone }` — get-or-create-or-reopen the one thread for a number (operator-initiated); 400 on invalid US number
- `GET /api/inquiries/{id}` — detail aggregate (inquiry + messages + notes + quotations + drafts + CRM card)
- `GET /api/inquiries/{id}/media?name=` — stream a stored inbound media file (proxy holds bytes durably via app-only Graph; client never needs carrier auth)
- `POST /api/inquiries/{id}/backfill-media` body `{ sid, mediaJson }` — dev/recovery re-pull of media for a message by SID (gated by `SignalWire:AllowDevSeed`)
- `POST /api/inquiries/backfill-identity` body `{ from, to, apply }` — admin one-time: rewrite an operator identity (Windows login → Shredder username) across `Inquiries.AssignedTo` / `InquiryNotes.NoteAuthor` / `InquiryQuotations.LinkedBy`. Dry-run unless `apply=true`; returns per-list matched/patched + affected inquiry ids
- `POST /api/inquiries/{id}/messages` body `{ body, from, fromDraftSpItemId? }` — send an SMS reply (opt-out-aware → 409; auto-assigns owner)
- `POST /api/inquiries/{id}/messages/mms` (multipart/form-data: `body`, `from`, `fromDraftSpItemId?`, `files[]`) — outbound MMS; images send as MMS, a PDF is rasterized to one image per page; durable copy kept in SharePoint `InquiryMedia`. 30 MB limit
- `POST /api/inquiries/{id}/notes` body `{ author, body }` — append-only note
- `POST /api/inquiries/{id}/quotations` body `{ hskNumber, linkedBy }` — link an HSK# quote (→ Quoted; validated `(HSK-)?(SO|PO|Q)<digits>`)
- `PATCH /api/inquiries/{id}` body `{ status?, assignedTo? }` — update status / reassign (steal)
- `POST /api/inquiries/{id}/read` — mark the whole inquiry read
- `POST /api/inquiries/{id}/messages/{messageSpItemId}/read` body `{ read }` — per-message read toggle (read state is button-only)
- `POST /api/inquiries/{id}/read-all` body `{ read }` — mark all messages read/unread
- `POST /api/inquiries/{id}/regenerate-draft` — regenerate the AI suggestion for the latest inbound (supersedes prior pending drafts)
- `POST /api/inquiries/{id}/drafts/{draftId}/accept` body `{ from? }` — send the AI draft (opt-out → 409)
- `POST /api/inquiries/{id}/drafts/{draftId}/dismiss` — dismiss a pending draft

**MessagingController endpoints (`/api/messages`)** — the older generic messaging surface (predates Pulse):
- `GET /api/messages/conversations?top=` · `GET /api/messages/conversation/{id}?top=` — thread summaries / one thread
- `POST /api/messages/send` body `{ from, to, body, subject?, channel }` — `channel` ∈ sms | email | internal
- `POST /api/messages/read/{conversationId}` · `GET /api/messages/users`

**SmsWebhookController endpoints (`/api/sms`)** — SignalWire ingress (auth-exempt, loopback-bound; HMAC-validated):
- `POST /api/sms/inbound` — inbound SMS/MMS webhook: validates the SignalWire signature, acks fast (empty TwiML), enqueues a `Job` (From/To/Body/Sid/MediaUrls) to the `sms-inbound-jobs` dedup queue. **The public ingress is the `OutlookShredder.SmsWebhook` Azure Function**, which enqueues to the same queue; this in-proxy route is the loopback/Cloudflare receiver.
- `POST /api/sms/status` — delivery-status callback → `InquiryService.UpdateMessageStatusAsync(sid, status)`
- `POST /api/sms/dev-seed` body `{ from, body }` — DEV-ONLY (gated by `SignalWire:AllowDevSeed`, 404 otherwise): inject an inbound straight into the pipeline, bypassing the signature, to populate a thread without a live webhook

### Services

**`IAiExtractionService`** (interface) + **`AiServiceFactory`** (singleton)
- Provider selected at startup via `AI:Provider` in appsettings (`"claude"` | `"gemini"` | `"roundrobin"`; default `"claude"`)
- `AiServiceFactory.GetService()` resolves once and caches the wrapper so any per-instance state (e.g. the round-robin counter) persists across calls.
- **`"claude"` / `"gemini"`**: returns the named primary. If the *other* provider's API key is also configured, wraps it in `FallbackAiExtractionService` so a thrown exception from the primary (after its own internal retries) is transparently retried against the secondary. If only one API key is configured, the primary is returned directly with no fallback.
- **`"roundrobin"`** (alias `"round-robin"`): alternates between Claude and Gemini per call via `RoundRobinAiExtractionService` (`Interlocked.Increment` counter, starts on Claude), cross-falling back to the other provider if the chosen one throws. Requires both API keys; degrades gracefully to single-provider mode with a warning if only one is set.
- Caller-initiated cancellation (`CancellationToken`) is honoured in both wrappers: the secondary is **not** tried when the caller has cancelled.
- To add a new provider: implement `IAiExtractionService`, register it in `Program.cs`, add a case in `AiServiceFactory.ResolveByName` and `IsProviderKeyConfigured`

**`ClaudeExtractionService`** (singleton) — provider `"claude"`
- `ExtractRfqAsync(ExtractRequest, CancellationToken)` → `RfqExtraction`
- Uses **tool_use** (`extract_rfq` tool) for schema-enforced output — no free-text JSON parsing
- Static system prompt + tool definition sent with **prompt caching** (`anthropic-beta: prompt-caching-2024-07-31`)
- Retries 429/5xx/network errors up to `Claude:MaxRetries` with randomised jitter
- Logs warning when `stop_reason == "max_tokens"` or content is truncated at `Claude:MaxContentChars`
- Extraction fields: jobReference, quoteReference, supplierName, freightTerms, products[]
  (dateOfQuote/estimatedDeliveryDate removed — dates come from the RFQ Reference record, not extraction)

**`GeminiExtractionService`** (singleton) — provider `"gemini"`
- Uses `Mscc.GenerativeAI` SDK (v3.x) with `GoogleAI` → `GenerativeModel` factory pattern
- JSON mode (`ResponseMimeType = "application/json"`) + `ResponseSchema` (ParameterType enum, not SchemaType)
- System prompt injected via `systemInstruction` on the model (not as a user turn)
- PDF/DOCX attachments passed as `InlineData { MimeType, Data }` parts in the parts list
- Retries up to `Gemini:MaxRetries` with same jitter table as Claude
- Config keys: `Google:ApiKey`, `Gemini:Model` (default `gemini-2.0-flash`), `Gemini:MaxRetries`, `Gemini:MaxContentChars`, `Gemini:MaxContextChars`
- System prompt is kept **aligned with Claude's** — both prompts are canonical and describe the same extraction schema. Edit both together when changing extraction behaviour.
- Regret handling: supplier returns `supplierProductComments` explaining inability to supply (e.g. "Supplier regrets — unable to supply"). Does **not** use a sentinel `productName="Regret"`; that sentinel is filtered out on the Shredder side as a safety net for older data.

**`SharePointService`** (singleton)
- All Graph API calls. Uses `ClientSecretCredential` (app-only, `Sites.FullControl.All`).
- `WriteProductRowAsync(extraction, productLine, request, source, sourceFile, index, messageId?)` → `SpWriteResult`
  - Deduplicates by email+product; prefers attachment source over body
  - OOF detection; resolves supplier via `SupplierCacheService`
  - `messageId` stored on both SR and SLI rows for dedup keying
- `BackfillMessageIdsAsync(mail, days, ct)` → `(Patched, Skipped)` — post-hoc MessageId population for older rows
- `DeduplicateSupplierResponsesAsync(days, ct)` → `(SrDeleted, SliDeleted)` — removes no-MessageId orphans and duplicate-MessageId extras
- `ReadQcListAsync()` → `{ columns, rows, itemIds, lastModified }` — itemIds are SharePoint item IDs, parallel-indexed with rows[]
- `GetQcLastModifiedAsync()` → `DateTime?`
- `UpdateQcLqAsync()` → `(updated count, misses list)` — derives $/lb from quote rows, updates QC 'LQ' column
- `UpdateQcRowAsync(itemId, qc, qcCut)` — patches QC and QC Cut fields on a single SP item; resolves internal column names automatically
- `GetPublishVersionAsync()` → version string
- `EnsureSupplierListsAsync()` — idempotent list creation (provisions SupplierResponses, SupplierLineItems, PurchaseOrders)
- `WritePurchaseOrderAsync(rfqId, supplierName, poNumber, receivedAt, messageId, lineItemsJson)` → deduped by MessageId
- `ReadPurchaseOrdersAsync()` → `List<PurchaseOrderRecord>` — all PO rows
- `GET /api/purchase-orders` controller endpoint — Shredder loads this on startup
- `PrewarmAsync(ct)` — site ID + hot list IDs + one throwaway Graph call. Called once from `Program.cs` at startup.
- `ReadConversationAsync(rfqId, supplierName)` → `List<ConversationMessage>` — inbound (SR) + outbound (SupplierConversations) merged, ordered by `SentAt`
- `ReadOutboundConversationAsync(rfqId, supplierName)` → outbound-only (skips SR scan — fast path)
- `WriteConversationMessageAsync(msg)` → SP item ID of the new row

**`MailService`** (singleton)
- Graph API for mailbox (app-only, `Mail.ReadWrite` + `Mail.Send`)
- `SendRfqEmailAsync(subject, body, bccAddresses)` — sends via Graph
- `SendSupplierInquiryAsync(to, subject, body, attachmentName?, attachmentBytes?, attachmentContentType?)` — sends a single-recipient follow-up to one supplier about an ongoing RFQ; saves to Sent Items
- `GetMessageByIdAsync(mailbox, messageId)` → message metadata + body
- `MarkProcessedAsync(mailbox, itemId, extra)` — stamps "RFQ-Processed" category
- Strips RE:/FW: prefixes, [EXTERNAL] tags, converts HTML → plain text
- Extracts job references via regex `RFQ\s+\[([A-Za-z0-9]+)\]`

**`MailPollerService`** (hosted service — background)
- Polls inbox every `Mail:PollIntervalSeconds` (default 30s) for messages without "RFQ-Processed"
- Per message: strips FW:/RE:/[EXTERNAL] prefixes, then routes:
  - Subject starts with `"Purchase Order #HSK-PO"` → `ProcessPurchaseOrderAsync` (extracts PDF via Claude, writes to `PurchaseOrders` SP list, publishes `EventType="PO"` to Service Bus, stamps "RFQ-Processed"+"PO-Processed")
  - Everything else → normal RFQ pipeline (Claude/Gemini extract → SharePoint → notify)
- Processes emails concurrently via `Parallel.ForEachAsync` with `MaxDegreeOfParallelism = MaxConcurrency`
- Rate-limited via sliding-window: tracks AI call timestamps in `_aiCallTimestamps`, blocks until a slot is available within the current 60-second window
- `ReprocessMessagesAsync(messageIds[])` — manual re-extraction of arbitrary message IDs; blocks until complete; routes POs correctly; tracks progress with `Interlocked.Increment`
- `GetStatus()` → `MailStatus` — live snapshot used by `GET /api/mail/status`
- In-flight tracking: `ProcessMessageAsync` registers each message in `_inFlight` (subject, from, startedAt) on entry and removes it in a `finally` block
- Config:
  - `Mail:MailboxAddress` — required
  - `Mail:LookbackHours` (default 24) — rolling window per poll cycle
  - `Mail:MaxEmailsPerMinute` (default **100**) — AI calls per minute across all concurrent slots
  - `Mail:MaxConcurrency` (default **8**) — parallel processing slots
  - `Mail:BodyContextChars` (default 2000) — email body chars included when processing an attachment
  - `Mail:PollIntervalSeconds` (default 30) — poll frequency
  - `Mail:ExtractBodyWithoutJobRef` (default false) — when false, body-only emails with no job ref get a placeholder row under [000000] without calling AI

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

**`MailStatus`** — returned by `GET /api/mail/status`
```
Poller: PollerStatus, Reprocess: ReprocessStatus, RateLimit: RateLimitStatus, InFlight: List<InFlightItem>
```
- `PollerStatus` — `{ Running, LastPollAt, MessagesFoundLastCycle }`
- `ReprocessStatus` — `{ Active, Total, Completed, Failed, PercentComplete }`
- `RateLimitStatus` — `{ CallsInLastMinute, MaxPerMinute, SlotsAvailable }`
- `InFlightItem` — `{ Subject, From, StartedAt }`

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
| `SupplierConversations` | RFQ_ID, SupplierName, Direction (in\|out), MessageId, InReplyTo, SentAt, EmailSubject, BodyText, HasAttachments, ExtractedPricing — indexed on RFQ_ID + SupplierName |

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
| `Mail:MaxEmailsPerMinute` | `100` | AI API calls per minute (across all concurrent slots) |
| `Mail:MaxConcurrency` | `8` | Parallel processing slots per poll/reprocess cycle |
| `Mail:BodyContextChars` | `2000` | Email body chars included as context when processing an attachment |
| `Mail:ExtractBodyWithoutJobRef` | `false` | When false, body-only emails with no job ref skip AI and get a placeholder row |
| `Proxy:AllowedOrigin` | `https://localhost:3000` | CORS origin for AddinHost |

### Logging
Serilog — console + rolling daily file at `Logs/proxy-.log`.

---

## AddinHost (`OutlookShredder.AddinHost/`)

Minimal ASP.NET Core 10 static file server (HTTPS on port 3000). Serves the Office.js taskpane.

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

## Import Directory — Customer / Contact Bulk Loads

**Endpoint:** `POST /api/import/run`

**Directory:** `Import:Directory` in `appsettings.json` (default: `%LOCALAPPDATA%\Shredder\Import\`).
Drop CSV files into this folder, then POST to the endpoint. Processed files move to `Import\processed\{yyyyMMdd_HHmmss}_{filename}` so they are never re-processed.

**File type detection (first match wins):**
- Filename contains `customer info` / `customerinfo` / `cust info` / `custinfo` → **Customer Info** (rich BP master)
- Filename contains `partner`, ` bp`, `_bp`, or starts with `bp` → **Business Partners** (Name + Popup Message)
- Filename contains `contact` → **Contacts** (Customer Name + Contact Name + Phone)
- Header row contains `Business Partner` **and** `Margin Type` → Customer Info
- Header row contains `Popup Message` → Business Partners
- Header row contains both `Contact Name` and `Customer Name` → Contacts
- Otherwise: file is skipped with a warning in the response

**Load order:** files are classified up front and processed **Business Partners → Contacts → Customer Info LAST**. Customer Info only *enriches* records the partner load already created, so it must run after every new record exists. All loads write in **paced batches** (`SharePointService.RunBatchedAsync`, throttle-retry via `PatchListItemWithRetryAsync`) so a 12k-row file doesn't flood Graph/SharePoint.

**Business Partner processing rules:**
- Unique key: BP name (case-insensitive)
- BP names containing `duplicate` or `do not use` (any case) are silently skipped — ERP marks these as invalid entries that interfere with lookups
- Duplicate names within the same file are also skipped (second occurrence)
- The partner load persists `Name` (Title) + `PopupMessage`. Richer per-customer fields (Payment Terms, credit limits, sales/margin stats, etc.) come from the **Customer Info** load below, which enriches the same record.

**Contact processing rules:**
- Unique key: (CustomerName, ContactName, Phone) triple
- Phones are normalised to 10-digit US format (strips formatting, drops leading country code `1` if 11 digits)
- Contacts with invalid/missing phones are skipped with a warning
- Phones shared by 2+ contacts at the same BP are treated as company/general-line numbers and de-prioritised (contact's own unique phone is preferred)
- On upsert: **additive only** — existing `(CustomerName, ContactName, Phone)` triples are never deleted. Only triples not already present in SP are inserted. Customers accumulate phone numbers over time.

**Customer Info processing rules** (the rich BP master — `ExportedData (5).csv` style, ~23 columns):
- Match key: `Business Partner` name == the Customers list `Title` (same key the partner load writes) — case-insensitive. Dedup within the file (first kept); dupes are reported.
- **Enrichment only — UPDATES existing Customers records; never creates them.** Names present in the file but absent from the Customers list are *reported* as "candidates to add" (in the response + the review CSV), not added — so run the partner load first.
- Only **changed** fields are written (canonical compare ignores formatting noise like `150000` vs `150000.0`); each write stamps `CustomerInfoUpdatedAt`. Blank cells never overwrite an existing value.
- Schema is one source of truth: `CustomerInfoSchema` (in `CustomerImportService.cs`) maps each CSV header → SharePoint column + kind, and drives parsing, **Customers**-list column provisioning (`EnsureCustomerListsAsync`, auto-run at startup), change detection, and writes. Columns added: `Active, BpCategory, PaymentTerms, OnHold, PaymentMethod, PrimaryContact, ContactPhone, ContactEmail, CreditLineLimit, CreditAvailable, AutoInvoice, AutoStatement, CurrentBalance, WinPct, WinPctTransactions, SalesLast6Mo, SalesLast12Mo, AvgInvoiceValue, AvgMarginPct, TaxExempt, HowDidYouHear, MarginType, CustomerInfoUpdatedAt`.
- Endpoints: drop-folder `POST /api/import/run` (used by the Shredder "Import Customer & Contact Data" dialog), or direct `POST /api/customers/import-customer-info[?dryRun=true]` (raw CSV body). Unparseable numeric/boolean cells are reported for review, never fatal.
- **Active = logical-delete.** `Active=false`/NO marks a customer inactive (ERP-hidden). Existing records are still enriched and flagged; an *unmatched* inactive row is dead data and is NOT surfaced as an add-candidate (`SkippedInactive`). **The data layer honors this:** `CustomerCacheService` excludes inactive customers AND their contacts (a contact inherits its customer's active status — no per-contact flag) from every lookup (`GetAllPartnerNames`, phone lookup, `GetAllContacts`, the `/business-partners` + `/contacts` endpoints). Default = hide inactive; set `Customers:IncludeInactive=true` to surface them. The import path reads SP directly, so it still sees/marks inactive records.

**Automation:** A `FileSystemWatcherImportService` (auto-trigger on file drop) is in `todos.md` as a future enhancement. Currently requires a manual `POST /api/import/run`.

## Debugging & Testing Methodology

**Always prefer testing and direct observation over inference.** When diagnosing a bug or verifying behavior:

1. **Test first**: Use proxy API endpoints (`/api/catalog/resolve`, `/api/catalog/diagnose`, `/api/items/by-rfq/{rfqId}`, `/api/mail/status`, etc.), proxy logs, and SP data to observe actual behavior before reading code.
2. **Then trace**: After observing the result, use code reading to understand *why* that path was taken — map the observed output back to specific code branches.
3. **Build test endpoints if needed**: If no API exists to observe a behavior directly, add a lightweight diagnostic endpoint rather than relying on inference from code reading alone.

This approach gives deterministic diagnosis. Inference from code reading alone is error-prone — the same code can produce different results depending on cache state, SP data, config, and timing. Observation anchors the diagnosis in reality.

## Rules

- **Date/time columns are native `dateTime`, never text.** SharePoint string-sorts text, which misorders dates
  (proven 0/25 correct on the legacy ErpDocuments text date). Store timestamps in native **indexed** `dateTime`
  columns and sort/filter server-side (`$orderby=fields/XxxDt desc`, `$filter=fields/XxxDt ge '…Z'`). Legacy text
  date columns are being migrated to native — see `wip/datetime-column-migration.md`. Pattern: add native `XxxDt`
  column + index; **dual-write it at EVERY write site** (values that aren't already ISO — e.g. "Jul-01-2026" —
  parse via `ToIsoOrNull` first); backfill via `SharePointService.BackfillTextDateToDateTimeAsync`; register one
  row in `SharePointService.DateTimeColumnMigrations()` (SP lists) or `SharePointInquiryStore
  .BackfillDateTimeColumnsAsync` (inquiry cluster) so the sweep + backfill-all cover it automatically. The
  `DateTimeBackfillSweepService` heals text-date drift from not-yet-updated fleet proxies every 30 min —
  **DISABLE it (`DateTimeBackfill:SweepEnabled=false`) once the whole fleet is on the dual-write build** and the
  log shows `[DtSweep] no drift`. See [[feedback_inquiry_timestamp_sort]].
- **Full-regret rule:** a SupplierResponse is a FULL regret (`IsRegret=true` at SR level) only when the email **body expresses regret AND there is no priced line in the extraction** (no PDF attachment with valid pricing — any price counts). A priced attachment means a real quote, or — when the body regrets but the PDF prices items — a substitute/partial; never a full regret. Per-line `IsRegret` on each SLI is computed independently (`!HasPrice(product) && (product.IsRegret || regret phrase)`). Set in `EnsureSupplierResponseAsync` (`blanketRegret`).
- **Job-ref mismatch:** when a PDF's AI-extracted JobReference is a valid RFQ id that differs from the email-subject ref, the priced data files under the PDF's **own** ref and a `⚠` warning is written to `SupplierProductComments` (visible in the grid). A foreign PDF that leaks a stale `SourceFile` pointer onto the wrong SR is cleaned with `POST /api/sr/{srId}/detach-file`. Diagnose any RFQ with `GET /api/diag/extraction-trace?rfqId=` (live email vs extracted vs stored, each block carries its source).
- **Graph MessageId is case-SENSITIVE.** Immutable message IDs are case-sensitive base64 — two IDs differing only by one character's case are DIFFERENT emails. Always compare/group MessageId with `StringComparison.Ordinal` / `StringComparer.Ordinal`, never `OrdinalIgnoreCase`. An ignore-case compare in the SupplierResponse dedup collapsed two distinct supplier emails (IDs differing by `P` vs `p`) onto one SR row, letting each overwrite the other's RFQ/attachment/line items — the root cause of the RSX9K3 wrong-PDF misroute (Penn Steel quote `…RPAAA=`→RSX9JW vs regret `…RpAAA=`→RSX9K3). Diagnose collisions with the extraction-trace; recover with detach-file / delete-SR + targeted reprocess.
- `@mithrilmetals.com` is never a valid supplier — never appear in extraction results or email targets
- All SharePoint writes go through `SharePointService` — no direct Graph calls from controllers
- Extraction prompt and JSON schema live in `ClaudeExtractionService.ExtractRfqAsync` and `GeminiExtractionService` — edit there to change extraction behaviour. The system prompt text is a `const string` near the top of each file.
- **RFQ ID format:** RFQ IDs are 2-letter user initials + 4 Crockford Base32 chars (6 total, e.g. `AW0001`). The first two chars are ALWAYS letters — this excludes a supplier's own job number such as Eastern Metal's `J06601` (single letter + digits), which the retired generic `[A-Z0-9]{6}` rule used to mis-capture. The `HQ`+6 (8-char) format and the generic 6-char rule were both **retired 2026-06**. The orphan sentinel `[000000]` is unchanged. The canonical pattern is `[A-Z]{2}[A-Z0-9]{4}` (IgnoreCase) — used by `JobRefRegex`/`JobRefBareRegex` in `MailPollerService` and `HackensackPollerService`, the reparent validators in `ExtractController`, `IsValidRfqId` in `SharePointService`, the extraction tool-schema descriptions/prompts in `ClaudeExtractionService`/`GeminiExtractionService`, the taskpane `JOB_REF_REGEX` in `jobRef.js`, and (client-side) `RfqDisplayProcessor.IsValidRfqId`. Note: the `[SHR:{rfqId}]` token parser in `ShrConvInRouter` keeps its own broader alternation — it's anchored by the `SHR:` prefix, so it's not a collision source.
