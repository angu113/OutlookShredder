using System.Text.Json;
using System.Text.RegularExpressions;
using Azure.Identity;
using Microsoft.Graph;
using Microsoft.Graph.Models;
using OutlookShredder.Proxy.Models;

namespace OutlookShredder.Proxy.Services;

/// <summary>
/// Manages all SharePoint list operations via Microsoft Graph (app-only auth).
///
/// Lists:
///   SupplierResponses  — one row per supplier email; holds email metadata + body + attachment
///   SupplierLineItems  — one row per extracted product line; child of SupplierResponses
///   RFQ References     — source RFQs written by ShredderXL; holds Notes field
///
/// Azure AD app requires:  Sites.FullControl.All  (Application, admin consented)
/// </summary>
public class SharePointService
{
    private readonly IConfiguration          _config;
    private readonly ILogger<SharePointService> _log;
    private readonly SupplierCacheService  _suppliers;
    private readonly ProductCatalogService _catalog;

    private GraphServiceClient?     _graph;
    private ClientSecretCredential? _spCredential;
    private string? _siteId;

    // Cached list IDs (lazy-resolved)
    private string? _srListId;      // SupplierResponses
    private string? _sliListId;     // SupplierLineItems
    private string? _rfqRefListId;  // RFQ References
    private string? _listId;        // RFQ Line Items (legacy — kept for EnsureColumnsAsync)

    private static readonly string[] _regretPhrases =
        ["regret", "no stock", "unable to supply", "cannot supply", "not available", "out of stock"];

    public SharePointService(IConfiguration config, ILogger<SharePointService> log,
        SupplierCacheService suppliers, ProductCatalogService catalog)
    {
        _config    = config;
        _log       = log;
        _suppliers = suppliers;
        _catalog   = catalog;
    }

    // ── Graph client (lazy init) ─────────────────────────────────────────────
    private GraphServiceClient GetGraph()
    {
        if (_graph is not null) return _graph;

        var tenantId     = _config["SharePoint:TenantId"]     ?? throw new InvalidOperationException("SharePoint:TenantId not set");
        var clientId     = _config["SharePoint:ClientId"]     ?? throw new InvalidOperationException("SharePoint:ClientId not set");
        var clientSecret = _config["SharePoint:ClientSecret"] ?? throw new InvalidOperationException("SharePoint:ClientSecret not set");

        var credential = new ClientSecretCredential(tenantId, clientId, clientSecret);
        _graph = new GraphServiceClient(credential, ["https://graph.microsoft.com/.default"]);
        return _graph;
    }

    // ── SharePoint REST credential (separate audience from Graph) ────────────
    private ClientSecretCredential GetSpCredential()
    {
        if (_spCredential is not null) return _spCredential;

        var tenantId     = _config["SharePoint:TenantId"]     ?? throw new InvalidOperationException("SharePoint:TenantId not set");
        var clientId     = _config["SharePoint:ClientId"]     ?? throw new InvalidOperationException("SharePoint:ClientId not set");
        var clientSecret = _config["SharePoint:ClientSecret"] ?? throw new InvalidOperationException("SharePoint:ClientSecret not set");

        _spCredential = new ClientSecretCredential(tenantId, clientId, clientSecret);
        return _spCredential;
    }

    // ── Site ID (cached) ─────────────────────────────────────────────────────
    private async Task<string> GetSiteIdAsync()
    {
        if (_siteId is not null) return _siteId;

        var siteUrl = _config["SharePoint:SiteUrl"]
            ?? "https://metalsupermarkets-my.sharepoint.com/personal/angus_mithrilmetals_com";

        var uri  = new Uri(siteUrl);
        var host = uri.Host;
        var path = uri.AbsolutePath;

        _log.LogInformation("[SP] Resolving site: {Host}{Path}", host, path);

        var site = await GetGraph().Sites[$"{host}:{path}"].GetAsync();
        _siteId  = site!.Id ?? throw new Exception("Could not resolve SharePoint site ID");

        _log.LogInformation("[SP] Site ID: {Id}", _siteId);
        return _siteId;
    }

    // ── List ID getters (cached) ─────────────────────────────────────────────

    private async Task<string> GetSupplierResponsesListIdAsync()
    {
        if (_srListId is not null) return _srListId;
        _srListId = await ResolveListIdAsync("SupplierResponses");
        return _srListId;
    }

    private async Task<string> GetSupplierLineItemsListIdAsync()
    {
        if (_sliListId is not null) return _sliListId;
        _sliListId = await ResolveListIdAsync("SupplierLineItems");
        return _sliListId;
    }

    private async Task<string> GetRfqReferencesListIdAsync()
    {
        if (_rfqRefListId is not null) return _rfqRefListId;
        _rfqRefListId = await ResolveListIdAsync("RFQ References");
        return _rfqRefListId;
    }

    // Legacy — used by EnsureColumnsAsync only
    private async Task<string> GetListIdAsync()
    {
        if (_listId is not null) return _listId;
        _listId = await ResolveListIdAsync(_config["SharePoint:ListName"] ?? "RFQ Line Items");
        return _listId;
    }

    private async Task<string> ResolveListIdAsync(string listName)
    {
        var siteId = await GetSiteIdAsync();
        var lists  = await GetGraph().Sites[siteId].Lists
            .GetAsync(req => req.QueryParameters.Filter = $"displayName eq '{listName}'");
        var list = lists?.Value?.FirstOrDefault()
            ?? throw new Exception($"SharePoint list '{listName}' not found. Run POST /api/setup-supplier-lists.");
        var id = list.Id ?? throw new Exception($"List '{listName}' ID was null");
        _log.LogInformation("[SP] List '{Name}' -> id: {Id}", listName, id);
        return id;
    }

    // ── Read: SupplierLineItems joined with SupplierResponses ────────────────

    /// <summary>
    /// Returns up to <paramref name="top"/> SupplierLineItems merged with their parent
    /// SupplierResponse fields as flat field dictionaries (matches the shape expected
    /// by the Shredder dashboard).
    /// </summary>
    public async Task<List<Dictionary<string, object?>>> ReadSupplierItemsAsync(int top = 500)
    {
        var siteId   = await GetSiteIdAsync();
        var srListId  = await GetSupplierResponsesListIdAsync();
        var sliListId = await GetSupplierLineItemsListIdAsync();

        // Fetch both lists in parallel
        var srTask  = GetGraph().Sites[siteId].Lists[srListId].Items
            .GetAsync(req => { req.QueryParameters.Expand = ["fields"]; req.QueryParameters.Top = top; });
        var sliTask = GetGraph().Sites[siteId].Lists[sliListId].Items
            .GetAsync(req => { req.QueryParameters.Expand = ["fields"]; req.QueryParameters.Top = top; });
        await Task.WhenAll(srTask, sliTask);

        // Extract a string value from an AdditionalData entry, handling JsonElement.
        // Graph SDK deserialises all field values as JsonElement; calling .GetString() on a
        // String-kind element returns the raw string without quotes.
        // Accepts both object and object? via the nullable-erased object? parameter.
        static string? Str(object? v) => v switch
        {
            string s => s,
            JsonElement je when je.ValueKind == JsonValueKind.String => je.GetString(),
            JsonElement je => je.ToString(),   // numbers, bools, etc.
            null => null,
            _ => v.ToString()
        };
        static string? GetStr(IDictionary<string, object?> d, string key) =>
            d.TryGetValue(key, out var v) ? Str(v) : null;
        static string? GetStrRaw(IDictionary<string, object> d, string key) =>
            d.TryGetValue(key, out var v) ? Str(v) : null;

        // Build lookup: SupplierResponse SP item ID → its fields
        var srById = (srTask.Result?.Value ?? [])
            .Where(i => i.Id is not null && i.Fields?.AdditionalData is not null)
            .ToDictionary(i => i.Id!, i => i.Fields!.AdditionalData!);

        // Fallback lookup: "RFQ_ID|SupplierName" → (SP item ID, SR fields).
        // Used when SupplierResponseId is missing or stale (e.g. data written before the
        // column existed, or before the upsert logic was corrected).
        // Stores the SR item ID so we can correct row["SupplierResponseId"] on a fallback hit,
        // which lets the client fetch the right SP attachment.
        var srBySupplierRfq = new Dictionary<string, (string SrId, IDictionary<string, object> Fields)>(StringComparer.OrdinalIgnoreCase);
        foreach (var (srItemId, srRaw) in srById)
        {
            // SR list may return RFQ_ID as "RFQ_ID" or "RFQ_x005F_ID" depending on how
            // the column was originally created.
            var rfq = GetStrRaw(srRaw, "RFQ_ID") ?? GetStrRaw(srRaw, "RFQ_x005F_ID");
            var sn  = GetStrRaw(srRaw, "SupplierName");
            if (rfq is not null && sn is not null)
                srBySupplierRfq.TryAdd($"{rfq}|{sn}", (srItemId, srRaw));  // first-wins; dedup
        }

        _log.LogDebug("[SP] ReadSupplierItems: {SrCount} SR rows, {SliCount} SLI rows",
            srById.Count, sliTask.Result?.Value?.Count ?? 0);

        static bool IsAppField(string key) =>
            !key.StartsWith('@') &&
            !key.StartsWith('_') &&
            key is not ("LinkTitle" or "LinkTitleNoMenu" or "ContentType"
                     or "Edit" or "Attachments" or "ItemChildCount" or "FolderChildCount"
                     or "AuthorLookupId" or "EditorLookupId"
                     or "AppAuthorLookupId" or "AppEditorLookupId");

        // Parent fields to promote into each line-item dict
        string[] parentFields = [
            "EmailFrom", "ReceivedAt", "ProcessedAt", "ProcessingSource",
            "SourceFile", "DateOfQuote", "EstimatedDeliveryDate",
            "QuoteReference", "FreightTerms", "EmailBody", "EmailSubject"
        ];

        var result = new List<Dictionary<string, object?>>();

        foreach (var sli in sliTask.Result?.Value ?? [])
        {
            if (sli.Fields?.AdditionalData is null) continue;

            var row = sli.Fields.AdditionalData
                .Where(kv => IsAppField(kv.Key))
                .ToDictionary(kv => kv.Key, kv => (object?)kv.Value);

            // ── Resolve parent SupplierResponse record ─────────────────────
            IDictionary<string, object>? srMatch = null;

            // Primary join: SupplierResponseId → SR item's SP integer ID
            var srId = GetStr(row, "SupplierResponseId");
            if (srId is not null)
            {
                srById.TryGetValue(srId, out srMatch);
                if (srMatch is null)
                    _log.LogDebug("[SP] SLI {SliId}: SupplierResponseId={SrId} not found in srById (keys: {Keys})",
                        sli.Id, srId, string.Join(",", srById.Keys.Take(10)));
            }

            // Fallback join: RFQ_ID + SupplierName (handles stale/missing SupplierResponseId)
            if (srMatch is null)
            {
                var sliRfq = GetStr(row, "RFQ_ID") ?? GetStr(row, "RFQ_x005F_ID");
                var sliSn  = GetStr(row, "SupplierName");
                if (sliRfq is not null && sliSn is not null &&
                    srBySupplierRfq.TryGetValue($"{sliRfq}|{sliSn}", out var fb))
                {
                    srMatch = fb.Fields;
                    // Correct the stale SupplierResponseId so the client fetches the right SP attachment
                    row["SupplierResponseId"] = fb.SrId;
                    _log.LogDebug("[SP] SLI {SliId} [{Rfq}/{Supplier}]: joined via fallback, corrected SrId {OldId}→{NewId}",
                        sli.Id, sliRfq, sliSn, srId ?? "null", fb.SrId);
                }
            }

            if (srMatch is not null)
            {
                // Rename RFQ_ID → JobReference for backward compat with the Shredder DTO.
                // Handle both possible internal column names.
                var rfqIdVal = GetStrRaw(srMatch, "RFQ_ID") ?? GetStrRaw(srMatch, "RFQ_x005F_ID");
                if (rfqIdVal is not null) row["JobReference"] = rfqIdVal;

                // Lift email-level fields if not already on the line item
                foreach (var f in parentFields)
                    if (!row.ContainsKey(f) && srMatch.TryGetValue(f, out var v))
                        row[f] = v;

                // If SR is blanket-regret and line item has no pricing, inherit regret flag
                if (srMatch.TryGetValue("IsRegret", out var srRegret) &&
                    srRegret is true or JsonElement { ValueKind: JsonValueKind.True })
                {
                    if (!row.ContainsKey("PricePerPound") || row["PricePerPound"] is null)
                        if (!row.ContainsKey("PricePerFoot") || row["PricePerFoot"] is null)
                            row["IsRegret"] = true;
                }
            }

            result.Add(row);
        }

        return result;
    }

    // ── Read: RFQ References (for Notes) ────────────────────────────────────

    public async Task<List<Dictionary<string, object?>>> ReadRfqReferencesAsync()
    {
        var siteId = await GetSiteIdAsync();
        var listId = await GetRfqReferencesListIdAsync();
        var col    = await ResolveRfqIdColumnAsync(siteId, listId);

        var result = await GetGraph().Sites[siteId].Lists[listId].Items
            .GetAsync(req =>
            {
                req.QueryParameters.Expand = [$"fields($select={col},Notes,Requester,DateCreated)"];
                req.QueryParameters.Top    = 500;
            });

        static string? FieldStr(IDictionary<string, object?> d, string key) =>
            d.TryGetValue(key, out var v) ? v?.ToString() : null;

        return result?.Value?
            .Where(i => i.Fields?.AdditionalData is not null)
            .Select(i =>
            {
                var d = i.Fields!.AdditionalData!;
                var rfqId = FieldStr(d, col)
                         ?? FieldStr(d, "RFQ_x005F_ID")
                         ?? FieldStr(d, "RFQ_ID");
                return new Dictionary<string, object?>
                {
                    ["Id"]          = i.Id,
                    ["RFQ_ID"]      = rfqId,
                    ["Notes"]       = FieldStr(d, "Notes"),
                    ["Requester"]   = FieldStr(d, "Requester"),
                    ["DateCreated"] = d.TryGetValue("DateCreated", out var dc) ? dc : null,
                };
            })
            .Where(d => d["RFQ_ID"] is not null)
            .ToList()
            ?? [];
    }

    // ── Write: update Notes on an RFQ Reference ──────────────────────────────

    public async Task UpdateRfqNotesAsync(string rfqId, string notes)
    {
        var siteId = await GetSiteIdAsync();
        var listId = await GetRfqReferencesListIdAsync();
        var col    = await ResolveRfqIdColumnAsync(siteId, listId);

        var result = await GetGraph().Sites[siteId].Lists[listId].Items
            .GetAsync(r =>
            {
                r.QueryParameters.Filter = $"fields/{col} eq '{EscapeOdata(rfqId)}'";
                r.QueryParameters.Expand = [$"fields($select=id,{col})"];
                r.QueryParameters.Top    = 5;
                r.Headers.Add("Prefer", "HonorNonIndexedQueriesWarningMayFailRandomly");
            });

        var item = result?.Value?.FirstOrDefault();

        if (item is null)
        {
            // RFQ came in via supplier email, not the import tab — create the row on first note save.
            _log.LogInformation("[SP] RFQ Reference '{Id}' not found — creating it", rfqId);
            await GetGraph().Sites[siteId].Lists[listId].Items
                .PostAsync(new ListItem
                {
                    Fields = new FieldValueSet
                    {
                        AdditionalData = new Dictionary<string, object?>
                        {
                            [col]      = rfqId,
                            ["Notes"]  = notes,
                        }
                    }
                });
            _log.LogInformation("[SP] Created RFQ Reference '{Id}' with Notes", rfqId);
            return;
        }

        await GetGraph().Sites[siteId].Lists[listId].Items[item.Id!].Fields
            .PatchAsync(new FieldValueSet
            {
                AdditionalData = new Dictionary<string, object?> { ["Notes"] = notes }
            });

        _log.LogInformation("[SP] Updated Notes for RFQ '{Id}'", rfqId);
    }

    // ── RFQ Import: read / write RFQ References + RFQ Line Items ────────────

    private string? _rfqLineItemsListId;

    private async Task<string> GetRfqLineItemsListIdAsync()
    {
        if (_rfqLineItemsListId is not null) return _rfqLineItemsListId;
        _rfqLineItemsListId = await ResolveListIdAsync("RFQ Line Items");
        return _rfqLineItemsListId;
    }

    /// <summary>Returns all RFQ_ID values that exist in the RFQ References list.</summary>
    public async Task<HashSet<string>> GetExistingRfqIdsAsync()
    {
        var siteId = await GetSiteIdAsync();
        var listId = await GetRfqReferencesListIdAsync();
        var col    = await ResolveRfqIdColumnAsync(siteId, listId);

        var items = await GetGraph().Sites[siteId].Lists[listId].Items
            .GetAsync(req =>
            {
                req.QueryParameters.Expand = [$"fields($select={col})"];
                req.QueryParameters.Top    = 5000;
            });

        var ids = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        foreach (var item in items?.Value ?? [])
        {
            if (item.Fields?.AdditionalData is null) continue;
            var d  = item.Fields.AdditionalData;
            // Try the resolved name first, then both fallback variants
            string? id = d.TryGetValue(col,            out var v0) ? v0?.ToString()
                       : d.TryGetValue("RFQ_x005F_ID", out var v1) ? v1?.ToString()
                       : d.TryGetValue("RFQ_ID",        out var v2) ? v2?.ToString()
                       : null;
            if (!string.IsNullOrEmpty(id)) ids.Add(id);
        }
        return ids;
    }

    /// <summary>Creates one row in the RFQ References list.</summary>
    public async Task CreateRfqReferenceAsync(RfqReferenceRequest req)
    {
        var siteId = await GetSiteIdAsync();
        var listId = await GetRfqReferencesListIdAsync();

        // Use whichever column name the list accepts (probe once, cache result).
        var col = await ResolveRfqIdColumnAsync(siteId, listId);

        await GetGraph().Sites[siteId].Lists[listId].Items
            .PostAsync(new ListItem
            {
                Fields = new FieldValueSet
                {
                    AdditionalData = new Dictionary<string, object?>
                    {
                        [col]               = req.RfqId,
                        ["Requester"]       = req.Requester,
                        ["DateCreated"]     = req.DateSent.ToUniversalTime().ToString("o"),
                        ["EmailRecipients"] = req.EmailRecipients,
                    }
                }
            });

        _log.LogInformation("[SP] Created RFQ Reference '{Id}'", req.RfqId);
    }

    /// <summary>Creates one row per entry in <paramref name="items"/> in the RFQ Line Items list.</summary>
    public async Task CreateRfqLineItemsAsync(IEnumerable<RfqLineItemRequest> items)
    {
        var siteId = await GetSiteIdAsync();
        var listId = await GetRfqLineItemsListIdAsync();
        var col    = await ResolveRfqIdColumnAsync(siteId, listId);

        foreach (var req in items)
        {
            var data = new Dictionary<string, object?> { [col] = req.RfqId };
            if (req.Mspc          is not null) data["MSPC"]           = req.Mspc;
            if (req.Product       is not null) data["Product"]        = req.Product;
            if (req.Units         is not null) data["Units"]          = req.Units;
            if (req.SizeOfUnits   is not null) data["SizeOfUnits"]    = req.SizeOfUnits;
            if (req.SupplierEmails is not null) data["SupplierEmails"] = req.SupplierEmails;

            await GetGraph().Sites[siteId].Lists[listId].Items
                .PostAsync(new ListItem { Fields = new FieldValueSet { AdditionalData = data } });
        }
    }

    // Cached: list ID → internal column name for RFQ_ID
    private readonly Dictionary<string, string> _rfqColByList = new();

    private async Task<string> ResolveRfqIdColumnAsync(string siteId, string listId)
    {
        if (_rfqColByList.TryGetValue(listId, out var cached)) return cached;

        var cols = await GetGraph().Sites[siteId].Lists[listId].Columns.GetAsync();
        var all  = cols?.Value ?? [];

        // Match on internal name (Name) or display name (DisplayName) — whichever has RFQ+ID.
        var col = all.FirstOrDefault(c =>
            c.Name?.Equals("RFQ_x005F_ID", StringComparison.OrdinalIgnoreCase) == true ||
            c.Name?.Equals("RFQ_ID",        StringComparison.OrdinalIgnoreCase) == true ||
            c.DisplayName?.Equals("RFQ_ID", StringComparison.OrdinalIgnoreCase) == true ||
            c.DisplayName?.Equals("RFQ ID", StringComparison.OrdinalIgnoreCase) == true);

        if (col is null)
        {
            var dump = string.Join("\n", all.Select(c => $"  Name={c.Name}  DisplayName={c.DisplayName}"));
            throw new InvalidOperationException(
                $"RFQ_ID column not found in list (id={listId}).\nAvailable columns:\n{dump}");
        }

        // Use the internal name (Name) for writes — DisplayName is just for matching.
        var name = col.Name
            ?? throw new InvalidOperationException(
                $"Column '{col.DisplayName}' has a null internal Name. Cannot write to it.");

        _rfqColByList[listId] = name;
        _log.LogInformation("[SP] RFQ_ID column resolved: Name={Name}  DisplayName={Display}",
            name, col.DisplayName);
        return name;
    }

    // ── Write: upsert one supplier email + all its extracted lines ───────────

    /// <summary>
    /// Main write entry point.  For each extracted email:
    ///   1. Upsert a <b>SupplierResponses</b> row (one per supplier email).
    ///   2. Upload the PDF attachment to that row.
    ///   3. Upsert a <b>SupplierLineItems</b> row for this product line.
    ///
    /// Uniqueness key for SupplierResponses: (RFQ_ID, EmailFrom).
    /// Uniqueness key for SupplierLineItems: (SupplierResponseId, ProductName).
    /// </summary>
    public async Task<SpWriteResult> WriteProductRowAsync(
        RfqExtraction  header,
        ProductLine    product,
        ExtractRequest emailMeta,
        string         source,
        string?        sourceFile,
        int            rowIndex)
    {
        var result = new SpWriteResult { ProductName = product.ProductName };
        try
        {
            var siteId = await GetSiteIdAsync();

            // ── Resolve job reference ────────────────────────────────────────
            var rawJobRef = (header.JobReference
                ?? emailMeta.JobRefs.FirstOrDefault()?.Trim('[', ']')
                ?? string.Empty).ToUpperInvariant();
            var jobRef = string.IsNullOrEmpty(rawJobRef) ? "000000" : rawJobRef;

            // ── Resolve supplier name ────────────────────────────────────────
            var rawSupplier = header.SupplierName;
            if (string.IsNullOrWhiteSpace(rawSupplier) && !string.IsNullOrWhiteSpace(emailMeta.EmailFrom))
            {
                var addr = emailMeta.EmailFrom;
                if (addr.Contains('@'))
                {
                    var domain = addr.Split('@').Last();
                    var parts  = domain.Split('.');
                    rawSupplier = parts.Length >= 2 ? parts[^2] : parts[0];
                }
                else rawSupplier = addr;
            }
            rawSupplier ??= string.Empty;
            var supplier = _suppliers.ResolveSupplierName(rawSupplier);
            if (supplier is null)
            {
                result.SupplierUnknown = true;
                supplier = "Unknown";
                jobRef   = "WHOIS";
                _log.LogInformation("[SP] Supplier '{Raw}' not in reference list — writing under [WHOIS]", rawSupplier);
            }

            // ── Upsert SupplierResponses ─────────────────────────────────────
            var srListId  = await GetSupplierResponsesListIdAsync();
            var srId      = await EnsureSupplierResponseAsync(
                siteId, srListId, jobRef, supplier, header, emailMeta, source, sourceFile);
            result.SpItemId = srId;

            // Upload the source attachment as a SharePoint list item attachment.
            if (result.SpItemId is not null &&
                emailMeta.SourceType == "attachment" &&
                !string.IsNullOrEmpty(emailMeta.FileName) &&
                !string.IsNullOrEmpty(emailMeta.Base64Data))
            {
                try
                {
                    var bytes = Convert.FromBase64String(emailMeta.Base64Data);
                    await UpsertItemAttachmentAsync(srId, srListId, emailMeta.FileName, bytes);
                }
                catch (Exception ex)
                {
                    _log.LogWarning(ex, "[SP] Attachment upload failed for SR {Id} ('{File}')", srId, emailMeta.FileName);
                }
            }

            // ── Upsert SupplierLineItems ─────────────────────────────────────
            var sliListId = await GetSupplierLineItemsListIdAsync();
            await WriteSupplierLineItemAsync(
                siteId, sliListId, srId, jobRef, supplier, product, rowIndex,
                sourceFile, emailMeta.EmailFrom);

            result.Success = true;
            result.Updated = false; // upsert logic handled internally
        }
        catch (Microsoft.Graph.Models.ODataErrors.ODataError odataEx)
        {
            result.Success = false;
            result.Error   = odataEx.Message;
            _log.LogError("[SP] ODataError: code={Code} msg={Msg}", odataEx.Error?.Code, odataEx.Error?.Message);
        }
        catch (Exception ex)
        {
            result.Success = false;
            result.Error   = ex.Message;
            _log.LogError(ex, "[SP] Failed to upsert product '{Name}'", product.ProductName);
        }
        return result;
    }

    // ── Upsert SupplierResponses (private) ───────────────────────────────────

    private async Task<string> EnsureSupplierResponseAsync(
        string siteId, string listId,
        string jobRef, string supplier,
        RfqExtraction header, ExtractRequest emailMeta,
        string source, string? sourceFile)
    {
        var existingId = await FindExistingSupplierResponseAsync(siteId, listId, jobRef, supplier);

        var emailBodyTrunc = (emailMeta.EmailBody ?? emailMeta.BodyContext) is string body
            ? body[..Math.Min(body.Length,
                  int.TryParse(_config["SharePoint:MaxEmailBodyChars"], out var mebc) ? mebc : 10_000)]
            : null;

        bool blanketRegret = HasRegretPhrase(emailMeta.EmailBody) || HasRegretPhrase(emailMeta.BodyContext);

        var title = $"[{jobRef}] {supplier} {(emailMeta.ReceivedAt is not null ? DateTime.Parse(emailMeta.ReceivedAt).ToString("yyyy-MM-dd") : "unknown")}";
        title = title[..Math.Min(title.Length, 255)];

        var fieldData = new Dictionary<string, object?>
        {
            ["Title"]                = title,
            ["RFQ_ID"]               = string.IsNullOrEmpty(jobRef) ? null : jobRef,
            ["SupplierName"]         = supplier,
            ["EmailFrom"]            = emailMeta.EmailFrom,
            ["ReceivedAt"]           = emailMeta.ReceivedAt,
            ["EmailSubject"]         = emailMeta.EmailSubject,
            ["EmailBody"]            = emailBodyTrunc,
            ["ProcessedAt"]          = DateTime.UtcNow.ToString("o"),
            ["ProcessingSource"]     = source,
            ["SourceFile"]           = sourceFile,
            ["QuoteReference"]       = header.QuoteReference,
            ["DateOfQuote"]          = header.DateOfQuote,
            ["EstimatedDeliveryDate"]= header.EstimatedDeliveryDate,
            ["FreightTerms"]         = header.FreightTerms,
            ["IsRegret"]             = blanketRegret,
        };

        if (existingId is not null)
        {
            await GetGraph().Sites[siteId].Lists[listId].Items[existingId].Fields
                .PatchAsync(new FieldValueSet { AdditionalData = fieldData });
            _log.LogInformation("[SP] Updated SupplierResponse {Id} for [{JobRef}] {Supplier}", existingId, jobRef, supplier);
            return existingId;
        }
        else
        {
            var item = await GetGraph().Sites[siteId].Lists[listId].Items
                .PostAsync(new ListItem { Fields = new FieldValueSet { AdditionalData = fieldData } });
            var newId = item!.Id!;
            _log.LogInformation("[SP] Created SupplierResponse {Id} for [{JobRef}] {Supplier}", newId, jobRef, supplier);
            return newId;
        }
    }

    private async Task<string?> FindExistingSupplierResponseAsync(
        string siteId, string listId, string jobRef, string supplierName)
    {
        if (string.IsNullOrEmpty(jobRef) || string.IsNullOrEmpty(supplierName))
            return null;

        var result = await GetGraph().Sites[siteId].Lists[listId].Items
            .GetAsync(r =>
            {
                r.QueryParameters.Filter = $"fields/RFQ_ID eq '{EscapeOdata(jobRef)}' and fields/SupplierName eq '{EscapeOdata(supplierName)}'";
                r.QueryParameters.Expand = ["fields($select=id,RFQ_ID,SupplierName)"];
                r.QueryParameters.Top    = 5;
                r.Headers.Add("Prefer", "HonorNonIndexedQueriesWarningMayFailRandomly");
            });

        return result?.Value?.FirstOrDefault()?.Id;
    }

    // ── Write SupplierLineItems (private) ────────────────────────────────────

    private async Task WriteSupplierLineItemAsync(
        string siteId, string listId,
        string supplierResponseId, string jobRef, string supplier,
        ProductLine product, int rowIndex,
        string? sourceFile, string? emailFrom)
    {
        var prodName   = product.ProductName ?? $"Product {rowIndex + 1}";
        var prodTokens = ProductTokens(prodName);

        var title = $"[{jobRef}] {supplier} - {prodName}";
        title = title[..Math.Min(title.Length, 255)];

        var fieldData = new Dictionary<string, object?>
        {
            ["Title"]                    = title,
            ["SupplierResponseId"]       = supplierResponseId,
            ["RFQ_ID"]                   = string.IsNullOrEmpty(jobRef) ? null : jobRef,
            ["SupplierName"]             = supplier,
            ["ProductName"]              = prodName,
            ["SourceFile"]               = sourceFile,
            ["EmailFrom"]                = emailFrom,
            ["UnitsRequested"]           = product.UnitsRequested,
            ["UnitsQuoted"]              = product.UnitsQuoted,
            ["LengthPerUnit"]            = product.LengthPerUnit,
            ["LengthUnit"]               = product.LengthUnit,
            ["WeightPerUnit"]            = product.WeightPerUnit,
            ["WeightUnit"]               = product.WeightUnit,
            ["PricePerPound"]            = product.PricePerPound,
            ["PricePerFoot"]             = product.PricePerFoot,
            ["PricePerPiece"]            = product.PricePerPiece,
            ["TotalPrice"]               = product.TotalPrice,
            ["LeadTimeText"]             = product.LeadTimeText,
            ["Certifications"]           = product.Certifications,
            ["SupplierProductComments"]  = product.SupplierProductComments,
            ["IsRegret"]                 = HasRegretPhrase(product.SupplierProductComments),
        };

        var existingId = await FindExistingSupplierLineItemAsync(
            siteId, listId, supplierResponseId, prodName, prodTokens);

        if (existingId is not null)
        {
            var update = new Dictionary<string, object?>(fieldData);
            update.Remove("ProductName"); // preserve canonical name from first write
            await GetGraph().Sites[siteId].Lists[listId].Items[existingId].Fields
                .PatchAsync(new FieldValueSet { AdditionalData = update });
            _log.LogInformation("[SP] Updated SupplierLineItem {Id} for '{Name}'", existingId, prodName);
        }
        else
        {
            var item = await GetGraph().Sites[siteId].Lists[listId].Items
                .PostAsync(new ListItem { Fields = new FieldValueSet { AdditionalData = fieldData } });
            _log.LogInformation("[SP] Created SupplierLineItem {Id} for '{Name}'", item!.Id, prodName);
        }
    }

    private async Task<string?> FindExistingSupplierLineItemAsync(
        string siteId, string listId,
        string supplierResponseId, string productName, HashSet<string> productTokens)
    {
        var result = await GetGraph().Sites[siteId].Lists[listId].Items
            .GetAsync(r =>
            {
                r.QueryParameters.Filter = $"fields/SupplierResponseId eq '{EscapeOdata(supplierResponseId)}'";
                r.QueryParameters.Expand = ["fields($select=id,SupplierResponseId,ProductName)"];
                r.QueryParameters.Top    = 100;
                r.Headers.Add("Prefer", "HonorNonIndexedQueriesWarningMayFailRandomly");
            });

        var match = result?.Value?.FirstOrDefault(i =>
        {
            var d = i.Fields?.AdditionalData;
            if (d is null) return false;
            var spProduct = d.TryGetValue("ProductName", out var p) ? p?.ToString() : null;
            if (NormalizeMatch(spProduct, productName)) return true;
            var spTokens = ProductTokens(spProduct ?? string.Empty);
            return NumericTokensCompatible(productTokens, spTokens)
                && ProductJaccard(spTokens, productTokens) >= 0.5;
        });

        return match?.Id;
    }

    // ── Regret detection ─────────────────────────────────────────────────────

    private static bool HasRegretPhrase(string? text) =>
        text is not null &&
        _regretPhrases.Any(p => text.Contains(p, StringComparison.OrdinalIgnoreCase));

    // ── OData helpers ────────────────────────────────────────────────────────

    private static string EscapeOdata(string s) => s.Replace("'", "''");

    // ── Product tokenisation ─────────────────────────────────────────────────

    private static readonly Regex _normaliseRegex  = new(@"[\s\W]+", RegexOptions.Compiled);
    private static bool NormalizeMatch(string? a, string? b)
    {
        if (a is null && b is null) return true;
        if (a is null || b is null) return false;
        static string N(string s) => _normaliseRegex.Replace(s.Trim().ToLowerInvariant(), " ").Trim();
        return N(a) == N(b);
    }

    private static readonly Regex _dimFraction  = new(@"(\d+)/(\d+)",                               RegexOptions.Compiled);
    private static readonly Regex _dimDecimal   = new(@"(\d+)\.(\d+)",                              RegexOptions.Compiled);
    private static readonly Regex _dimSeparator = new(@"(\d[a-z0-9]*)[""']?\s*[xX×]\s*[""']?(\d[a-z0-9]*)", RegexOptions.Compiled);
    private static readonly Regex _dimSplit     = new(@"[^a-z0-9]+",                                RegexOptions.Compiled);
    private static readonly Regex _orLength     = new(@"\bor\s+\d+[a-z""']*\b", RegexOptions.Compiled | RegexOptions.IgnoreCase);

    private static string PreprocessProduct(string s)
    {
        s = s.ToLowerInvariant();
        s = _orLength.Replace(s, "");
        s = Regex.Replace(s, @"\brandom\s+lengths?\b|\bmill\s+lengths?\b|\bfull\s+lengths?\b|\blengths?\b", "");
        s = _dimFraction.Replace(s, "$1f$2");
        s = _dimDecimal.Replace(s, "$1d$2");
        s = Regex.Replace(s, @"d(\d+)", m =>
        {
            var stripped = m.Groups[1].Value.TrimEnd('0');
            return "d" + (stripped.Length == 0 ? "0" : stripped);
        });
        s = _dimSeparator.Replace(s, "$1x$2");
        s = _dimSeparator.Replace(s, "$1x$2");
        s = Regex.Replace(s, @"[""']", "");
        return s;
    }

    private static HashSet<string> ProductTokens(string s)
    {
        var p = PreprocessProduct(s);
        return _dimSplit.Split(p)
                        .Where(t => t.Length > 1 || (t.Length == 1 && char.IsDigit(t[0])))
                        .ToHashSet();
    }

    private static double ProductJaccard(HashSet<string> a, HashSet<string> b)
    {
        if (a.Count == 0 && b.Count == 0) return 1.0;
        var intersection = a.Count(t => b.Contains(t));
        var union        = a.Count + b.Count - intersection;
        return union == 0 ? 0 : (double)intersection / union;
    }

    private static bool HasDigit(string t) => t.Any(char.IsDigit);
    private static bool IsDimToken(string t) =>
        t.Any(char.IsDigit) && t.Any(c => c == 'x' || c == 'f' || c == 'd');

    private static bool NumericTokensCompatible(HashSet<string> a, HashSet<string> b)
    {
        var numA = a.Where(HasDigit).ToHashSet();
        var numB = b.Where(HasDigit).ToHashSet();
        var dimA = numA.Where(IsDimToken).ToHashSet();
        var dimB = numB.Where(IsDimToken).ToHashSet();

        if (dimA.Count > 0 && dimB.Count > 0)
        {
            if (!dimA.SetEquals(dimB)) return false;
            var gradeA = numA.Where(t => !IsDimToken(t)).ToHashSet();
            var gradeB = numB.Where(t => !IsDimToken(t)).ToHashSet();
            return gradeA.IsSubsetOf(gradeB) || gradeB.IsSubsetOf(gradeA);
        }
        if (dimA.Count > 0) return false;

        var gA = numA.Where(t => !IsDimToken(t)).ToHashSet();
        var gB = numB.Where(t => !IsDimToken(t)).ToHashSet();
        return gA.IsSubsetOf(gB) || gB.IsSubsetOf(gA);
    }

    // ── Attachment upload (SharePoint REST API) ──────────────────────────────

    private async Task UpsertItemAttachmentAsync(string spItemId, string listId, string fileName, byte[] bytes)
    {
        var siteUrl  = _config["SharePoint:SiteUrl"] ?? "https://metalsupermarkets-my.sharepoint.com/personal/angus_mithrilmetals_com";
        var uri      = new Uri(siteUrl);
        var host     = uri.Host;
        var sitePath = uri.AbsolutePath.TrimEnd('/');

        var tokenCtx = new Azure.Core.TokenRequestContext([$"https://{host}/.default"]);
        var token    = await GetSpCredential().GetTokenAsync(tokenCtx);

        var attBase = $"https://{host}{sitePath}/_api/web/lists(guid'{listId}')/items({spItemId})/AttachmentFiles";

        using var http = new HttpClient();
        http.DefaultRequestHeaders.Authorization =
            new System.Net.Http.Headers.AuthenticationHeaderValue("Bearer", token.Token);
        http.DefaultRequestHeaders.Add("Accept", "application/json;odata=nometadata");

        // Delete existing same-named attachment if present
        var listResp = await http.GetAsync(attBase);
        if (listResp.IsSuccessStatusCode)
        {
            var listJson = await listResp.Content.ReadAsStringAsync();
            using var listDoc = System.Text.Json.JsonDocument.Parse(listJson);
            var alreadyExists = listDoc.RootElement.TryGetProperty("value", out var val) &&
                val.EnumerateArray()
                   .Any(e => e.TryGetProperty("FileName", out var fn) &&
                             string.Equals(fn.GetString(), fileName, StringComparison.OrdinalIgnoreCase));

            if (alreadyExists)
            {
                var delUrl = $"{attBase}/getByFileName('{Uri.EscapeDataString(fileName)}')";
                using var delReq = new HttpRequestMessage(HttpMethod.Delete, delUrl);
                delReq.Headers.Add("IF-MATCH", "*");
                var delResp = await http.SendAsync(delReq);
                if (!delResp.IsSuccessStatusCode)
                    _log.LogWarning("[SP] Could not delete existing attachment '{File}': {Status}", fileName, delResp.StatusCode);
            }
        }

        // Upload
        var uploadUrl  = $"{attBase}/add(FileName='{Uri.EscapeDataString(fileName)}')";
        var fileContent = new ByteArrayContent(bytes);
        fileContent.Headers.ContentType =
            new System.Net.Http.Headers.MediaTypeHeaderValue("application/octet-stream");

        var addResp = await http.PostAsync(uploadUrl, fileContent);
        if (addResp.IsSuccessStatusCode)
            _log.LogInformation("[SP] Uploaded attachment '{File}' ({Bytes} bytes) to item {Id}", fileName, bytes.Length, spItemId);
        else
        {
            var err = await addResp.Content.ReadAsStringAsync();
            _log.LogWarning("[SP] Failed to upload attachment '{File}': {Status} {Body}",
                fileName, addResp.StatusCode, err[..Math.Min(err.Length, 400)]);
        }
    }

    // ── Clean: delete all derived email-processing data ──────────────────────

    /// <summary>
    /// Deletes every item in SupplierResponses and SupplierLineItems.
    /// Returns counts of items deleted from each list.
    /// Does NOT touch RFQ References (notes / dates).
    /// </summary>
    public async Task<(int SrDeleted, int SliDeleted)> CleanSupplierDataAsync()
    {
        var siteId    = await GetSiteIdAsync();
        var srListId  = await GetSupplierResponsesListIdAsync();
        var sliListId = await GetSupplierLineItemsListIdAsync();

        var srDeleted  = await DeleteAllItemsAsync(siteId, srListId,  "SupplierResponses");
        var sliDeleted = await DeleteAllItemsAsync(siteId, sliListId, "SupplierLineItems");

        return (srDeleted, sliDeleted);
    }

    private async Task<int> DeleteAllItemsAsync(string siteId, string listId, string listName)
    {
        int deleted = 0;

        while (true)
        {
            // Fetch a page of item IDs only — no fields needed
            var page = await GetGraph().Sites[siteId].Lists[listId].Items
                .GetAsync(req => { req.QueryParameters.Top = 100; });

            var items = page?.Value;
            if (items is null || items.Count == 0) break;

            // Delete in parallel (Graph throttles at ~4 concurrent writes; keep modest)
            var tasks = items
                .Where(i => i.Id is not null)
                .Select(i => GetGraph().Sites[siteId].Lists[listId].Items[i.Id!].DeleteAsync());

            await Task.WhenAll(tasks);
            deleted += items.Count;
            _log.LogInformation("[SP] Deleted {Count} items from {List} (total so far: {Total})",
                items.Count, listName, deleted);

            // If fewer items than page size were returned, we're done
            if (items.Count < 100) break;
        }

        _log.LogInformation("[SP] Finished cleaning {List}: {Total} items deleted", listName, deleted);
        return deleted;
    }

    // ── Fetch SP list item attachment (SupplierResponses PDF) ────────────────

    /// <summary>
    /// Downloads the named attachment from a SupplierResponses list item via the SP REST API.
    /// Returns null if the item or file is not found.
    /// </summary>
    public async Task<(string ContentType, byte[] Bytes, string FileName)?> GetSpItemAttachmentAsync(
        string srItemId, string fileName)
    {
        var siteUrl  = _config["SharePoint:SiteUrl"] ?? "https://metalsupermarkets-my.sharepoint.com/personal/angus_mithrilmetals_com";
        var uri      = new Uri(siteUrl);
        var host     = uri.Host;
        var sitePath = uri.AbsolutePath.TrimEnd('/');
        var listId   = await GetSupplierResponsesListIdAsync();

        var tokenCtx = new Azure.Core.TokenRequestContext([$"https://{host}/.default"]);
        var token    = await GetSpCredential().GetTokenAsync(tokenCtx);

        var url = $"https://{host}{sitePath}/_api/web/lists(guid'{listId}')" +
                  $"/items({srItemId})/AttachmentFiles" +
                  $"/getByFileName('{Uri.EscapeDataString(fileName)}')/$value";

        using var http = new HttpClient();
        http.DefaultRequestHeaders.Authorization =
            new System.Net.Http.Headers.AuthenticationHeaderValue("Bearer", token.Token);

        var resp = await http.GetAsync(url);
        if (!resp.IsSuccessStatusCode)
        {
            _log.LogWarning("[SP] Attachment not found: srItemId={Id} file={File} status={Status}",
                srItemId, fileName, resp.StatusCode);
            return null;
        }

        var bytes = await resp.Content.ReadAsByteArrayAsync();
        var ct    = resp.Content.Headers.ContentType?.MediaType ?? "application/octet-stream";
        _log.LogInformation("[SP] Fetched attachment '{File}' ({Bytes} bytes) from SR item {Id}", fileName, bytes.Length, srItemId);
        return (ct, bytes, fileName);
    }

    // ── Diagnostics ──────────────────────────────────────────────────────────

    public async Task<object> DiagnoseAsync()
    {
        var steps = new List<object>();
        try
        {
            steps.Add(new { step = "token", status = "trying" });
            var tenantId     = _config["SharePoint:TenantId"]     ?? throw new Exception("SharePoint:TenantId not set");
            var clientId     = _config["SharePoint:ClientId"]     ?? throw new Exception("SharePoint:ClientId not set");
            var clientSecret = _config["SharePoint:ClientSecret"] ?? throw new Exception("SharePoint:ClientSecret not set");
            var credential   = new Azure.Identity.ClientSecretCredential(tenantId, clientId, clientSecret);
            var tokenCtx     = new Azure.Core.TokenRequestContext(["https://graph.microsoft.com/.default"]);
            var token        = await credential.GetTokenAsync(tokenCtx);
            var jwtParts     = token.Token.Split('.');
            var claimsJson   = jwtParts.Length > 1
                ? System.Text.Encoding.UTF8.GetString(
                    Convert.FromBase64String(jwtParts[1].PadRight((jwtParts[1].Length + 3) & ~3, '=')))
                : "{}";
            using var claimsDoc = System.Text.Json.JsonDocument.Parse(claimsJson);
            var roles = claimsDoc.RootElement.TryGetProperty("roles", out var r) ? r.ToString() : "NONE";
            var aud   = claimsDoc.RootElement.TryGetProperty("aud",   out var a) ? a.ToString() : "?";
            var tid   = claimsDoc.RootElement.TryGetProperty("tid",   out var t) ? t.ToString() : "?";
            steps[^1] = new { step = "token", status = "ok", expiresOn = token.ExpiresOn, aud, tid, roles };

            var graph = GetGraph();

            steps.Add(new { step = "sites/root", status = "trying" });
            var root = await graph.Sites["root"].GetAsync();
            steps[^1] = new { step = "sites/root", status = "ok", siteId = root?.Id, webUrl = root?.WebUrl };

            var siteUrl = _config["SharePoint:SiteUrl"] ?? "https://metalsupermarkets-my.sharepoint.com/personal/angus_mithrilmetals_com";
            var uri     = new Uri(siteUrl);
            var siteKey = $"{uri.Host}:{uri.AbsolutePath}";
            steps.Add(new { step = $"sites/{siteKey}", status = "trying" });
            var site = await graph.Sites[siteKey].GetAsync();
            steps[^1] = new { step = $"sites/{siteKey}", status = "ok", siteId = site?.Id };

            foreach (var listName in new[] { "SupplierResponses", "SupplierLineItems", "RFQ References" })
            {
                steps.Add(new { step = "list lookup", status = "trying" });
                var lists = await graph.Sites[site!.Id!].Lists
                    .GetAsync(r => r.QueryParameters.Filter = $"displayName eq '{listName}'");
                var found = lists?.Value?.FirstOrDefault();
                steps[^1] = new { step = "list lookup",
                                  status = found != null ? "ok" : "not_found",
                                  listId = found?.Id, listName };
            }
        }
        catch (Microsoft.Graph.Models.ODataErrors.ODataError ex)
        {
            steps.Add(new { step = "error", code = ex.Error?.Code, message = ex.Error?.Message });
        }
        catch (Exception ex)
        {
            steps.Add(new { step = "error", message = ex.Message });
        }
        return new { steps };
    }

    // ── Provision new supplier lists (run once) ──────────────────────────────

    public async Task<Dictionary<string, object>> EnsureSupplierListsAsync()
    {
        var siteId  = await GetSiteIdAsync();
        var results = new Dictionary<string, object>
        {
            ["SupplierResponses"]  = await EnsureListColumnsAsync(siteId, "SupplierResponses",
            [
                ("RFQ_ID",               "text"),
                ("SupplierName",         "text"),
                ("EmailFrom",            "text"),
                ("ReceivedAt",           "dateTime"),
                ("EmailSubject",         "text"),
                ("EmailBody",            "note"),
                ("ProcessedAt",          "dateTime"),
                ("ProcessingSource",     "text"),
                ("SourceFile",           "text"),
                ("QuoteReference",       "text"),
                ("DateOfQuote",          "dateTime"),
                ("EstimatedDeliveryDate","dateTime"),
                ("FreightTerms",         "text"),
                ("IsRegret",             "boolean"),
            ]),
            ["SupplierLineItems"] = await EnsureListColumnsAsync(siteId, "SupplierLineItems",
            [
                ("SupplierResponseId",       "text"),
                ("RFQ_ID",                   "text"),
                ("SupplierName",             "text"),
                ("SourceFile",               "text"),
                ("EmailFrom",                "text"),
                ("ProductName",              "text"),
                ("CatalogProductName",       "text"),
                ("ProductSearchKey",         "text"),
                ("UnitsRequested",           "number"),
                ("UnitsQuoted",              "number"),
                ("LengthPerUnit",            "number"),
                ("LengthUnit",               "text"),
                ("WeightPerUnit",            "number"),
                ("WeightUnit",               "text"),
                ("PricePerPound",            "number"),
                ("PricePerFoot",             "number"),
                ("PricePerPiece",            "number"),
                ("TotalPrice",               "number"),
                ("LeadTimeText",             "text"),
                ("Certifications",           "text"),
                ("SupplierProductComments",  "note"),
                ("IsRegret",                 "boolean"),
            ]),
        };
        return results;
    }

    private async Task<Dictionary<string, string>> EnsureListColumnsAsync(
        string siteId, string listName, (string Name, string Type)[] columns)
    {
        // Create list if absent
        var lists = await GetGraph().Sites[siteId].Lists
            .GetAsync(r => r.QueryParameters.Filter = $"displayName eq '{listName}'");
        string listId;
        if (lists?.Value?.FirstOrDefault() is null)
        {
            var newList = await GetGraph().Sites[siteId].Lists.PostAsync(new List
            {
                DisplayName = listName,
                ListProp    = new ListInfo { Template = "genericList" },
            });
            listId = newList?.Id ?? throw new Exception($"Failed to create list '{listName}'");
            _log.LogInformation("[SP] Created list '{Name}' -> id: {Id}", listName, listId);
        }
        else
        {
            listId = lists.Value.First().Id!;
        }

        // Cache the newly-resolved IDs
        if (listName == "SupplierResponses") _srListId  = listId;
        if (listName == "SupplierLineItems") _sliListId = listId;

        var existing = await GetGraph().Sites[siteId].Lists[listId].Columns.GetAsync();
        var existingNames = existing?.Value?
            .Select(c => c.Name ?? "").ToHashSet(StringComparer.OrdinalIgnoreCase) ?? [];

        var results = new Dictionary<string, string>();
        foreach (var (name, type) in columns)
        {
            if (existingNames.Contains(name)) { results[name] = "exists"; continue; }
            try
            {
                var col = type switch
                {
                    "text"     => new ColumnDefinition { Name = name, Text     = new TextColumn() },
                    "number"   => new ColumnDefinition { Name = name, Number   = new NumberColumn() },
                    "dateTime" => new ColumnDefinition { Name = name, DateTime = new DateTimeColumn() },
                    "note"     => new ColumnDefinition { Name = name, Text     = new TextColumn { AllowMultipleLines = true, LinesForEditing = 6 } },
                    "boolean"  => new ColumnDefinition { Name = name, Boolean  = new BooleanColumn() },
                    _          => new ColumnDefinition { Name = name, Text     = new TextColumn() }
                };
                await GetGraph().Sites[siteId].Lists[listId].Columns.PostAsync(col);
                results[name] = "created";
                _log.LogInformation("[SP] Created column '{Name}' ({Type}) on '{List}'", name, type, listName);
            }
            catch (Exception ex)
            {
                results[name] = $"error: {ex.Message}";
                _log.LogWarning("[SP] Column '{Name}' on '{List}': {Err}", name, listName, ex.Message);
            }
        }
        return results;
    }

    // ── Legacy: provision old RFQ Line Items list (kept for recovery) ────────

    public async Task<Dictionary<string, string>> EnsureColumnsAsync()
    {
        var siteId  = await GetSiteIdAsync();
        var listName = _config["SharePoint:ListName"] ?? "RFQ Line Items";

        var lists = await GetGraph().Sites[siteId].Lists
            .GetAsync(r => r.QueryParameters.Filter = $"displayName eq '{listName}'");
        if (lists?.Value?.FirstOrDefault() is null)
        {
            var newList = await GetGraph().Sites[siteId].Lists.PostAsync(new List
            {
                DisplayName = listName,
                ListProp    = new ListInfo { Template = "genericList" },
            });
            _listId = newList?.Id ?? throw new Exception($"Failed to create list '{listName}'");
        }

        var listId = await GetListIdAsync();
        var results = new Dictionary<string, string>();
        var existing = await GetGraph().Sites[siteId].Lists[listId].Columns.GetAsync();
        var existingNames = existing?.Value?.Select(c => c.Name ?? "").ToHashSet(StringComparer.OrdinalIgnoreCase) ?? [];

        var columns = new (string Name, string Type)[]
        {
            ("JobReference",            "text"),
            ("EmailFrom",               "text"),
            ("ReceivedAt",              "dateTime"),
            ("ProcessedAt",             "dateTime"),
            ("ProcessingSource",        "text"),
            ("SourceFile",              "text"),
            ("SupplierName",            "text"),
            ("QuoteReference",          "text"),
            ("DateOfQuote",             "dateTime"),
            ("EstimatedDeliveryDate",   "dateTime"),
            ("ProductName",             "text"),
            ("UnitsRequested",          "number"),
            ("UnitsQuoted",             "number"),
            ("LengthPerUnit",           "number"),
            ("LengthUnit",              "text"),
            ("WeightPerUnit",           "number"),
            ("WeightUnit",              "text"),
            ("PricePerPound",           "number"),
            ("PricePerFoot",            "number"),
            ("PricePerPiece",           "number"),
            ("TotalPrice",              "number"),
            ("LeadTimeText",            "text"),
            ("Certifications",          "text"),
            ("FreightTerms",            "text"),
            ("SupplierProductComments", "note"),
            ("CatalogProductName",      "text"),
            ("ProductSearchKey",        "text"),
            ("EmailBody",               "note"),
        };

        foreach (var (name, type) in columns)
        {
            if (existingNames.Contains(name)) { results[name] = "exists"; continue; }
            try
            {
                var col = type switch
                {
                    "text"     => new ColumnDefinition { Name = name, Text     = new TextColumn() },
                    "number"   => new ColumnDefinition { Name = name, Number   = new NumberColumn() },
                    "dateTime" => new ColumnDefinition { Name = name, DateTime = new DateTimeColumn() },
                    "note"     => new ColumnDefinition { Name = name, Text     = new TextColumn { AllowMultipleLines = true, LinesForEditing = 6 } },
                    _          => new ColumnDefinition { Name = name, Text     = new TextColumn() }
                };
                await GetGraph().Sites[siteId].Lists[listId].Columns.PostAsync(col);
                results[name] = "created";
            }
            catch (Exception ex) { results[name] = $"error: {ex.Message}"; }
        }
        return results;
    }

    // ── Legacy: read from old RFQ Line Items (kept for migration) ───────────

    public async Task<List<Dictionary<string, object?>>> ReadItemsAsync(int top = 500)
    {
        var siteId = await GetSiteIdAsync();
        var listId = await GetListIdAsync();

        var result = await GetGraph().Sites[siteId].Lists[listId].Items
            .GetAsync(req => { req.QueryParameters.Expand = ["fields"]; req.QueryParameters.Top = top; });

        static bool IsAppField(string key) =>
            !key.StartsWith('@') && !key.StartsWith('_') &&
            key is not ("LinkTitle" or "LinkTitleNoMenu" or "ContentType" or "Edit"
                     or "Attachments" or "ItemChildCount" or "FolderChildCount"
                     or "Modified" or "Created"
                     or "AuthorLookupId" or "EditorLookupId"
                     or "AppAuthorLookupId" or "AppEditorLookupId");

        return result?.Value?
            .Where(i => i.Fields?.AdditionalData is not null)
            .Select(i => i.Fields!.AdditionalData!
                .Where(kv => IsAppField(kv.Key))
                .ToDictionary(kv => kv.Key, kv => kv.Value))
            .ToList()
            ?? [];
    }
}
