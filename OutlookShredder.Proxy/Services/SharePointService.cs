using Azure.Identity;
using Microsoft.Graph;
using Microsoft.Graph.Models;
using OutlookShredder.Proxy.Models;

namespace OutlookShredder.Proxy.Services;

/// <summary>
/// Writes extracted RFQ product lines to the SharePoint RFQLineItems list
/// using Microsoft Graph API with client-credential (app-only) authentication.
///
/// Azure AD app requires:  Sites.ReadWrite.All  (Application permission, admin consented)
/// </summary>
public class SharePointService
{
    private readonly IConfiguration _config;
    private readonly ILogger<SharePointService> _log;

    private GraphServiceClient? _graph;
    private string? _siteId;
    private string? _listId;

    public SharePointService(IConfiguration config, ILogger<SharePointService> log)
    {
        _config = config;
        _log    = log;
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

    // ── Resolve site ID (cached) ─────────────────────────────────────────────
    private async Task<string> GetSiteIdAsync()
    {
        if (_siteId is not null) return _siteId;

        var siteUrl = _config["SharePoint:SiteUrl"]
            ?? "https://metalsupermarkets.sharepoint.com/sites/hackensack";

        var uri  = new Uri(siteUrl);
        var host = uri.Host;
        var path = uri.AbsolutePath;

        _log.LogInformation("[SP] Resolving site: {Host}{Path}", host, path);

        var site = await GetGraph().Sites[$"{host}:{path}"].GetAsync();
        _siteId  = site!.Id ?? throw new Exception("Could not resolve SharePoint site ID");

        _log.LogInformation("[SP] Site ID: {Id}", _siteId);
        return _siteId;
    }

    // ── Resolve list ID (cached) ─────────────────────────────────────────────
    private async Task<string> GetListIdAsync()
    {
        if (_listId is not null) return _listId;

        var siteId   = await GetSiteIdAsync();
        var listName = _config["SharePoint:ListName"] ?? "RFQLineItems";

        var lists = await GetGraph().Sites[siteId].Lists
            .GetAsync(req => req.QueryParameters.Filter = $"displayName eq '{listName}'");

        var list  = lists?.Value?.FirstOrDefault()
            ?? throw new Exception($"SharePoint list '{listName}' not found. Run POST /api/setup-columns first.");

        _listId = list.Id ?? throw new Exception("List ID was null");
        _log.LogInformation("[SP] List '{Name}' -> id: {Id}", listName, _listId);
        return _listId;
    }

    // ── Read all items (for dashboard) ───────────────────────────────────────
    /// <summary>
    /// Returns up to <paramref name="top"/> items from the RFQLineItems list as
    /// plain field dictionaries, suitable for JSON serialisation to the dashboard.
    /// Uses the same app-only credentials as writes — no browser token required.
    /// </summary>
    public async Task<List<Dictionary<string, object?>>> ReadItemsAsync(int top = 500)
    {
        var siteId = await GetSiteIdAsync();
        var listId = await GetListIdAsync();

        var result = await GetGraph().Sites[siteId].Lists[listId].Items
            .GetAsync(req =>
            {
                req.QueryParameters.Expand = ["fields"];
                req.QueryParameters.Top    = top;
            });

        // Strip SharePoint system keys (@odata.*, LinkTitle*, ContentType, Lookup IDs, etc.)
        // so the dashboard only receives the columns we care about.
        static bool IsAppField(string key) =>
            !key.StartsWith('@') &&
            !key.StartsWith('_') &&
            key is not ("LinkTitle" or "LinkTitleNoMenu" or "ContentType"
                     or "Edit" or "Attachments" or "ItemChildCount" or "FolderChildCount"
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

    // ── Write one product row ────────────────────────────────────────────────
    public async Task<SpWriteResult> WriteProductRowAsync(
        RfqExtraction header,
        ProductLine   product,
        ExtractRequest emailMeta,
        string        source,
        string?       sourceFile,
        int           rowIndex)
    {
        var result = new SpWriteResult { ProductName = product.ProductName };
        try
        {
            var siteId  = await GetSiteIdAsync();
            var listId  = await GetListIdAsync();
            var jobRef  = (header.JobReference ?? emailMeta.JobRefs.FirstOrDefault()?.Trim('[', ']') ?? "UNKNOWN").ToUpperInvariant();
            var supplier = header.SupplierName ?? emailMeta.EmailFrom ?? "Unknown";
            var prodName = product.ProductName ?? $"Product {rowIndex + 1}";

            var fields = new FieldValueSet
            {
                AdditionalData = new Dictionary<string, object?>
                {
                    ["Title"]                    = $"[{jobRef}] {supplier} - {prodName}"[..Math.Min($"[{jobRef}] {supplier} - {prodName}".Length, 255)],
                    ["JobReference"]             = jobRef,
                    ["EmailFrom"]                = emailMeta.EmailFrom,
                    ["ReceivedAt"]               = emailMeta.ReceivedAt,
                    ["ProcessedAt"]              = DateTime.UtcNow.ToString("o"),
                    ["ProcessingSource"]         = source,
                    ["SourceFile"]               = sourceFile,
                    ["SupplierName"]             = supplier,
                    ["DateOfQuote"]              = header.DateOfQuote,
                    ["EstimatedDeliveryDate"]    = header.EstimatedDeliveryDate,
                    ["ProductName"]              = prodName,
                    ["UnitsRequested"]           = product.UnitsRequested,
                    ["UnitsQuoted"]              = product.UnitsQuoted,
                    ["LengthPerUnit"]            = product.LengthPerUnit,
                    ["LengthUnit"]               = product.LengthUnit,
                    ["WeightPerUnit"]            = product.WeightPerUnit,
                    ["WeightUnit"]               = product.WeightUnit,
                    ["PricePerPound"]            = product.PricePerPound,
                    ["PricePerFoot"]             = product.PricePerFoot,
                    ["SupplierProductComments"]  = product.SupplierProductComments,
                }
            };

            var item = await GetGraph().Sites[siteId].Lists[listId].Items
                .PostAsync(new ListItem { Fields = fields });

            result.Success  = true;
            result.SpItemId = item!.Id;
            result.SpWebUrl = item.WebUrl;
            _log.LogInformation("[SP] Wrote row for product '{Name}' -> item {Id}", prodName, item.Id);
        }
        catch (Microsoft.Graph.Models.ODataErrors.ODataError odataEx)
        {
            result.Success = false;
            result.Error   = odataEx.Message;
            _log.LogError("[SP] ODataError writing '{Name}': code={Code} message={Msg} inner={Inner}",
                product.ProductName,
                odataEx.Error?.Code,
                odataEx.Error?.Message,
                odataEx.Error?.InnerError?.AdditionalData != null
                    ? string.Join(", ", odataEx.Error.InnerError.AdditionalData.Select(k => $"{k.Key}={k.Value}"))
                    : "none");
        }
        catch (Exception ex)
        {
            result.Success = false;
            result.Error   = ex.Message;
            _log.LogError(ex, "[SP] Failed to write row for product '{Name}'", product.ProductName);
        }
        return result;
    }

    // ── Diagnostics ──────────────────────────────────────────────────────────
    public async Task<object> DiagnoseAsync()
    {
        var steps = new List<object>();
        try
        {
            // Step 1: explicitly acquire a token so we know credentials are valid
            steps.Add(new { step = "token", status = "trying" });
            var tenantId     = _config["SharePoint:TenantId"]     ?? throw new Exception("SharePoint:TenantId not set");
            var clientId     = _config["SharePoint:ClientId"]     ?? throw new Exception("SharePoint:ClientId not set");
            var clientSecret = _config["SharePoint:ClientSecret"] ?? throw new Exception("SharePoint:ClientSecret not set");
            var credential   = new Azure.Identity.ClientSecretCredential(tenantId, clientId, clientSecret);
            var tokenCtx     = new Azure.Core.TokenRequestContext(["https://graph.microsoft.com/.default"]);
            var token        = await credential.GetTokenAsync(tokenCtx);
            // Decode token claims (middle JWT segment) without validating signature
            var jwtParts  = token.Token.Split('.');
            var claimsJson = jwtParts.Length > 1
                ? System.Text.Encoding.UTF8.GetString(
                    Convert.FromBase64String(jwtParts[1].PadRight((jwtParts[1].Length + 3) & ~3, '=')))
                : "{}";
            using var claimsDoc = System.Text.Json.JsonDocument.Parse(claimsJson);
            var roles = claimsDoc.RootElement.TryGetProperty("roles", out var r) ? r.ToString() : "NONE";
            var aud   = claimsDoc.RootElement.TryGetProperty("aud",   out var a) ? a.ToString() : "?";
            var tid   = claimsDoc.RootElement.TryGetProperty("tid",   out var t) ? t.ToString() : "?";
            steps[^1] = new { step = "token", status = "ok", expiresOn = token.ExpiresOn, aud, tid, roles };
            var graph = GetGraph();

            // Step 2: raw HTTP call to /sites/root so we see the exact response
            steps.Add(new { step = "sites/root (raw)", status = "trying" });
            using var http = new System.Net.Http.HttpClient();
            http.DefaultRequestHeaders.Authorization =
                new System.Net.Http.Headers.AuthenticationHeaderValue("Bearer", token.Token);
            var rawResp = await http.GetAsync("https://graph.microsoft.com/v1.0/sites/root");
            var rawBody = await rawResp.Content.ReadAsStringAsync();
            steps[^1] = new { step = "sites/root (raw)", status = ((int)rawResp.StatusCode).ToString(), body = rawBody };

            if (!rawResp.IsSuccessStatusCode) return new { steps };

            // Step 3: can we reach Graph /sites/root via SDK?
            steps.Add(new { step = "sites/root", status = "trying" });
            var root = await graph.Sites["root"].GetAsync();
            steps[^1] = new { step = "sites/root", status = "ok", siteId = root?.Id, webUrl = root?.WebUrl };

            // Step 4: can we resolve the configured site by host:path?
            var siteUrl  = _config["SharePoint:SiteUrl"] ?? "https://metalsupermarkets.sharepoint.com/sites/hackensack";
            var uri      = new Uri(siteUrl);
            var siteKey  = $"{uri.Host}:{uri.AbsolutePath}";
            steps.Add(new { step = $"sites/{siteKey}", status = "trying" });
            var site = await graph.Sites[siteKey].GetAsync();
            steps[^1] = new { step = $"sites/{siteKey}", status = "ok", siteId = site?.Id };

            // Step 5: can we list lists on that site?
            steps.Add(new { step = "list lookup", status = "trying" });
            var listName = _config["SharePoint:ListName"] ?? "RFQLineItems";
            var lists    = await graph.Sites[site!.Id!].Lists
                .GetAsync(r => r.QueryParameters.Filter = $"displayName eq '{listName}'");
            var found = lists?.Value?.FirstOrDefault();
            steps[^1] = new { step = "list lookup", status = found != null ? "ok" : "not_found",
                              listId = found?.Id, listName };
        }
        catch (Microsoft.Graph.Models.ODataErrors.ODataError ex)
        {
            steps.Add(new { step = "error", code = ex.Error?.Code, message = ex.Error?.Message,
                            inner = ex.Error?.InnerError?.AdditionalData?
                                .Select(k => $"{k.Key}={k.Value}") });
        }
        catch (Exception ex)
        {
            steps.Add(new { step = "error", message = ex.Message });
        }
        return new { steps };
    }

    // ── Provision columns (run once) ─────────────────────────────────────────
    public async Task<Dictionary<string, string>> EnsureColumnsAsync()
    {
        var siteId = await GetSiteIdAsync();
        var listId = await GetListIdAsync();
        var results = new Dictionary<string, string>();

        var existing = await GetGraph().Sites[siteId].Lists[listId].Columns.GetAsync();
        var existingNames = existing?.Value?.Select(c => c.Name ?? "").ToHashSet(StringComparer.OrdinalIgnoreCase)
            ?? [];

        var columns = new (string Name, string Type)[]
        {
            ("JobReference",            "text"),
            ("EmailFrom",               "text"),
            ("ReceivedAt",              "dateTime"),
            ("ProcessedAt",             "dateTime"),
            ("ProcessingSource",        "text"),
            ("SourceFile",              "text"),
            ("SupplierName",            "text"),
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
            ("SupplierProductComments", "note"),
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
                _log.LogInformation("[SP] Created column '{Name}' ({Type})", name, type);
            }
            catch (Exception ex)
            {
                results[name] = $"error: {ex.Message}";
            }
        }
        return results;
    }
}
