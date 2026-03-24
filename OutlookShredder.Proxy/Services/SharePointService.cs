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
                    ["Title"]                    = $"[{jobRef}] {supplier} – {prodName}"[..Math.Min($"[{jobRef}] {supplier} – {prodName}".Length, 255)],
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
        catch (Exception ex)
        {
            result.Success = false;
            result.Error   = ex.Message;
            _log.LogError(ex, "[SP] Failed to write row for product '{Name}'", product.ProductName);
        }
        return result;
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
