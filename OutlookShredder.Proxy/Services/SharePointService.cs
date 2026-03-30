using System.Text.RegularExpressions;
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
    private readonly SupplierCacheService  _suppliers;
    private readonly ProductCatalogService _catalog;

    private GraphServiceClient?       _graph;
    private ClientSecretCredential?   _spCredential;
    private string? _siteId;
    private string? _listId;

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

    // ── SharePoint REST credential (lazy init, separate audience from Graph) ─
    private ClientSecretCredential GetSpCredential()
    {
        if (_spCredential is not null) return _spCredential;

        var tenantId     = _config["SharePoint:TenantId"]     ?? throw new InvalidOperationException("SharePoint:TenantId not set");
        var clientId     = _config["SharePoint:ClientId"]     ?? throw new InvalidOperationException("SharePoint:ClientId not set");
        var clientSecret = _config["SharePoint:ClientSecret"] ?? throw new InvalidOperationException("SharePoint:ClientSecret not set");

        _spCredential = new ClientSecretCredential(tenantId, clientId, clientSecret);
        return _spCredential;
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

    // ── Upsert one product row ───────────────────────────────────────────────
    /// <summary>
    /// Inserts or updates the SharePoint row for a single extracted product line.
    /// Uniqueness key: (JobReference, SupplierName, ProductName).
    /// If a matching row already exists it is updated in-place; otherwise a new row is inserted.
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
            var siteId      = await GetSiteIdAsync();
            var listId      = await GetListIdAsync();
            var rawJobRef = (header.JobReference ?? emailMeta.JobRefs.FirstOrDefault()?.Trim('[', ']') ?? string.Empty).ToUpperInvariant();
            // No job reference found → record under sentinel [000000] so no emails are silently dropped.
            var jobRef = string.IsNullOrEmpty(rawJobRef) ? "000000" : rawJobRef;

            // When Claude can't extract a supplier name, derive a matchable token from the
            // sender's email domain (e.g. "sales@ryerson.com" → "ryerson") so the fuzzy
            // resolver can still find the canonical supplier name for deduplication.
            var rawSupplier = header.SupplierName;
            if (string.IsNullOrWhiteSpace(rawSupplier) && !string.IsNullOrWhiteSpace(emailMeta.EmailFrom))
            {
                var addr = emailMeta.EmailFrom;
                if (addr.Contains('@'))
                {
                    var domain = addr.Split('@').Last();        // e.g. "ryerson.com"
                    var parts  = domain.Split('.');
                    rawSupplier = parts.Length >= 2 ? parts[^2] : parts[0]; // "ryerson"
                }
                else
                {
                    rawSupplier = addr;
                }
            }
            rawSupplier ??= string.Empty;
            var supplier = _suppliers.ResolveSupplierName(rawSupplier);
            if (supplier is null)
            {
                // Supplier not in the reference list — still write the row so nothing is lost.
                // Override job ref to [WHOIS] as a diagnostic flag, and use "Unknown" as the supplier name.
                result.SupplierUnknown = true;
                supplier = "Unknown";
                jobRef   = "WHOIS";
                _log.LogInformation("[SP] Supplier '{Raw}' not in reference list — writing row under [WHOIS]", rawSupplier);
            }

            var prodName   = product.ProductName ?? $"Product {rowIndex + 1}";
            var prodTokens = ProductTokens(prodName);

            // Fuzzy-match against the product catalog. Falls back to null (vendor's description used by dashboard).
            var catalogMatch   = _catalog.ResolveProduct(prodName);
            var catalogName    = catalogMatch?.Name;
            var catalogKey     = catalogMatch?.SearchKey;
            if (catalogName is not null && !string.Equals(catalogName, prodName, StringComparison.OrdinalIgnoreCase))
                _log.LogDebug("[SP] Product '{Raw}' → catalog '{Catalog}'", prodName, catalogName);

            var title = $"[{jobRef}] {supplier} - {prodName}";
            title = title[..Math.Min(title.Length, 255)];

            var fieldData = new Dictionary<string, object?>
            {
                ["Title"]                   = title,
                ["JobReference"]            = string.IsNullOrEmpty(jobRef) ? null : jobRef,
                ["EmailFrom"]               = emailMeta.EmailFrom,
                ["ReceivedAt"]              = emailMeta.ReceivedAt,
                ["ProcessedAt"]             = DateTime.UtcNow.ToString("o"),
                ["ProcessingSource"]        = source,
                ["SourceFile"]              = sourceFile,
                ["SupplierName"]            = supplier,
                ["QuoteReference"]          = header.QuoteReference,
                ["DateOfQuote"]             = header.DateOfQuote,
                ["EstimatedDeliveryDate"]   = header.EstimatedDeliveryDate,
                ["FreightTerms"]            = header.FreightTerms,
                ["ProductName"]             = prodName,
                ["UnitsRequested"]          = product.UnitsRequested,
                ["UnitsQuoted"]             = product.UnitsQuoted,
                ["LengthPerUnit"]           = product.LengthPerUnit,
                ["LengthUnit"]              = product.LengthUnit,
                ["WeightPerUnit"]           = product.WeightPerUnit,
                ["WeightUnit"]              = product.WeightUnit,
                ["PricePerPound"]           = product.PricePerPound,
                ["PricePerFoot"]            = product.PricePerFoot,
                ["PricePerPiece"]           = product.PricePerPiece,
                ["TotalPrice"]              = product.TotalPrice,
                ["LeadTimeText"]            = product.LeadTimeText,
                ["Certifications"]          = product.Certifications,
                ["SupplierProductComments"] = product.SupplierProductComments,
                ["CatalogProductName"]      = catalogName,
                ["ProductSearchKey"]        = catalogKey,
                ["EmailBody"]               = emailMeta.EmailBody is not null
                    ? emailMeta.EmailBody[..Math.Min(emailMeta.EmailBody.Length,
                          int.TryParse(_config["SharePoint:MaxEmailBodyChars"], out var mebc) ? mebc : 10_000)]
                    : emailMeta.BodyContext,
            };

            var fields = new FieldValueSet { AdditionalData = fieldData };

            // Check for an existing row with the same job / supplier / quote / product.
            var existing = await FindExistingItemAsync(siteId, listId, jobRef, supplier,
                header.QuoteReference, prodName, prodTokens);

            if (existing is not null)
            {
                // Update in-place — preserve the canonical ProductName from the first extraction
                // so that minor Claude wording variations don't rename the row on every reprocess.
                var updateData = new Dictionary<string, object?>(fieldData);
                updateData.Remove("ProductName");
                await GetGraph().Sites[siteId].Lists[listId].Items[existing.Value.Id].Fields
                    .PatchAsync(new FieldValueSet { AdditionalData = updateData });

                result.Success  = true;
                result.Updated  = true;
                result.SpItemId = existing.Value.Id;
                result.SpWebUrl = existing.Value.WebUrl;
                _log.LogInformation("[SP] Updated existing row for '{Name}' (item {Id})", prodName, existing.Value.Id);
            }
            else
            {
                // Insert new row.
                var item = await GetGraph().Sites[siteId].Lists[listId].Items
                    .PostAsync(new ListItem { Fields = fields });

                result.Success  = true;
                result.Updated  = false;
                result.SpItemId = item!.Id;
                result.SpWebUrl = item.WebUrl;
                _log.LogInformation("[SP] Inserted new row for '{Name}' -> item {Id}", prodName, item.Id);
            }

            // Upload the source attachment as a SharePoint list item attachment.
            // On insert:  always upload if attachment data is present.
            // On update:  only upload (and replace) if new attachment data is present;
            //             leave any existing attachment untouched when there is nothing to replace it with.
            if (result.SpItemId is not null &&
                emailMeta.SourceType == "attachment" &&
                !string.IsNullOrEmpty(emailMeta.FileName) &&
                !string.IsNullOrEmpty(emailMeta.Base64Data))
            {
                try
                {
                    var bytes = Convert.FromBase64String(emailMeta.Base64Data);
                    await UpsertItemAttachmentAsync(result.SpItemId, emailMeta.FileName, bytes);
                }
                catch (Exception ex)
                {
                    // Attachment failure is non-fatal — the row data is already written.
                    _log.LogWarning(ex, "[SP] Attachment upload failed for item {Id} ('{File}')", result.SpItemId, emailMeta.FileName);
                }
            }
        }
        catch (Microsoft.Graph.Models.ODataErrors.ODataError odataEx)
        {
            result.Success = false;
            result.Error   = odataEx.Message;
            _log.LogError("[SP] ODataError upserting '{Name}': code={Code} message={Msg} inner={Inner}",
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
            _log.LogError(ex, "[SP] Failed to upsert row for product '{Name}'", product.ProductName);
        }
        return result;
    }

    // ── Existence check ──────────────────────────────────────────────────────

    /// <summary>
    /// Returns the ID and WebUrl of an existing list item matching the given
    /// (JobReference, SupplierName, ProductName) key, or null if none exists.
    /// JobReference is filtered server-side; supplier and product are matched in
    /// memory using normalised comparison so minor formatting differences don't
    /// produce duplicate rows.
    /// </summary>
    private async Task<(string Id, string? WebUrl)?> FindExistingItemAsync(
        string siteId, string listId, string jobRef, string supplierName,
        string? quoteReference, string productName, HashSet<string> productTokens)
    {
        var filter = string.IsNullOrEmpty(jobRef)
            ? null
            : $"fields/JobReference eq '{EscapeOdata(jobRef)}'";

        // NOTE: JobReference must be indexed in SharePoint for this filter to be reliable at scale.
        // List settings → Indexed columns → Add JobReference.
        // The Prefer header lets it work on non-indexed columns but may be slow on large lists.
        var result = await GetGraph().Sites[siteId].Lists[listId].Items
            .GetAsync(r =>
            {
                if (filter is not null) r.QueryParameters.Filter = filter;
                r.QueryParameters.Expand = ["fields($select=SupplierName,QuoteReference,ProductName)"];
                r.QueryParameters.Top    = int.TryParse(_config["SharePoint:MaxSearchItems"], out var msi) ? msi : 500;
                r.Headers.Add("Prefer", "HonorNonIndexedQueriesWarningMayFailRandomly");
            });

        var match = result?.Value?.FirstOrDefault(i =>
        {
            var data = i.Fields?.AdditionalData;
            if (data is null) return false;
            var spSupplier = data.TryGetValue("SupplierName",   out var s) ? s?.ToString() : null;
            var spQuoteRef = data.TryGetValue("QuoteReference", out var q) ? q?.ToString() : null;
            var spProduct  = data.TryGetValue("ProductName",    out var p) ? p?.ToString() : null;

            // Resolve stored supplier through the fuzzy resolver so variants of the same
            // company name (e.g. "Hadco Metal (Casey Krauss)" vs "Hadco") match correctly.
            var resolvedSp = _suppliers.ResolveSupplierName(spSupplier) ?? spSupplier;
            if (!string.Equals(resolvedSp, supplierName, StringComparison.OrdinalIgnoreCase)) return false;

            // Quote reference check: when both sides have a quote ref they must agree.
            // If either is absent, skip this check (fall through to product matching).
            if (!string.IsNullOrWhiteSpace(quoteReference) && !string.IsNullOrWhiteSpace(spQuoteRef))
            {
                if (!string.Equals(quoteReference.Trim(), spQuoteRef.Trim(), StringComparison.OrdinalIgnoreCase))
                    return false;
            }

            // Accept product match via exact-normalised string OR fuzzy token match.
            // The numeric-superset rule must hold before Jaccard is applied so that
            // different grades / sizes never collapse into the same row.
            if (NormalizeMatch(spProduct, productName)) return true;
            var spTokens = ProductTokens(spProduct ?? string.Empty);
            return NumericTokensCompatible(productTokens, spTokens)
                && ProductJaccard(spTokens, productTokens) >= 0.5;
        });

        if (match is null) return null;
        return (match.Id!, match.WebUrl);
    }

    private static string EscapeOdata(string s) => s.Replace("'", "''");

    // Normalise for comparison: lowercase, collapse whitespace and punctuation.
    private static readonly Regex _normaliseRegex = new(@"[\s\W]+", RegexOptions.Compiled);
    private static bool NormalizeMatch(string? a, string? b)
    {
        if (a is null && b is null) return true;
        if (a is null || b is null) return false;
        static string N(string s) => _normaliseRegex.Replace(s.Trim().ToLowerInvariant(), " ").Trim();
        return N(a) == N(b);
    }

    // ── Product tokenisation ─────────────────────────────────────────────────
    //
    // Dimension notation is normalised into single compound tokens BEFORE splitting
    // so that "1\" x 2\"", "1 x 2", and "1x2" all become the token "1x2", and
    // "3/16" becomes "3f16".  Single-character tokens are only dropped when they
    // contain no digit (removes stray 'x' separators, connector words, etc.),
    // so bare numeric dimensions like a "2" in "2\" OD" are preserved.

    private static readonly Regex _dimFraction  = new(@"(\d+)/(\d+)",                               RegexOptions.Compiled);
    private static readonly Regex _dimDecimal   = new(@"(\d+)\.(\d+)",                              RegexOptions.Compiled);
    private static readonly Regex _dimSeparator = new(@"(\d[a-z0-9]*)[""']?\s*[xX×]\s*[""']?(\d[a-z0-9]*)", RegexOptions.Compiled);
    private static readonly Regex _dimSplit     = new(@"[^a-z0-9]+",                                RegexOptions.Compiled);

    // Strips alternate-length clauses like "or 12'" / "or 12\" lengths" that Claude
    // sometimes appends when a supplier quotes multiple length options.
    private static readonly Regex _orLength =
        new(@"\bor\s+\d+[a-z""']*\b", RegexOptions.Compiled | RegexOptions.IgnoreCase);

    private static string PreprocessProduct(string s)
    {
        s = s.ToLowerInvariant();
        // Strip noise before tokenising: trailing length qualifiers and alternate-length clauses.
        s = _orLength.Replace(s, "");           // "10' or 12'" → "10'"
        s = Regex.Replace(s, @"\brandom\s+lengths?\b|\bmill\s+lengths?\b|\bfull\s+lengths?\b|\blengths?\b", "");
        s = _dimFraction.Replace(s, "$1f$2");   // 3/16  → 3f16
        s = _dimDecimal.Replace(s, "$1d$2");    // 1.5   → 1d5
        // Strip trailing zeros from d-decimal tokens so 6d000 == 6d0 == 6d
        s = Regex.Replace(s, @"d(\d+)", m =>
        {
            var stripped = m.Groups[1].Value.TrimEnd('0');
            return "d" + (stripped.Length == 0 ? "0" : stripped);
        });                                     // 6d000 → 6d0, 6d500 → 6d5
        s = _dimSeparator.Replace(s, "$1x$2");  // 1 x 2 → 1x2  (first pass)
        s = _dimSeparator.Replace(s, "$1x$2");  // second pass for 3-part: 1x2 x 3 → 1x2x3
        s = Regex.Replace(s, @"[""']", "");     // strip remaining inch/foot markers
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

    // Numeric/dimension tokens — any token that contains at least one digit.
    private static bool HasDigit(string t) => t.Any(char.IsDigit);

    /// <summary>
    /// Returns true when the numeric/dimension tokens of two product names are
    /// compatible for matching purposes.
    ///
    /// Three cases:
    ///   1. Both carry compound dimension tokens (contain x / f / d) → dimension
    ///      sets must be identical AND grade numbers must be compatible (subset).
    ///      Prevents 1x2 ↔ 2x4 and 304 ↔ 316 false matches.
    ///   2. Only A carries dimension tokens → block.  Matching would discard A's
    ///      dimension information into a dimensionless cache entry.
    ///   3. Only B carries dimension tokens (or neither does) → allow.  A is a
    ///      less-detailed form; B's dimensional canonical name will be preserved.
    /// </summary>
    private static bool NumericTokensCompatible(HashSet<string> a, HashSet<string> b)
    {
        var numA = a.Where(HasDigit).ToHashSet();
        var numB = b.Where(HasDigit).ToHashSet();

        var dimA = numA.Where(IsDimToken).ToHashSet();
        var dimB = numB.Where(IsDimToken).ToHashSet();

        if (dimA.Count > 0 && dimB.Count > 0)
        {
            // Case 1: both dimensioned — dims must match exactly, grades must be compatible.
            if (!dimA.SetEquals(dimB)) return false;
            var gradeA = numA.Where(t => !IsDimToken(t)).ToHashSet();
            var gradeB = numB.Where(t => !IsDimToken(t)).ToHashSet();
            return gradeA.IsSubsetOf(gradeB) || gradeB.IsSubsetOf(gradeA);
        }

        if (dimA.Count > 0) // dimB.Count == 0
            // Case 2: A has dims, B doesn't — block to avoid losing dimension info.
            return false;

        // Case 3: B has dims and A doesn't, OR neither has dims.
        // Grade numbers must be compatible (subset in either direction).
        var gA = numA.Where(t => !IsDimToken(t)).ToHashSet();
        var gB = numB.Where(t => !IsDimToken(t)).ToHashSet();
        return gA.IsSubsetOf(gB) || gB.IsSubsetOf(gA);
    }

    // A "compound dimension" token contains at least one digit AND one of the
    // separator characters introduced by PreprocessProduct:
    //   x  — cross-section separator (1x2, 2x4)
    //   f  — fraction separator     (3f16 from 3/16)
    //   d  — decimal separator      (1d5  from 1.5)
    private static bool IsDimToken(string t) =>
        t.Any(char.IsDigit) && t.Any(c => c == 'x' || c == 'f' || c == 'd');

    // ── Attachment upload (SharePoint REST API) ──────────────────────────────

    /// <summary>
    /// Uploads <paramref name="bytes"/> as a named attachment on a SharePoint list item.
    /// If an attachment with the same filename already exists it is deleted first,
    /// so this is effectively an upsert — the caller is responsible for only calling
    /// this method when a new file is actually available (preserving existing attachments
    /// when no replacement is provided is enforced by the caller, not here).
    /// </summary>
    private async Task UpsertItemAttachmentAsync(string spItemId, string fileName, byte[] bytes)
    {
        var siteUrl  = _config["SharePoint:SiteUrl"] ?? "https://metalsupermarkets.sharepoint.com/sites/hackensack";
        var uri      = new Uri(siteUrl);
        var host     = uri.Host;
        var sitePath = uri.AbsolutePath.TrimEnd('/');
        var listId   = await GetListIdAsync();

        // Acquire a SharePoint-scoped bearer token (audience differs from Graph).
        var tokenCtx = new Azure.Core.TokenRequestContext([$"https://{host}/.default"]);
        var token    = await GetSpCredential().GetTokenAsync(tokenCtx);

        // REST endpoint for list item attachments.
        var attBase = $"https://{host}{sitePath}/_api/web/lists(guid'{listId}')/items({spItemId})/AttachmentFiles";

        using var http = new HttpClient();
        http.DefaultRequestHeaders.Authorization =
            new System.Net.Http.Headers.AuthenticationHeaderValue("Bearer", token.Token);
        http.DefaultRequestHeaders.Add("Accept", "application/json;odata=nometadata");

        // List existing attachments to see if same-named file already exists.
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
                // Delete the existing attachment so we can replace it.
                var delUrl = $"{attBase}/getByFileName('{Uri.EscapeDataString(fileName)}')";
                using var delReq = new HttpRequestMessage(HttpMethod.Delete, delUrl);
                delReq.Headers.Add("IF-MATCH", "*");
                var delResp = await http.SendAsync(delReq);
                if (!delResp.IsSuccessStatusCode)
                {
                    var delErr = await delResp.Content.ReadAsStringAsync();
                    _log.LogWarning("[SP] Could not delete existing attachment '{File}': {Status} {Body}",
                        fileName, delResp.StatusCode, delErr[..Math.Min(delErr.Length, 200)]);
                }
            }
        }
        else
        {
            _log.LogDebug("[SP] Could not list attachments for item {Id}: {Status}", spItemId, listResp.StatusCode);
        }

        // Upload the new attachment.
        var uploadUrl = $"{attBase}/add(FileName='{Uri.EscapeDataString(fileName)}')";
        var fileContent = new ByteArrayContent(bytes);
        fileContent.Headers.ContentType =
            new System.Net.Http.Headers.MediaTypeHeaderValue("application/octet-stream");

        var addResp = await http.PostAsync(uploadUrl, fileContent);
        if (addResp.IsSuccessStatusCode)
        {
            _log.LogInformation("[SP] Uploaded attachment '{File}' ({Bytes} bytes) to item {Id}",
                fileName, bytes.Length, spItemId);
        }
        else
        {
            var err = await addResp.Content.ReadAsStringAsync();
            _log.LogWarning("[SP] Failed to upload attachment '{File}': {Status} {Body}",
                fileName, addResp.StatusCode, err[..Math.Min(err.Length, 400)]);
        }
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
            ("PricePerFoot",           "number"),
            ("PricePerPiece",          "number"),
            ("TotalPrice",             "number"),
            ("LeadTimeText",           "text"),
            ("Certifications",         "text"),
            ("FreightTerms",           "text"),
            ("SupplierProductComments","note"),
            ("CatalogProductName",     "text"),
            ("ProductSearchKey",       "text"),
            ("EmailBody",              "note"),
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
