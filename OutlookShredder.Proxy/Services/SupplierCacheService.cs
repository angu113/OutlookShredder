using Azure.Identity;
using Microsoft.Graph;

namespace OutlookShredder.Proxy.Services;

/// <summary>
/// Singleton background service that fetches the canonical supplier list from a
/// SharePoint/OneDrive list, caches it in memory, and refreshes every hour.
///
/// Config keys (optional — defaults shown):
///   Suppliers:SiteUrl   — full URL of the personal/OneDrive site that owns the list
///                         default: https://metalsupermarkets-my.sharepoint.com/personal/angus_mithrilmetals_com
///   Suppliers:ListName  — display name of the list
///                         default: Suppliers
///
/// Uses the same app-only credentials as SharePointService.
/// The Azure AD app must have Sites.ReadWrite.All (or Sites.Read.All) consented for
/// the OneDrive personal site as well as the main SharePoint site.
/// </summary>
public class SupplierCacheService : BackgroundService
{
    private readonly IConfiguration _config;
    private readonly ILogger<SupplierCacheService> _log;
    private GraphServiceClient? _graph;

    // Pre-tokenised entries for lock-free reads — reference swap is atomic.
    private sealed record Entry(string Name, HashSet<string> Tokens, string? EmailDomain);
    private volatile IReadOnlyList<Entry> _cache = [];

    public SupplierCacheService(IConfiguration config, ILogger<SupplierCacheService> log)
    {
        _config = config;
        _log    = log;
    }

    // ── Background loop ───────────────────────────────────────────────────────

    protected override async Task ExecuteAsync(CancellationToken stoppingToken)
    {
        // Refresh immediately on startup, then every hour.
        while (!stoppingToken.IsCancellationRequested)
        {
            await RefreshAsync();
            try { await Task.Delay(TimeSpan.FromHours(1), stoppingToken); }
            catch (OperationCanceledException) { break; }
        }
    }

    private async Task RefreshAsync()
    {
        try
        {
            var entries = await FetchEntriesAsync();
            _cache = entries.AsReadOnly();

            _log.LogInformation("[Suppliers] Cache refreshed — {Count} supplier(s) loaded", entries.Count);
        }
        catch (Exception ex)
        {
            _log.LogWarning(ex, "[Suppliers] Cache refresh failed — stale cache will be used");
        }
    }

    // ── Graph fetch ───────────────────────────────────────────────────────────

    private async Task<List<Entry>> FetchEntriesAsync()
    {
        var graph = GetGraph();

        var siteUrl = _config["Suppliers:SiteUrl"]
            ?? "https://metalsupermarkets-my.sharepoint.com/personal/angus_mithrilmetals_com";
        var uri     = new Uri(siteUrl);
        var siteKey = $"{uri.Host}:{uri.AbsolutePath}";

        var site = await graph.Sites[siteKey].GetAsync();
        if (site?.Id is null)
        {
            _log.LogWarning("[Suppliers] Could not resolve site '{Key}'", siteKey);
            return [];
        }

        var listName = _config["Suppliers:ListName"] ?? "Suppliers";
        var lists = await graph.Sites[site.Id].Lists
            .GetAsync(r => r.QueryParameters.Filter = $"displayName eq '{listName}'");

        var list = lists?.Value?.FirstOrDefault();
        if (list?.Id is null)
        {
            _log.LogWarning("[Suppliers] List '{Name}' not found in site '{Key}'", listName, siteKey);
            return [];
        }

        var items = await graph.Sites[site.Id].Lists[list.Id].Items
            .GetAsync(r =>
            {
                r.QueryParameters.Expand = ["fields($select=Title,ContactEmail)"];
                r.QueryParameters.Top    = 1000;
            });

        return items?.Value?
            .Where(i => i.Fields?.AdditionalData != null)
            .Select(i =>
            {
                var d = i.Fields!.AdditionalData!;
                var name  = d.TryGetValue("Title",        out var t) ? t?.ToString() : null;
                var email = d.TryGetValue("ContactEmail", out var e) ? e?.ToString() : null;
                if (string.IsNullOrWhiteSpace(name)) return null;
                var domain = ExtractDomain(email);
                return new Entry(name, Tokenize(name), domain);
            })
            .Where(e => e is not null)
            .Select(e => e!)
            .ToList()
            ?? [];
    }

    private static string? ExtractDomain(string? email)
    {
        if (string.IsNullOrWhiteSpace(email)) return null;
        var at = email.IndexOf('@');
        return at >= 0 ? email[(at + 1)..].ToLowerInvariant() : null;
    }

    private GraphServiceClient GetGraph()
    {
        if (_graph is not null) return _graph;
        var tenantId     = _config["SharePoint:TenantId"]     ?? throw new InvalidOperationException("SharePoint:TenantId not set");
        var clientId     = _config["SharePoint:ClientId"]     ?? throw new InvalidOperationException("SharePoint:ClientId not set");
        var clientSecret = _config["SharePoint:ClientSecret"] ?? throw new InvalidOperationException("SharePoint:ClientSecret not set");
        var credential   = new ClientSecretCredential(tenantId, clientId, clientSecret);
        _graph = new GraphServiceClient(credential, ["https://graph.microsoft.com/.default"]);
        return _graph;
    }

    // ── Public resolver ───────────────────────────────────────────────────────

    /// <summary>
    /// Returns the canonical supplier name from the Suppliers list if a fuzzy
    /// match is found (Jaccard token similarity ≥ 0.35).
    /// Returns <see langword="null"/> when the cache is loaded and no match exists.
    /// Returns <paramref name="rawName"/> unchanged only when the cache is empty
    /// (not yet loaded), so startup processing is not blocked.
    /// </summary>
    public string? ResolveSupplierName(string? rawName)
    {
        if (string.IsNullOrWhiteSpace(rawName)) return null;

        var cache = _cache;   // snapshot — no lock needed
        if (cache.Count == 0) return null;   // cache not yet loaded — skip, don't use raw name

        var rawTokens = Tokenize(rawName);

        // Strategy 1: Jaccard similarity ≥ 0.35.
        var best = cache
            .Select(e => (e.Name, Score: JaccardScore(rawTokens, e.Tokens)))
            .Where(x => x.Score >= 0.35)
            .OrderByDescending(x => x.Score)
            .FirstOrDefault();

        if (best.Name is not null)
        {
            if (!string.Equals(best.Name, rawName, StringComparison.OrdinalIgnoreCase))
                _log.LogDebug("[Suppliers] '{Raw}' → '{Canonical}' (jaccard {Score:F2})", rawName, best.Name, best.Score);
            return best.Name;
        }

        // Strategy 2: all canonical tokens are present in the raw name (containment).
        // Handles "Hadco Metal (Casey Krauss)" → "Hadco" where Jaccard is diluted by
        // extra personal-name tokens but the canonical name is clearly embedded.
        var contained = cache
            .Where(e => e.Tokens.Count > 0 && e.Tokens.IsSubsetOf(rawTokens))
            .OrderByDescending(e => e.Tokens.Count)   // prefer most specific match
            .FirstOrDefault();

        if (contained is not null)
        {
            if (!string.Equals(contained.Name, rawName, StringComparison.OrdinalIgnoreCase))
                _log.LogDebug("[Suppliers] '{Raw}' → '{Canonical}' (containment)", rawName, contained.Name);
            return contained.Name;
        }

        return null;   // no match in reference list
    }

    /// <summary>Exposes the current cached supplier names (for API/dashboard use).</summary>
    public IReadOnlyList<string> CachedNames => _cache.Select(e => e.Name).ToList().AsReadOnly();

    /// <summary>
    /// Returns a map of email domain → canonical supplier name, built from the
    /// ContactEmail column of the Suppliers list.
    /// </summary>
    public IReadOnlyDictionary<string, string> DomainMap =>
        _cache
            .Where(e => e.EmailDomain is not null)
            .GroupBy(e => e.EmailDomain!, StringComparer.OrdinalIgnoreCase)
            .ToDictionary(g => g.Key, g => g.First().Name, StringComparer.OrdinalIgnoreCase);

    /// <summary>
    /// Fallback when the domain is not in <see cref="DomainMap"/>: checks whether all
    /// canonical name tokens appear as substrings in <paramref name="domainPart"/>.
    /// e.g. "certifiedsteel" matches "Certified Steel" (tokens: certified, steel).
    /// Returns the first matching canonical name, or <see langword="null"/>.
    /// </summary>
    public string? ResolveByDomainSubstring(string domainPart)
    {
        if (string.IsNullOrWhiteSpace(domainPart)) return null;
        var lower = domainPart.ToLowerInvariant();
        return _cache
            .Where(e => e.Tokens.Count > 0 && e.Tokens.All(t => lower.Contains(t)))
            .Select(e => e.Name)
            .FirstOrDefault();
    }

    // ── Helpers ───────────────────────────────────────────────────────────────

    private static readonly char[] _delimiters = [' ', ',', '.', '-', '/', '&', '(', ')', '_'];

    // Strip common legal suffixes before tokenising so "Acme Inc." matches "Acme".
    private static readonly string[] _stopWords =
        ["inc", "llc", "ltd", "co", "corp", "company", "group", "international", "the"];

    private static HashSet<string> Tokenize(string s) =>
        s.ToLowerInvariant()
         .Split(_delimiters, StringSplitOptions.RemoveEmptyEntries)
         .Where(t => t.Length > 1 && !_stopWords.Contains(t))
         .ToHashSet();

    private static double JaccardScore(HashSet<string> a, HashSet<string> b)
    {
        if (a.Count == 0 && b.Count == 0) return 1.0;
        var intersection = a.Count(t => b.Contains(t));
        var union        = a.Count + b.Count - intersection;
        return union == 0 ? 0 : (double)intersection / union;
    }
}
