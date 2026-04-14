using Microsoft.AspNetCore.Mvc;
using OutlookShredder.Proxy.Models;
using OutlookShredder.Proxy.Services;

namespace OutlookShredder.Proxy.Controllers;

/// <summary>
/// Dry-run endpoint for testing RLI anchoring against existing SharePoint data.
/// Never writes to SharePoint — read-only comparison only.
/// </summary>
[ApiController]
[Route("api/rli-anchoring")]
public class RliAnchoringController : ControllerBase
{
    private readonly SharePointService _sp;
    private readonly ClaudeService     _claude;
    private readonly ILogger<RliAnchoringController> _log;

    public RliAnchoringController(
        SharePointService sp,
        ClaudeService     claude,
        ILogger<RliAnchoringController> log)
    {
        _sp     = sp;
        _claude = claude;
        _log    = log;
    }

    /// <summary>
    /// Compares existing SLI ProductSearchKey values against the RLI MSPC for the same RFQ.
    /// When withClaude=true, re-runs Claude extraction on the stored SR EmailBody with RLI
    /// context injected and shows what productSearchKey Claude would now return.
    ///
    /// GET /api/rli-anchoring/compare?sampleSize=20[&withClaude=true&claudeLimit=5]
    /// </summary>
    [HttpGet("compare")]
    public async Task<IActionResult> Compare(
        [FromQuery] int  sampleSize   = 20,
        [FromQuery] bool withClaude   = false,
        [FromQuery] int  claudeLimit  = 5)
    {
        _log.LogInformation("[RliTest] Starting dry-run compare: sampleSize={N}, withClaude={WC}",
            sampleSize, withClaude);

        // ── 1. Sample RFQ IDs that have RLI rows ─────────────────────────────
        var allRli = await _sp.ReadAllRfqLineItemsAsync();

        // Group by RFQ ID, drop IDs that look invalid, sample up to sampleSize
        var rfqIds = allRli
            .Select(r => r.RfqId)
            .Where(id => !string.IsNullOrEmpty(id) && id != "000000" && id != "WHOIS")
            .Distinct()
            .OrderByDescending(id => id)  // most recent first (lexicographic, works for YYYY.NNN style)
            .Take(sampleSize)
            .ToList();

        _log.LogInformation("[RliTest] Found {Count} unique RFQ IDs with RLI rows (capped at {N})",
            rfqIds.Count, sampleSize);

        var results = new List<RfqCompareResult>();
        int claudeCallsUsed = 0;

        foreach (var rfqId in rfqIds)
        {
            // ── 2. Get RLI items for this RFQ ─────────────────────────────
            var rliItems = allRli
                .Where(r => string.Equals(r.RfqId, rfqId, StringComparison.OrdinalIgnoreCase))
                .Select(r => new RliContextItem { Mspc = r.Mspc, ProductName = r.Product })
                .ToList();

            // ── 3. Get existing SLI rows for this RFQ ─────────────────────
            var sliRows = await _sp.ReadSliCompactByRfqIdAsync(rfqId);

            // ── 4. Build data-only comparison ─────────────────────────────
            var productRows = sliRows.Select(sli =>
            {
                // Find best-matching RLI item by name (simple substring heuristic)
                var bestRli = FindBestRliMatch(sli.ProductName ?? sli.SupplierProductName, rliItems);

                return new ProductCompareRow
                {
                    SupplierName        = sli.SupplierName,
                    SupplierProductName = sli.SupplierProductName ?? sli.ProductName,
                    ExistingSearchKey   = sli.ProductSearchKey,
                    ExistingCatalogName = sli.CatalogProductName,
                    BestRliMspc         = bestRli?.Mspc,
                    BestRliProductName  = bestRli?.ProductName,
                    DataMatch           = bestRli?.Mspc is not null &&
                                         string.Equals(sli.ProductSearchKey, bestRli.Mspc,
                                             StringComparison.OrdinalIgnoreCase),
                };
            }).ToList();

            // ── 5. Optionally re-run Claude with RLI context ───────────────
            if (withClaude && claudeCallsUsed < claudeLimit && rliItems.Count > 0)
            {
                var srRows = await _sp.ReadSrEmailsByRfqIdAsync(rfqId);
                foreach (var sr in srRows)
                {
                    if (claudeCallsUsed >= claudeLimit) break;
                    if (string.IsNullOrWhiteSpace(sr.EmailBody)) continue;

                    var req = new ExtractRequest
                    {
                        Content      = sr.EmailBody,
                        SourceType   = "body",
                        JobRefs      = [rfqId],
                        EmailSubject = sr.EmailSubject,
                        EmailFrom    = sr.EmailFrom,
                        RliItems     = rliItems,
                    };

                    try
                    {
                        _log.LogInformation("[RliTest] Claude dry-run for [{RfqId}] / {Supplier}",
                            rfqId, sr.SupplierName);
                        var extraction = await _claude.ExtractAsync(req);
                        claudeCallsUsed++;

                        // Annotate matching product rows with Claude's new suggestion
                        if (extraction?.Products is { Count: > 0 } products)
                        {
                            foreach (var (p, idx) in products.Select((p, i) => (p, i)))
                            {
                                // Match by index (products should be in same order)
                                var target = productRows
                                    .Where(r => string.Equals(r.SupplierName, sr.SupplierName,
                                        StringComparison.OrdinalIgnoreCase))
                                    .Skip(idx).FirstOrDefault();

                                if (target is not null)
                                {
                                    target.ClaudeProductName  = p.ProductName;
                                    target.ClaudeSearchKey    = p.ProductSearchKey;
                                    target.ClaudeAgreesWithRli = p.ProductSearchKey is not null &&
                                        string.Equals(p.ProductSearchKey, target.BestRliMspc,
                                            StringComparison.OrdinalIgnoreCase);
                                    target.ClaudeAgreesWithExisting = string.Equals(
                                        p.ProductSearchKey, target.ExistingSearchKey,
                                        StringComparison.OrdinalIgnoreCase);
                                }
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                        _log.LogWarning(ex, "[RliTest] Claude dry-run failed for [{RfqId}]", rfqId);
                    }
                }
            }

            results.Add(new RfqCompareResult
            {
                RfqId       = rfqId,
                RliItems    = rliItems,
                ProductRows = productRows,
            });
        }

        var summary = new
        {
            RfqsSampled     = results.Count,
            TotalSliRows    = results.Sum(r => r.ProductRows.Count),
            DataMatches     = results.SelectMany(r => r.ProductRows).Count(p => p.DataMatch),
            DataMismatches  = results.SelectMany(r => r.ProductRows).Count(p => !p.DataMatch && p.BestRliMspc is not null),
            NoRliMatch      = results.SelectMany(r => r.ProductRows).Count(p => p.BestRliMspc is null),
            ClaudeCallsUsed = claudeCallsUsed,
        };

        _log.LogInformation("[RliTest] Done. {Summary}", System.Text.Json.JsonSerializer.Serialize(summary));

        return Ok(new { Summary = summary, Results = results });
    }

    /// <summary>
    /// Naive best-match: finds the RLI item whose ProductName shares the most
    /// tokens with the supplier's product name. Good enough for a dry-run comparison.
    /// </summary>
    private static RliContextItem? FindBestRliMatch(string? supplierName, List<RliContextItem> rliItems)
    {
        if (string.IsNullOrWhiteSpace(supplierName) || rliItems.Count == 0) return null;

        var supplierTokens = Tokenize(supplierName);
        if (supplierTokens.Count == 0) return null;

        return rliItems
            .Where(r => !string.IsNullOrEmpty(r.ProductName))
            .Select(r => new
            {
                Item  = r,
                Score = Tokenize(r.ProductName!).Intersect(supplierTokens, StringComparer.OrdinalIgnoreCase).Count(),
            })
            .Where(x => x.Score > 0)
            .OrderByDescending(x => x.Score)
            .FirstOrDefault()?.Item;
    }

    private static HashSet<string> Tokenize(string s) =>
        [.. s.ToLowerInvariant()
             .Split([' ', '-', '/', '.', ',', '×', 'x', '"', '\'', '\t'], StringSplitOptions.RemoveEmptyEntries)
             .Where(t => t.Length >= 2)];
}

// ── Response models ───────────────────────────────────────────────────────────

public class RfqCompareResult
{
    public string              RfqId       { get; set; } = "";
    public List<RliContextItem> RliItems   { get; set; } = [];
    public List<ProductCompareRow> ProductRows { get; set; } = [];
}

public class ProductCompareRow
{
    public string? SupplierName        { get; set; }
    public string? SupplierProductName { get; set; }

    // ── Existing (fuzzy match) ────────────────────────────────────────────────
    public string? ExistingSearchKey   { get; set; }
    public string? ExistingCatalogName { get; set; }

    // ── Best data-level RLI match (by token overlap) ─────────────────────────
    public string? BestRliMspc         { get; set; }
    public string? BestRliProductName  { get; set; }
    /// <summary>True when ExistingSearchKey already matches the RLI MSPC — fuzzy match got it right.</summary>
    public bool    DataMatch           { get; set; }

    // ── Claude dry-run result (withClaude=true only) ──────────────────────────
    public string? ClaudeProductName         { get; set; }
    public string? ClaudeSearchKey           { get; set; }
    /// <summary>Claude's productSearchKey agrees with the RLI MSPC.</summary>
    public bool?   ClaudeAgreesWithRli       { get; set; }
    /// <summary>Claude's productSearchKey agrees with the existing fuzzy match.</summary>
    public bool?   ClaudeAgreesWithExisting  { get; set; }
}
