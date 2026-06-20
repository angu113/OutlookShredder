using Microsoft.AspNetCore.Mvc;
using OutlookShredder.Proxy.Services;

namespace OutlookShredder.Proxy.Controllers;

/// <summary>
/// Sales-Order history: the one-time bulk load of the Sales-Orders export into the <c>SalesOrderHistory</c>
/// SP list, and the per-customer serving endpoint that backs the Raptor incoming-call card.
/// </summary>
[ApiController]
[Route("api/sales-orders")]
public class SalesOrdersController(
    CustomerImportService     importer,
    SharePointService         sp,
    SalesOrderHistoryService  history,
    IConfiguration            config,
    ILogger<SalesOrdersController> log) : ControllerBase
{
    /// <summary>
    /// POST /api/sales-orders/import-history[?dryRun=true][&amp;file=Name.csv] — loads the Sales-Orders export
    /// from the Import directory into the SalesOrderHistory list.
    ///
    /// dryRun=true (default false): parses + computes dedup-vs-existing counts; writes nothing.
    /// dryRun=false: kicks off a background, batched, resumable load and returns immediately (202). Poll
    /// GET /api/sales-orders/import-history/status for progress.
    ///
    /// The source CSV is the file named by <c>file</c>, else the first Import-dir CSV whose header carries
    /// the distinctive <c>Order Date</c> + <c>Doc #</c> + <c>Customer</c> columns.
    /// </summary>
    [HttpPost("import-history")]
    public async Task<IActionResult> ImportHistory(
        [FromQuery] bool dryRun = false, [FromQuery] string? file = null, CancellationToken ct = default)
    {
        // Provision the list (indexed at construction) before any read/write — idempotent.
        await sp.EnsureSalesOrderHistoryListAsync(ct);

        var importDir = ResolveImportDir();
        var (path, readError) = ResolveSalesOrdersCsv(importDir, file);
        if (path is null)
            return BadRequest(new { importDir, error = readError ?? "No Sales-Orders CSV found in the import directory." });

        string csv;
        try { csv = await System.IO.File.ReadAllTextAsync(path, ct); }
        catch (Exception ex) { return BadRequest(new { file = Path.GetFileName(path), error = ex.Message }); }

        var parsed = importer.ParseSalesOrders(csv);
        log.LogInformation("[SalesOrders] {Mode} {File}: parsed={P}",
            dryRun ? "DryRun" : "Import", Path.GetFileName(path), parsed.Rows.Count);

        if (dryRun)
        {
            var existing = await sp.ReadAllSalesOrderIdsAsync(ct);
            int present  = parsed.Rows.Count(r => existing.Contains(r.OrderId));
            int distinctCustomers = parsed.Rows
                .Select(r => r.CustomerName.Trim()).Where(n => n.Length > 0)
                .Distinct(StringComparer.OrdinalIgnoreCase).Count();
            return Ok(new
            {
                file = Path.GetFileName(path), dryRun = true,
                parsed = parsed.Rows.Count,
                distinctCustomers,
                existingRows = existing.Count,
                alreadyPresent = present,
                toAdd = parsed.Rows.Count - present,
                warnings = parsed.Warnings,
            });
        }

        if (parsed.Rows.Count == 0)
            return BadRequest(new { file = Path.GetFileName(path), error = "Parsed 0 rows.", warnings = parsed.Warnings });

        if (!history.StartImport(parsed.Rows, out var message))
            return Conflict(new { message, status = history.ImportStatus });

        return Accepted(new { file = Path.GetFileName(path), message, status = history.ImportStatus });
    }

    /// <summary>GET /api/sales-orders/import-history/status — live progress of the background bulk load.</summary>
    [HttpGet("import-history/status")]
    public IActionResult ImportStatus() => Ok(history.ImportStatus);

    /// <summary>GET /api/sales-orders/by-customer?customer={bp}&amp;top=5 — a customer's recent orders +
    /// summary header for the Raptor call card. Served from the in-memory cache.</summary>
    [HttpGet("by-customer")]
    public async Task<IActionResult> ByCustomer(
        [FromQuery] string customer, [FromQuery] int top = 5, CancellationToken ct = default)
    {
        if (string.IsNullOrWhiteSpace(customer))
            return BadRequest(new { error = "customer is required" });

        var result = await history.GetOrdersForCustomerAsync(customer, top, ct);
        return Ok(result);
    }

    // ── Helpers ──────────────────────────────────────────────────────────────

    private string ResolveImportDir()
    {
        var configured = config["Import:Directory"];
        if (!string.IsNullOrWhiteSpace(configured))
            return configured;
        var localApp = Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData);
        return Path.Combine(localApp, "Shredder", "Import");
    }

    /// <summary>Picks the Sales-Orders CSV: the named <paramref name="preferred"/> file if given, else the
    /// first Import-dir CSV whose header has Order Date + Doc # + Customer. Returns (null, reason) on miss.</summary>
    private (string? Path, string? Error) ResolveSalesOrdersCsv(string importDir, string? preferred)
    {
        if (!Directory.Exists(importDir))
            return (null, $"Import directory does not exist: {importDir}");

        if (!string.IsNullOrWhiteSpace(preferred))
        {
            var p = Path.Combine(importDir, preferred);
            return System.IO.File.Exists(p) ? (p, null) : (null, $"File not found: {preferred}");
        }

        foreach (var fp in Directory.GetFiles(importDir, "*.csv"))
        {
            try
            {
                using var reader = new StreamReader(fp);
                var firstLine = (reader.ReadLine() ?? "").ToLowerInvariant();
                if (firstLine.Contains("order date") && firstLine.Contains("doc #") && firstLine.Contains("customer"))
                    return (fp, null);
            }
            catch { /* unreadable file — skip */ }
        }
        return (null, null);
    }
}
