using Microsoft.AspNetCore.Mvc;
using OutlookShredder.Proxy.Services;

namespace OutlookShredder.Proxy.Controllers;

[ApiController]
[Route("api/import")]
public class ImportController(
    CustomerImportService importer,
    SharePointService     sp,
    CustomerCacheService  crm,
    CallLogCrmBackfillService callLogBackfill,
    IConfiguration        config,
    ILogger<ImportController> log) : ControllerBase
{
    // File types that add/map customers or contacts — a successful real run of any of these can make
    // historic PhoneCallLog rows newly resolvable, so we re-run the blanks-only CRM backfill afterwards.
    private static readonly string[] CrmAffectingTypes = ["partners", "contacts", "ordercontacts", "customerinfo"];

    /// <summary>
    /// POST /api/import/run[?dryRun=true] — scans the import directory for CSV files and
    /// processes each one.
    ///
    /// dryRun=true (default false): parses files, downloads current SP data, returns the diff
    /// (what would be added/updated/deleted) without writing anything and without moving files.
    ///
    /// dryRun=false: applies changes to SP and moves each processed file to
    /// Import\processed\{timestamp}_{filename}.
    ///
    /// File type detection (first match wins):
    ///   - Filename contains "customer info" / "customerinfo" → customer info (enrichment)
    ///   - Filename contains "partner" or "bp"  → business partners
    ///   - Filename contains "sales order"       → order contacts (mined from the Sales-Orders export)
    ///   - Filename contains "contact"           → contacts
    ///   - Header contains "Popup Message"       → business partners
    ///   - Header contains "Order Date" + "Doc #" + "Contact" → order contacts
    ///   - Header contains "Contact Name" and "Customer Name" → contacts
    ///   - Otherwise: skipped with a warning
    ///
    /// If any BP rows were filtered out by a phrase match (duplicate, do not use, dupe),
    /// a review file is written to Import\_skipped_review_{timestamp}.csv.
    /// </summary>
    [HttpPost("run")]
    public async Task<IActionResult> Run([FromQuery] bool dryRun = false, CancellationToken ct = default)
    {
        var importDir = ResolveImportDir();

        Directory.CreateDirectory(importDir);
        var processedDir = Path.Combine(importDir, "processed");
        if (!dryRun) Directory.CreateDirectory(processedDir);

        var files = Directory.GetFiles(importDir, "*.csv");
        if (files.Length == 0)
            return Ok(new { importDir, dryRun, message = "No CSV files found in import directory.", files = Array.Empty<object>() });

        var timestamp  = DateTimeOffset.Now.ToString("yyyyMMdd_HHmmss");
        var results    = new List<object>();
        var reviewRows = new List<(string SourceFile, string BpName, string Reason)>();

        // Read + classify every file up front so the loads can be ORDERED: business partners, then
        // contacts, then Customer Info LAST — the Customer Info load only enriches records the partner
        // load has already created, so it must run after every new record exists.
        var loaded = new List<(string Path, string Name, string Csv, string Type, string? ReadError)>();
        foreach (var fp in files)
        {
            var fn = Path.GetFileName(fp);
            try
            {
                var text = await System.IO.File.ReadAllTextAsync(fp, ct);
                loaded.Add((fp, fn, text, DetectFileType(fn, text), null));
            }
            catch (Exception ex)
            {
                loaded.Add((fp, fn, "", "unknown", ex.Message));
            }
        }

        static int LoadOrder(string t) => t switch
        {
            "partners"      => 0,
            "contacts"      => 1,
            "ordercontacts" => 2,   // additive contacts mined from the orders export (needs customers to exist)
            "customerinfo"  => 3,   // enrichment — always last
            _               => 4,
        };
        var ordered = loaded
            .OrderBy(f => LoadOrder(f.Type))
            .ThenBy(f => f.Name, StringComparer.OrdinalIgnoreCase)
            .ToList();

        foreach (var (filePath, filename, csv, fileType, readError) in ordered)
        {
            log.LogInformation("[Import] {Mode} {File} ({Type})",
                dryRun ? "DryRun" : "Processing", filename, fileType);

            if (readError is not null)
            {
                results.Add(new { file = filename, error = readError });
                continue;
            }

            object result;
            try
            {
                (result, var skipped) = fileType switch
                {
                    "partners" => dryRun
                        ? await DiffPartnersAsync(filename, csv, ct)
                        : await ProcessPartnersAsync(filename, csv, ct),
                    "contacts" => dryRun
                        ? await DiffContactsAsync(filename, csv, ct)
                        : await ProcessContactsAsync(filename, csv, ct),
                    "ordercontacts" => dryRun
                        ? await DiffOrderContactsAsync(filename, csv, ct)
                        : await ProcessOrderContactsAsync(filename, csv, ct),
                    "customerinfo" => dryRun
                        ? await DiffCustomerInfoAsync(filename, csv, ct)
                        : await ProcessCustomerInfoAsync(filename, csv, ct),
                    _ => ((object)new { file = filename, skipped = true,
                                        reason = "Could not detect file type from filename or headers." },
                          Array.Empty<CustomerImportService.SkippedItem>())
                };
                foreach (var s in skipped)
                    reviewRows.Add((filename, s.Name, s.Reason));
            }
            catch (Exception ex)
            {
                log.LogError(ex, "[Import] Failed to process {File}", filename);
                result = new { file = filename, error = ex.Message };
            }

            results.Add(result);

            // Only move files on a real run
            if (!dryRun)
            {
                var dest = Path.Combine(processedDir, $"{timestamp}_{filename}");
                try { System.IO.File.Move(filePath, dest, overwrite: true); }
                catch (Exception ex)
                {
                    log.LogWarning(ex, "[Import] Could not move {File} to processed/", filename);
                }
            }
        }

        // Write review file if any phrase-matched rows were skipped
        string? reviewFile = null;
        if (reviewRows.Count > 0)
        {
            reviewFile = Path.Combine(importDir, $"_skipped_review_{timestamp}.csv");
            try
            {
                var lines = new List<string> { "SourceFile,BPName,Reason" };
                lines.AddRange(reviewRows.Select(r =>
                    $"{CsvQuote(r.SourceFile)},{CsvQuote(r.BpName)},{CsvQuote(r.Reason)}"));
                await System.IO.File.WriteAllLinesAsync(reviewFile, lines, ct);
                log.LogWarning("[Import] {Count} BP rows require manual review — see {File}",
                    reviewRows.Count, reviewFile);
            }
            catch (Exception ex)
            {
                log.LogError(ex, "[Import] Failed to write skipped review file");
                reviewFile = null;
            }
        }

        // After a real run that added/mapped customers or contacts, refresh the CRM cache and backfill blank
        // CRM fields on historic call-log rows so the call log / Raptor reflect the new mappings. Blanks-only
        // (never overwrites an existing BpName — that stays the manual includeChanged=true review path).
        // Fire-and-forget so the import response isn't blocked by the call-log scan.
        var crmBackfillQueued = false;
        if (!dryRun && ordered.Any(f => CrmAffectingTypes.Contains(f.Type)))
        {
            crmBackfillQueued = true;
            _ = Task.Run(async () =>
            {
                try
                {
                    await crm.RefreshNowAsync();   // pick up the just-imported customers/contacts before lookup
                    var r = await callLogBackfill.RunAsync(dryRun: false, includeChanged: false);
                    log.LogInformation(
                        "[Import] post-import call-log backfill: {Updated} filled, {Failed} failed (of {Total} rows; {Changed} changed left for review)",
                        r.Updated, r.Failed, r.TotalRecords, r.ChangeExisting);
                }
                catch (Exception ex)
                {
                    log.LogWarning(ex, "[Import] post-import call-log backfill failed");
                }
            });
        }

        return Ok(new { importDir, dryRun, reviewFile, skippedForReview = reviewRows.Count,
                        crmBackfillQueued, files = results });
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

    private static string DetectFileType(string filename, string csv)
    {
        var fn = filename.ToLowerInvariant();
        if (fn.Contains("customer info") || fn.Contains("customerinfo") || fn.Contains("cust info") || fn.Contains("custinfo"))
            return "customerinfo";
        if (fn.Contains("partner") || fn.Contains(" bp") || fn.StartsWith("bp") || fn.Contains("_bp"))
            return "partners";
        // Sales-Orders export mined for contacts — check before the generic "contact" hint so a file
        // named e.g. "sales orders - contacts.csv" routes to the orders parser, not the contacts parser.
        if (fn.Contains("sales order") || fn.Contains("salesorder") || fn.Contains("order contact"))
            return "ordercontacts";
        if (fn.Contains("contact"))
            return "contacts";

        var firstLine = csv.Split('\n', 2)[0].ToLowerInvariant();
        // Customer Info master: keyed on "Business Partner" plus the distinctive stats column.
        if (firstLine.Contains("business partner") && firstLine.Contains("margin type"))
            return "customerinfo";
        if (firstLine.Contains("popup message"))
            return "partners";
        // Sales-Orders export: distinctive Order Date + Doc # columns alongside the "Contact" column
        // ("[phone] - [name]"). Disjoint from the Contacts file (which has "Contact Name"/"Customer Name").
        if (firstLine.Contains("order date") && firstLine.Contains("doc #") && firstLine.Contains("contact"))
            return "ordercontacts";
        if (firstLine.Contains("contact name") && firstLine.Contains("customer name"))
            return "contacts";

        return "unknown";
    }

    // ── Live run ─────────────────────────────────────────────────────────────

    private async Task<(object result, IReadOnlyList<CustomerImportService.SkippedItem> skipped)>
        ProcessPartnersAsync(string filename, string csv, CancellationToken ct)
    {
        var parsed = importer.ParsePartners(csv);
        if (parsed.Rows.Count == 0 && parsed.Warnings.Count > 0)
            return (new { file = filename, type = "partners", parsed = 0,
                          skippedForReview = parsed.Skipped.Count, warnings = parsed.Warnings }, parsed.Skipped);

        var (added, updated, spSkipped) = await sp.UpsertBusinessPartnersAsync(parsed.Rows, ct);
        log.LogInformation("[Import] {File} partners: parsed={P} added={A} updated={U} spSkipped={S}",
            filename, parsed.Rows.Count, added, updated, spSkipped);
        return (new { file = filename, type = "partners", parsed = parsed.Rows.Count,
                      added, updated, spSkipped, skippedForReview = parsed.Skipped.Count,
                      warnings = parsed.Warnings }, parsed.Skipped);
    }

    private async Task<(object result, IReadOnlyList<CustomerImportService.SkippedItem> skipped)>
        ProcessContactsAsync(string filename, string csv, CancellationToken ct)
    {
        var parsed = importer.ParseContacts(csv);
        if (parsed.Rows.Count == 0 && parsed.Warnings.Count > 0)
            return (new { file = filename, type = "contacts", parsed = 0,
                          warnings = parsed.Warnings }, parsed.Skipped);

        var (added, unchanged) = await sp.UpsertContactsAsync(parsed.Rows, ct);
        log.LogInformation("[Import] {File} contacts: parsed={P} added={A} unchanged={U}",
            filename, parsed.Rows.Count, added, unchanged);
        return (new { file = filename, type = "contacts", parsed = parsed.Rows.Count,
                      added, unchanged, warnings = parsed.Warnings }, parsed.Skipped);
    }

    // ── Dry run ──────────────────────────────────────────────────────────────

    private async Task<(object result, IReadOnlyList<CustomerImportService.SkippedItem> skipped)>
        DiffPartnersAsync(string filename, string csv, CancellationToken ct)
    {
        var parsed = importer.ParsePartners(csv);
        if (parsed.Rows.Count == 0 && parsed.Warnings.Count > 0)
            return (new { file = filename, type = "partners", dryRun = true, parsed = 0,
                          skippedForReview = parsed.Skipped.Count, warnings = parsed.Warnings }, parsed.Skipped);

        var diff = await sp.DiffBusinessPartnersAsync(parsed.Rows, ct);
        return (new
        {
            file     = filename,
            type     = "partners",
            dryRun   = true,
            parsed   = parsed.Rows.Count,
            skippedForReview = parsed.Skipped.Count,
            toAdd    = diff.ToAdd,
            toUpdate = diff.ToUpdate,
            unchanged = diff.Unchanged,
            addSample    = diff.AddSample,
            updateSamples = diff.UpdateSamples,
            warnings = parsed.Warnings,
        }, parsed.Skipped);
    }

    private async Task<(object result, IReadOnlyList<CustomerImportService.SkippedItem> skipped)>
        DiffContactsAsync(string filename, string csv, CancellationToken ct)
    {
        var parsed = importer.ParseContacts(csv);
        if (parsed.Rows.Count == 0 && parsed.Warnings.Count > 0)
            return (new { file = filename, type = "contacts", dryRun = true, parsed = 0,
                          warnings = parsed.Warnings }, parsed.Skipped);

        var diff = await sp.DiffContactsAsync(parsed.Rows, ct);
        return (new
        {
            file         = filename,
            type         = "contacts",
            dryRun       = true,
            parsed       = parsed.Rows.Count,
            rowsToAdd    = diff.RowsToAdd,
            rowsUnchanged = diff.RowsUnchanged,
            warnings     = parsed.Warnings,
        }, parsed.Skipped);
    }

    // ── Order-sourced contacts (additive; existing customers only, unmatched reported) ──

    private async Task<(object result, IReadOnlyList<CustomerImportService.SkippedItem> skipped)>
        ProcessOrderContactsAsync(string filename, string csv, CancellationToken ct)
    {
        var parsed = importer.ParseContactsFromOrders(csv);
        if (parsed.Rows.Count == 0 && parsed.Warnings.Count > 0)
            return (new { file = filename, type = "ordercontacts", parsed = 0,
                          skippedForReview = parsed.Skipped.Count, warnings = parsed.Warnings }, parsed.Skipped);

        var r = await sp.UpsertContactsExistingOnlyAsync(parsed.Rows, ct);
        log.LogInformation(
            "[Import] {File} ordercontacts: parsed={P} matched={M} added={A} unchanged={U} unmatchedCustomers={X}",
            filename, parsed.Rows.Count, r.MatchedRows, r.Added, r.Unchanged, r.UnmatchedCustomers);

        var review = BuildOrderContactsReview(parsed.Skipped, r.UnmatchedSample, r.UnmatchedCustomers);
        return (new
        {
            file = filename, type = "ordercontacts", parsed = parsed.Rows.Count,
            matched = r.MatchedRows, added = r.Added, unchanged = r.Unchanged,
            unmatchedCustomers = r.UnmatchedCustomers, unmatchedSample = r.UnmatchedSample,
            skippedForReview = review.Count, warnings = parsed.Warnings,
        }, review);
    }

    private async Task<(object result, IReadOnlyList<CustomerImportService.SkippedItem> skipped)>
        DiffOrderContactsAsync(string filename, string csv, CancellationToken ct)
    {
        var parsed = importer.ParseContactsFromOrders(csv);
        if (parsed.Rows.Count == 0 && parsed.Warnings.Count > 0)
            return (new { file = filename, type = "ordercontacts", dryRun = true, parsed = 0,
                          skippedForReview = parsed.Skipped.Count, warnings = parsed.Warnings }, parsed.Skipped);

        var diff   = await sp.DiffContactsExistingOnlyAsync(parsed.Rows, ct);
        var review = BuildOrderContactsReview(parsed.Skipped, diff.UnmatchedSample, diff.UnmatchedCustomers);
        return (new
        {
            file = filename, type = "ordercontacts", dryRun = true, parsed = parsed.Rows.Count,
            matched = diff.MatchedRows, rowsToAdd = diff.RowsToAdd, rowsUnchanged = diff.RowsUnchanged,
            // Add-quality breakdown so a large rowsToAdd can be inspected before committing.
            samePhoneAdds = diff.SamePhoneAdds, sameNameNewPhoneAdds = diff.SameNameNewPhoneAdds,
            brandNewAdds = diff.BrandNewAdds, samePhoneSample = diff.SamePhoneSample,
            unmatchedCustomers = diff.UnmatchedCustomers, unmatchedSample = diff.UnmatchedSample,
            skippedForReview = review.Count, warnings = parsed.Warnings,
        }, review);
    }

    /// <summary>Folds parse rejects (no phone / no name) AND unmatched customer names into review-file
    /// rows, so everything needing a human lands in _skipped_review_{ts}.csv.</summary>
    private static IReadOnlyList<CustomerImportService.SkippedItem> BuildOrderContactsReview(
        IReadOnlyList<CustomerImportService.SkippedItem> parseSkipped,
        IReadOnlyList<string> unmatchedSample, int unmatchedTotal)
    {
        var list = new List<CustomerImportService.SkippedItem>(parseSkipped);
        foreach (var name in unmatchedSample)
            list.Add(new CustomerImportService.SkippedItem(name, "customer not in Customers list (contacts skipped)"));
        if (unmatchedTotal > unmatchedSample.Count)
            list.Add(new CustomerImportService.SkippedItem(
                $"(+{unmatchedTotal - unmatchedSample.Count} more unmatched)",
                "customer not in Customers list (contacts skipped)"));
        return list;
    }

    // ── Customer Info (enrichment of existing Customers records) ──────────────

    private async Task<(object result, IReadOnlyList<CustomerImportService.SkippedItem> skipped)>
        ProcessCustomerInfoAsync(string filename, string csv, CancellationToken ct)
    {
        var parsed = importer.ParseCustomerInfo(csv);
        if (parsed.Rows.Count == 0 && parsed.Warnings.Count > 0)
            return (new { file = filename, type = "customerinfo", parsed = 0,
                          warnings = parsed.Warnings }, parsed.Skipped);

        var r = await sp.EnrichCustomersAsync(parsed.Rows, ct);
        log.LogInformation(
            "[Import] {File} customerinfo: parsed={P} matched={M} updated={U} unchanged={C} unmatched={X}",
            filename, parsed.Rows.Count, r.Matched, r.Updated, r.Unchanged, r.Unmatched);

        var review = BuildCustomerInfoReview(parsed.Skipped, r.UnmatchedSample, r.Unmatched);
        return (new
        {
            file = filename, type = "customerinfo", parsed = parsed.Rows.Count,
            matched = r.Matched, updated = r.Updated, unchanged = r.Unchanged,
            unmatchedCandidates = r.Unmatched, skippedInactive = r.SkippedInactive,
            unmatchedSample = r.UnmatchedSample,
            skippedForReview = review.Count, warnings = parsed.Warnings,
        }, review);
    }

    private async Task<(object result, IReadOnlyList<CustomerImportService.SkippedItem> skipped)>
        DiffCustomerInfoAsync(string filename, string csv, CancellationToken ct)
    {
        var parsed = importer.ParseCustomerInfo(csv);
        if (parsed.Rows.Count == 0 && parsed.Warnings.Count > 0)
            return (new { file = filename, type = "customerinfo", dryRun = true, parsed = 0,
                          warnings = parsed.Warnings }, parsed.Skipped);

        var diff   = await sp.DiffCustomerInfoAsync(parsed.Rows, ct);
        var review = BuildCustomerInfoReview(parsed.Skipped, diff.UnmatchedSample, diff.Unmatched);
        return (new
        {
            file = filename, type = "customerinfo", dryRun = true, parsed = parsed.Rows.Count,
            matched = diff.Matched, toUpdate = diff.ToUpdate, unchanged = diff.Unchanged,
            unmatchedCandidates = diff.Unmatched, skippedInactive = diff.SkippedInactive,
            unmatchedSample = diff.UnmatchedSample,
            infoUpdateSamples = diff.UpdateSamples.Select(s => new { name = s.Name, changedFields = s.ChangedFields }),
            skippedForReview = review.Count, warnings = parsed.Warnings,
        }, review);
    }

    /// <summary>Folds parse oddities (dupes / unparseable cells) AND unmatched "candidate to add" names
    /// into review-file rows, so everything that needs a human lands in _skipped_review_{ts}.csv.</summary>
    private static IReadOnlyList<CustomerImportService.SkippedItem> BuildCustomerInfoReview(
        IReadOnlyList<CustomerImportService.SkippedItem> parseSkipped,
        IReadOnlyList<string> unmatchedSample, int unmatchedTotal)
    {
        var list = new List<CustomerImportService.SkippedItem>(parseSkipped);
        foreach (var name in unmatchedSample)
            list.Add(new CustomerImportService.SkippedItem(name, "not in Customers list (candidate to add)"));
        if (unmatchedTotal > unmatchedSample.Count)
            list.Add(new CustomerImportService.SkippedItem(
                $"(+{unmatchedTotal - unmatchedSample.Count} more unmatched)",
                "not in Customers list (candidate to add)"));
        return list;
    }

    private static string CsvQuote(string value) =>
        value.Contains(',') || value.Contains('"') || value.Contains('\n')
            ? $"\"{value.Replace("\"", "\"\"")}\""
            : value;
}
