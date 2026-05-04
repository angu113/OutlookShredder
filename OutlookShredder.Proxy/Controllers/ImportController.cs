using Microsoft.AspNetCore.Mvc;
using OutlookShredder.Proxy.Services;

namespace OutlookShredder.Proxy.Controllers;

[ApiController]
[Route("api/import")]
public class ImportController(
    CustomerImportService importer,
    SharePointService     sp,
    IConfiguration        config,
    ILogger<ImportController> log) : ControllerBase
{
    /// <summary>
    /// POST /api/import/run — scans the import directory for CSV files, processes each one,
    /// and moves processed files to a timestamped subfolder.
    ///
    /// File type detection (first match wins):
    ///   - Filename contains "partner" or "bp"  → business partners
    ///   - Filename contains "contact"           → contacts
    ///   - Header contains "Popup Message"       → business partners
    ///   - Header contains "Contact Name" and "Customer Name" → contacts
    ///   - Otherwise: skipped with a warning
    ///
    /// If any BP rows were filtered out by a phrase match (duplicate, do not use, dupe),
    /// a review file is written to Import\_skipped_review_{timestamp}.csv for manual inspection.
    ///
    /// Returns a per-file summary array.
    /// </summary>
    [HttpPost("run")]
    public async Task<IActionResult> Run(CancellationToken ct)
    {
        var importDir = ResolveImportDir();

        Directory.CreateDirectory(importDir);
        var processedDir = Path.Combine(importDir, "processed");
        Directory.CreateDirectory(processedDir);

        var files = Directory.GetFiles(importDir, "*.csv");
        if (files.Length == 0)
            return Ok(new { importDir, message = "No CSV files found in import directory.", files = Array.Empty<object>() });

        var timestamp    = DateTimeOffset.Now.ToString("yyyyMMdd_HHmmss");
        var results      = new List<object>();
        // Accumulate phrase-matched skipped BPs across all files for the review log
        var reviewRows   = new List<(string SourceFile, string BpName, string Reason)>();

        foreach (var filePath in files)
        {
            var filename = Path.GetFileName(filePath);
            log.LogInformation("[Import] Processing {File}", filename);

            string csv;
            try { csv = await System.IO.File.ReadAllTextAsync(filePath, ct); }
            catch (Exception ex)
            {
                results.Add(new { file = filename, error = ex.Message });
                continue;
            }

            var fileType = DetectFileType(filename, csv);
            object result;

            try
            {
                (result, var skipped) = fileType switch
                {
                    "partners" => await ProcessPartnersAsync(filename, csv, ct),
                    "contacts" => await ProcessContactsAsync(filename, csv, ct),
                    _          => ((object)new { file = filename, skipped = true, reason = "Could not detect file type from filename or headers." },
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

            // Move to processed/ regardless of outcome so the file isn't re-processed
            var dest = Path.Combine(processedDir, $"{timestamp}_{filename}");
            try { System.IO.File.Move(filePath, dest, overwrite: true); }
            catch (Exception ex)
            {
                log.LogWarning(ex, "[Import] Could not move {File} to processed/", filename);
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

        return Ok(new { importDir, reviewFile, skippedForReview = reviewRows.Count, files = results });
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
        if (fn.Contains("partner") || fn.Contains(" bp") || fn.StartsWith("bp") || fn.Contains("_bp"))
            return "partners";
        if (fn.Contains("contact"))
            return "contacts";

        // Header fallback — read first line
        var firstLine = csv.Split('\n', 2)[0].ToLowerInvariant();
        if (firstLine.Contains("popup message"))
            return "partners";
        if (firstLine.Contains("contact name") && firstLine.Contains("customer name"))
            return "contacts";

        return "unknown";
    }

    private async Task<(object result, IReadOnlyList<CustomerImportService.SkippedItem> skipped)>
        ProcessPartnersAsync(string filename, string csv, CancellationToken ct)
    {
        var parsed = importer.ParsePartners(csv);
        if (parsed.Rows.Count == 0 && parsed.Warnings.Count > 0)
            return (new { file = filename, type = "partners", parsed = 0, warnings = parsed.Warnings,
                          skippedForReview = parsed.Skipped.Count }, parsed.Skipped);

        var (added, updated, skipped) = await sp.UpsertBusinessPartnersAsync(parsed.Rows, ct);
        log.LogInformation("[Import] {File} partners: parsed={P} added={A} updated={U} skipped={S}",
            filename, parsed.Rows.Count, added, updated, skipped);
        return (new { file = filename, type = "partners", parsed = parsed.Rows.Count,
                      added, updated, skipped, skippedForReview = parsed.Skipped.Count,
                      warnings = parsed.Warnings }, parsed.Skipped);
    }

    private async Task<(object result, IReadOnlyList<CustomerImportService.SkippedItem> skipped)>
        ProcessContactsAsync(string filename, string csv, CancellationToken ct)
    {
        var parsed = importer.ParseContacts(csv);
        if (parsed.Rows.Count == 0 && parsed.Warnings.Count > 0)
            return (new { file = filename, type = "contacts", parsed = 0,
                          warnings = parsed.Warnings }, parsed.Skipped);

        var (added, deleted) = await sp.UpsertContactsAsync(parsed.Rows, ct);
        log.LogInformation("[Import] {File} contacts: parsed={P} added={A} deleted={D}",
            filename, parsed.Rows.Count, added, deleted);
        return (new { file = filename, type = "contacts", parsed = parsed.Rows.Count,
                      added, deleted, warnings = parsed.Warnings }, parsed.Skipped);
    }

    private static string CsvQuote(string value) =>
        value.Contains(',') || value.Contains('"') || value.Contains('\n')
            ? $"\"{value.Replace("\"", "\"\"")}\""
            : value;
}
