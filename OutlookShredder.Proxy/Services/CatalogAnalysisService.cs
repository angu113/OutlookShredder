using System.Diagnostics;
using System.Net.Http.Headers;
using System.Text;
using System.Text.Json;
using System.Text.Json.Serialization;
using OutlookShredder.Proxy.Models;

namespace OutlookShredder.Proxy.Services;

/// <summary>
/// Disk-cached analysis pipeline for developing and testing the tokenised product
/// matching approach.  All heavy work (SP reads, AI tokenisation, match scoring) runs
/// once and the results are cached under %LOCALAPPDATA%\Shredder\Proxy\analysis-cache\.
/// No SharePoint writes happen here — commit is a separate, explicit step.
///
/// Endpoints (via AnalysisController):
///   POST /api/analysis/catalog/fetch    — snapshot ProductCatalogService into catalog.json
///   POST /api/analysis/sli/fetch        — snapshot SliCacheService rows (PSK != null) into sli-sample.json
///   POST /api/analysis/catalog/tokenize — AI-tokenise catalog.json → catalog-tokens.json (resumable)
///   POST /api/analysis/sli/tokenize     — AI-tokenise sli-sample.json → sli-tokens.json (resumable)
///   GET  /api/analysis/match-test       — score sli-tokens against catalog-tokens in-memory
///   GET  /api/analysis/status           — cache file sizes + counts
/// </summary>
public class CatalogAnalysisService
{
    private readonly IConfiguration                _config;
    private readonly ILogger<CatalogAnalysisService> _log;
    private readonly ProductCatalogService         _catalog;
    private readonly SliCacheService               _sli;
    private readonly HttpClient                    _http;

    private static readonly JsonSerializerOptions Json = new()
    {
        WriteIndented        = false,
        PropertyNamingPolicy = JsonNamingPolicy.CamelCase,
        DefaultIgnoreCondition = JsonIgnoreCondition.WhenWritingNull
    };

    private string CacheDir => Path.Combine(
        Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
        "Shredder", "Proxy", "analysis-cache");

    private string CatalogPath      => Path.Combine(CacheDir, "catalog.json");
    private string SliPath          => Path.Combine(CacheDir, "sli-sample.json");
    private string CatalogTokensPath=> Path.Combine(CacheDir, "catalog-tokens.json");
    private string SliTokensPath    => Path.Combine(CacheDir, "sli-tokens.json");
    private string MatchResultsPath => Path.Combine(CacheDir, "match-test-results.json");

    private const string ApiUrl      = "https://api.anthropic.com/v1/messages";
    private const string ApiVersion  = "2023-06-01";
    private const int    BatchSize   = 25;

    public CatalogAnalysisService(
        IConfiguration config,
        ILogger<CatalogAnalysisService> log,
        ProductCatalogService catalog,
        SliCacheService sli)
    {
        _config  = config;
        _log     = log;
        _catalog = catalog;
        _sli     = sli;

        _http = new HttpClient { Timeout = TimeSpan.FromSeconds(60) };
        _http.DefaultRequestHeaders.Accept.Add(
            new MediaTypeWithQualityHeaderValue("application/json"));
    }

    // ── Status ────────────────────────────────────────────────────────────────

    public object GetStatus()
    {
        static object FileInfo(string path)
        {
            if (!File.Exists(path)) return new { exists = false };
            var fi = new FileInfo(path);
            return new { exists = true, bytes = fi.Length, modified = fi.LastWriteTimeUtc };
        }
        return new
        {
            catalogJson       = FileInfo(CatalogPath),
            sliJson           = FileInfo(SliPath),
            catalogTokensJson = FileInfo(CatalogTokensPath),
            sliTokensJson     = FileInfo(SliTokensPath),
            matchResultsJson  = FileInfo(MatchResultsPath),
        };
    }

    // ── Fetch ─────────────────────────────────────────────────────────────────

    /// <summary>Snapshot ProductCatalogService in-memory cache → catalog.json</summary>
    public object FetchCatalog()
    {
        var entries = _catalog.CachedEntries
            .Select(e => new CatalogEntry(e.Name, e.SearchKey))
            .ToList();

        EnsureDir();
        File.WriteAllText(CatalogPath,
            JsonSerializer.Serialize(entries, Json), Encoding.UTF8);

        _log.LogInformation("[Analysis] catalog.json written — {Count} entries", entries.Count);
        return new { count = entries.Count, path = CatalogPath };
    }

    /// <summary>Snapshot SliCacheService rows with a known ProductSearchKey → sli-sample.json</summary>
    public async Task<object> FetchSliAsync(CancellationToken ct)
    {
        // Ensure cache is warm — will populate from disk or SP
        if (_sli.TryGet() is null)
            await _sli.PopulateAsync(force: false, ct);

        var rows = _sli.TryGet();
        if (rows is null)
            return new { error = "SLI cache is empty — proxy may still be starting up" };

        var sample = rows
            .Select(d => new SliEntry(
                ProductName:        GetStr(d, "ProductName") ?? "",
                ProductSearchKey:   GetStr(d, "CatalogProductSearchKey", "ProductSearchKey"),
                CatalogProductName: GetStr(d, "CatalogProductName"),
                SupplierName:       GetStr(d, "SupplierName"),
                RfqId:              GetStr(d, "JobReference")))
            .Where(e => !string.IsNullOrWhiteSpace(e.ProductName)
                     && !string.IsNullOrWhiteSpace(e.ProductSearchKey))
            .ToList();

        EnsureDir();
        File.WriteAllText(SliPath,
            JsonSerializer.Serialize(sample, Json), Encoding.UTF8);

        _log.LogInformation("[Analysis] sli-sample.json written — {Count} rows (from {Total} total)",
            sample.Count, rows.Count);
        return new { count = sample.Count, totalRows = rows.Count, path = SliPath };
    }

    // ── Tokenise ──────────────────────────────────────────────────────────────

    /// <summary>
    /// AI-tokenise the catalog or SLI source file.  Resumable — entries already
    /// present in the tokens file are skipped.  Saves progress after every batch
    /// so interrupting mid-run doesn't lose work.
    /// clearShapes: if non-empty, removes entries with those TkShape values from the
    /// existing tokens file before starting, forcing them to be re-tokenised.
    /// </summary>
    public async Task<object> TokenizeAsync(string target, string[] clearShapes, CancellationToken ct)
    {
        var (sourcePath, tokensPath, label) = target == "sli"
            ? (SliPath, SliTokensPath, "sli")
            : (CatalogPath, CatalogTokensPath, "catalog");

        if (!File.Exists(sourcePath))
            return new { error = $"{label}.json not found — run fetch first" };

        var apiKey = _config["Anthropic:ApiKey"];
        if (string.IsNullOrWhiteSpace(apiKey))
            return new { error = "Anthropic:ApiKey not configured" };

        var model = _config["Claude:AnalysisModel"] ?? "claude-haiku-4-5-20251001";

        // Load source names
        var sourceJson = await File.ReadAllTextAsync(sourcePath, ct);
        List<(string Name, string? SearchKey)> source;
        if (target == "sli")
        {
            var sliEntries = JsonSerializer.Deserialize<List<SliEntry>>(sourceJson, Json)!;
            source = sliEntries.Select(e => (e.ProductName, e.ProductSearchKey)).ToList();
        }
        else
        {
            var catEntries = JsonSerializer.Deserialize<List<CatalogEntry>>(sourceJson, Json)!;
            source = catEntries.Select(e => (e.Name, e.SearchKey)).ToList();
        }

        // Load already-tokenised entries (resume support)
        var existing = new Dictionary<string, ProductTokens>(StringComparer.OrdinalIgnoreCase);
        if (File.Exists(tokensPath))
        {
            var existingJson = await File.ReadAllTextAsync(tokensPath, ct);
            var existingList = JsonSerializer.Deserialize<List<ProductTokens>>(existingJson, Json)
                               ?? [];
            foreach (var t in existingList.Where(t => !t.TokenizationFailed))
                existing[t.Name] = t;

            // Drop entries whose shape is in the clear-list so they get re-tokenised
            if (clearShapes.Length > 0)
            {
                var toClear = existing.Values
                    .Where(t => t.TkShape != null &&
                                clearShapes.Contains(t.TkShape, StringComparer.OrdinalIgnoreCase))
                    .Select(t => t.Name)
                    .ToList();
                foreach (var n in toClear) existing.Remove(n);
                _log.LogInformation("[Analysis] {Label} clear-shapes: removed {Count} entries for shapes [{Shapes}]",
                    label, toClear.Count, string.Join(", ", clearShapes));
            }

            _log.LogInformation("[Analysis] {Label} resume: {Done}/{Total} already tokenised",
                label, existing.Count, source.Count);
        }

        var pending = source.Where(s => !existing.ContainsKey(s.Name)).ToList();
        var sw       = Stopwatch.StartNew();
        int processed = 0, failed = 0;

        for (int i = 0; i < pending.Count && !ct.IsCancellationRequested; i += BatchSize)
        {
            var batch = pending.Skip(i).Take(BatchSize).ToList();
            var results = await TokeniseBatchAsync(batch.Select(b => b.Name).ToList(),
                apiKey, model, ct);

            for (int j = 0; j < batch.Count; j++)
            {
                var (name, searchKey) = batch[j];
                if (j < results.Count && results[j] is not null)
                {
                    results[j]!.Name       = name;
                    results[j]!.SearchKey  = searchKey;
                    results[j]!.TokenizedAt = DateTime.UtcNow;
                    existing[name] = results[j]!;
                    processed++;
                }
                else
                {
                    existing[name] = new ProductTokens
                    {
                        Name = name, SearchKey = searchKey,
                        TokenizationFailed = true, TokenizedAt = DateTime.UtcNow
                    };
                    failed++;
                }
            }

            // Save progress after every batch
            var snapshot = existing.Values.ToList();
            await File.WriteAllTextAsync(tokensPath,
                JsonSerializer.Serialize(snapshot, Json), Encoding.UTF8, ct);

            _log.LogInformation("[Analysis] {Label} tokenise: {Done}/{Total} ({Fail} failed)",
                label, processed + failed, pending.Count, failed);
        }

        return new
        {
            processed,
            failed,
            skipped    = existing.Count - processed - failed,
            total      = source.Count,
            elapsed    = sw.Elapsed.ToString(@"mm\:ss"),
            path       = tokensPath
        };
    }

    // ── Match test ────────────────────────────────────────────────────────────

    /// <summary>
    /// Scores each sli-tokens entry against catalog-tokens in-memory.
    /// Returns a MatchTestRun with per-row details.  No network calls.
    /// </summary>
    public async Task<object> RunMatchTestAsync(int limit, CancellationToken ct)
    {
        if (!File.Exists(CatalogTokensPath))
            return new { error = "catalog-tokens.json not found — run catalog/tokenize first" };
        if (!File.Exists(SliTokensPath))
            return new { error = "sli-tokens.json not found — run sli/tokenize first" };

        var catJson  = await File.ReadAllTextAsync(CatalogTokensPath, ct);
        var sliJson  = await File.ReadAllTextAsync(SliTokensPath, ct);

        var catalog   = JsonSerializer.Deserialize<List<ProductTokens>>(catJson, Json)!
                            .Where(t => !t.TokenizationFailed).ToList();
        var sliTokens = JsonSerializer.Deserialize<List<ProductTokens>>(sliJson, Json)!
                            .Where(t => !t.TokenizationFailed).ToList();

        // Load sli-sample.json once — keyed by name (first entry wins for duplicates)
        Dictionary<string, SliEntry> sliLookup = [];
        if (File.Exists(SliPath))
        {
            var rawSli   = await File.ReadAllTextAsync(SliPath, ct);
            var allSli   = JsonSerializer.Deserialize<List<SliEntry>>(rawSli, Json)!;
            foreach (var e in allSli)
                sliLookup.TryAdd(e.ProductName, e);  // first entry wins on duplicate names

            foreach (var t in sliTokens)
            {
                if (sliLookup.TryGetValue(t.Name, out var entry))
                    t.SearchKey = entry.ProductSearchKey;
            }
        }

        var run    = new MatchTestRun { RunAt = DateTime.UtcNow };
        var sample = limit > 0 ? sliTokens.Take(limit).ToList() : sliTokens;

        foreach (var supplier in sample)
        {
            var (match, score, failReason) = FindBestMatch(supplier, catalog);

            var mc = new MatchCase
            {
                ProductName         = supplier.Name,
                ExpectedSearchKey   = supplier.SearchKey,
                Score               = score,
                FailReason          = failReason
            };

            // Attach supplier/catalog name from pre-loaded lookup
            if (sliLookup.TryGetValue(supplier.Name, out var sliEntry))
            {
                mc.SupplierName        = sliEntry.SupplierName;
                mc.ExpectedCatalogName = sliEntry.CatalogProductName;
            }

            if (match is null)
            {
                mc.IsNoMatch = true;
                run.NoMatches++;
            }
            else
            {
                mc.ActualSearchKey   = match.SearchKey;
                mc.ActualCatalogName = match.Name;
                mc.IsHit             = string.Equals(match.SearchKey, supplier.SearchKey,
                                           StringComparison.OrdinalIgnoreCase);
                if (mc.IsHit) run.Hits++; else run.Misses++;
            }

            run.Cases.Add(mc);
        }

        run.Total = sample.Count;

        await File.WriteAllTextAsync(MatchResultsPath,
            JsonSerializer.Serialize(run, Json), Encoding.UTF8, ct);

        return new
        {
            total      = run.Total,
            hits       = run.Hits,
            misses     = run.Misses,
            noMatches  = run.NoMatches,
            hitRate    = $"{run.HitRate}%",
            path       = MatchResultsPath,
            // Top misses for quick inspection
            topMisses  = run.Cases
                .Where(c => !c.IsHit && !c.IsNoMatch)
                .Take(20)
                .Select(c => new
                {
                    c.ProductName,
                    c.SupplierName,
                    c.ExpectedSearchKey,
                    c.ExpectedCatalogName,
                    c.ActualSearchKey,
                    c.ActualCatalogName,
                    c.Score,
                    c.FailReason
                })
        };
    }

    // ── Matching algorithm ────────────────────────────────────────────────────

    // Conditions that are "exclusive" — if the catalog entry carries one of these,
    // the supplier description must also mention it.  Prevents a plain sheet matching
    // a perforated sheet, or a plain tube matching a DOM tube, etc.
    private static readonly HashSet<string> ExclusiveConditions =
        new(StringComparer.OrdinalIgnoreCase)
        { "perforated", "expanded", "dom", "grating", "treadplate" };

    private static (ProductTokens? match, double score, string? failReason)
        FindBestMatch(ProductTokens supplier, List<ProductTokens> catalog)
    {
        ProductTokens? best            = null;
        double         bestScore       = -1;
        string?        noMatchReason   = "no candidates passed gates";

        foreach (var cat in catalog)
        {
            // Hard gates: if both sides specify a field and they differ → skip
            if (!GatePasses(supplier.TkMetal,  cat.TkMetal,  out var r1)) { noMatchReason = r1; continue; }
            if (!GatePasses(supplier.TkShape,  cat.TkShape,  out var r2)) { noMatchReason = r2; continue; }
            if (!GatePasses(supplier.TkAlloy,  cat.TkAlloy,  out var r3)) { noMatchReason = r3; continue; }
            if (!GatePasses(supplier.TkTemper, cat.TkTemper, out var r4)) { noMatchReason = r4; continue; }

            // Exclusive-condition gate: catalog carries an exclusive tag the supplier
            // didn't mention → not this product
            var catExclusive = cat.TkConditions
                .Where(c => ExclusiveConditions.Contains(c)).ToArray();
            if (catExclusive.Length > 0)
            {
                bool supplierMentionsIt = catExclusive.Any(c =>
                    supplier.TkConditions.Contains(c, StringComparer.OrdinalIgnoreCase));
                if (!supplierMentionsIt)
                {
                    noMatchReason = $"catalog requires condition '{catExclusive[0]}'";
                    continue;
                }
            }

            // Scoring
            double score = 0;
            if (DimsMatch(supplier.TkDims, cat.TkDims))                        score += 3;
            if (ConditionsOverlap(supplier.TkConditions, cat.TkConditions))     score += 1;

            // Signal check: must score >= 1 (dims match OR conditions overlap),
            // OR both sides specify alloy AND temper (strong identity even without dims)
            bool bothHaveAlloyAndTemper =
                supplier.TkAlloy  != null && cat.TkAlloy  != null &&
                supplier.TkTemper != null && cat.TkTemper != null;

            if (score < 1 && !bothHaveAlloyAndTemper)
            {
                noMatchReason = "insufficient signal (no dims match, no conditions overlap, alloy+temper not both present)";
                continue;
            }

            if (score > bestScore)
            {
                bestScore    = score;
                best         = cat;
            }
        }

        // Only return a failReason when we have no match at all
        return (best, bestScore < 0 ? 0 : bestScore, best is null ? noMatchReason : null);
    }

    private static bool GatePasses(string? a, string? b, out string? reason)
    {
        reason = null;
        if (a is null || b is null) return true; // one side unspecified → pass
        if (string.Equals(a, b, StringComparison.OrdinalIgnoreCase)) return true;
        reason = $"{a} vs {b}";
        return false;
    }

    private static bool DimsMatch(string? a, string? b, double tolerance = 0.03)
    {
        var da = ParseDims(a);
        var db = ParseDims(b);
        if (da is null || db is null || da.Length != db.Length) return false;
        Array.Sort(da);
        Array.Sort(db);
        for (int i = 0; i < da.Length; i++)
        {
            double avg = (da[i] + db[i]) / 2.0;
            if (avg < 0.001) continue;
            if (Math.Abs(da[i] - db[i]) / avg > tolerance) return false;
        }
        return true;
    }

    private static double[]? ParseDims(string? dims)
    {
        if (string.IsNullOrWhiteSpace(dims)) return null;
        var parts = dims.Split('x', StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries);
        var result = new double[parts.Length];
        for (int i = 0; i < parts.Length; i++)
            if (!double.TryParse(parts[i], System.Globalization.NumberStyles.Any,
                System.Globalization.CultureInfo.InvariantCulture, out result[i]))
                return null;
        return result;
    }

    private static bool ConditionsOverlap(string[]? a, string[]? b)
    {
        if (a is null || a.Length == 0 || b is null || b.Length == 0) return false;
        return a.Any(c => b.Contains(c, StringComparer.OrdinalIgnoreCase));
    }

    // ── Claude batch tokenisation ─────────────────────────────────────────────

    private async Task<List<ProductTokens?>> TokeniseBatchAsync(
        List<string> names, string apiKey, string model, CancellationToken ct)
    {
        var numbered = string.Join("\n",
            names.Select((n, i) => $"{i + 1}. {n}"));

        var prompt = $"""
            Extract structured product attributes from these metal product names.
            Return a JSON array — one object per product in the same order as numbered.

            Each object must have exactly these fields (null if unspecified):
            - metal: one of aluminum, steel, stainless, copper, brass, bronze, titanium, nickel — or null
            - alloy: grade/series as lowercase string ("6061", "304", "a36", "1018") — or null
            - temper: temper code as lowercase string ("t6511", "t651", "h14") — or null
            - shape: one of flatbar, roundbar, squarebar, hexbar, sheet, plate, angle, channel, tube_round, tube_square, tube_rect, pipe, wideflange, beam_s, coil, strip, rod, wire, expanded, grating, treadplate — or null
            - dims: decimal inches, shape-specific rules below — or null
            - conditions: array from [hot_rolled, cold_rolled, cold_drawn, stress_proof, galvanized, anodized, seamless, welded, dom, polished, drawn, extruded, key_stock, perforated] — empty [] if none

            Dims rules — always ignore cut/length/panel dimensions:
            - sheet, plate, coil, strip: thickness ONLY → "0.050"  (ignore 48x120 panel size)
            - flatbar: thickness×width → "0.250x1.250"
            - roundbar, rod, wire: diameter → "0.500"
            - squarebar, hexbar: size → "1.000"
            - angle: leg1×leg2×thickness → "4.000x4.000x0.375"  (legs first, wall last)
            - channel: depth×weight-per-foot → "6x13"  (e.g. C6×13#, MC6×12#)
            - tube_square: width×width×wall → "4.000x4.000x0.188"
            - tube_rect: width×height×wall → "4.000x3.000x0.120"
            - tube_round: OD×wall → "2.000x0.120"
            - pipe: nominal size ONLY → "1.0"  (ignore OD and wall thickness)
            - wideflange, beam_s: depth×weight-per-foot → "8x15"  (e.g. W8×15#, S12×40#)

            Products:
            {numbered}

            Respond with only the JSON array, no explanation or markdown fences.
            """;

        var body = JsonSerializer.Serialize(new
        {
            model,
            max_tokens = 2048,
            messages   = new[] { new { role = "user", content = prompt } }
        });

        using var req = new HttpRequestMessage(HttpMethod.Post, ApiUrl);
        req.Headers.Add("x-api-key", apiKey);
        req.Headers.Add("anthropic-version", ApiVersion);
        req.Content = new StringContent(body, Encoding.UTF8, "application/json");

        HttpResponseMessage resp;
        try { resp = await _http.SendAsync(req, ct); }
        catch (Exception ex)
        {
            _log.LogWarning(ex, "[Analysis] Claude call failed for batch");
            return Enumerable.Repeat<ProductTokens?>(null, names.Count).ToList();
        }

        var raw = await resp.Content.ReadAsStringAsync(ct);
        if (!resp.IsSuccessStatusCode)
        {
            _log.LogWarning("[Analysis] Claude returned {Status}: {Body}", resp.StatusCode, raw[..Math.Min(200, raw.Length)]);
            return Enumerable.Repeat<ProductTokens?>(null, names.Count).ToList();
        }

        try
        {
            using var doc    = JsonDocument.Parse(raw);
            var       text   = doc.RootElement
                .GetProperty("content")[0]
                .GetProperty("text")
                .GetString() ?? "[]";

            // Strip any accidental markdown fences
            text = text.Trim();
            if (text.StartsWith("```")) text = text[(text.IndexOf('\n') + 1)..];
            if (text.EndsWith("```"))   text = text[..text.LastIndexOf("```")].TrimEnd();

            var tokenDtos = JsonSerializer.Deserialize<List<ProductTokens>>(text, Json)
                            ?? [];

            // Pad with nulls if Claude returned fewer entries than expected
            while (tokenDtos.Count < names.Count) tokenDtos.Add(null!);
            return tokenDtos.Take(names.Count).Select(t => (ProductTokens?)t).ToList();
        }
        catch (Exception ex)
        {
            _log.LogWarning(ex, "[Analysis] Failed to parse Claude batch response");
            return Enumerable.Repeat<ProductTokens?>(null, names.Count).ToList();
        }
    }

    // ── Helpers ───────────────────────────────────────────────────────────────

    private void EnsureDir()
        => Directory.CreateDirectory(CacheDir);

    private static string? GetStr(Dictionary<string, object?> d, params string[] keys)
    {
        foreach (var key in keys)
        {
            if (!d.TryGetValue(key, out var v)) continue;
            if (v is string s) return string.IsNullOrWhiteSpace(s) ? null : s;
            if (v is JsonElement je)
            {
                var str = je.ValueKind == JsonValueKind.String ? je.GetString() : je.ToString();
                return string.IsNullOrWhiteSpace(str) ? null : str;
            }
        }
        return null;
    }
}
