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
///   POST /api/analysis/catalog/tokenize — AI-tokenise catalog.json -> catalog-tokens.json (resumable)
///   POST /api/analysis/sli/tokenize     — AI-tokenise sli-sample.json -> sli-tokens.json (resumable)
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

    // Static tokenisation instructions sent as a cached system prompt.
    // Cached by the Anthropic prompt-caching API (beta header set in constructor).
    // Minimum cacheable: 1024 tokens for Sonnet, 2048 tokens for Haiku.
    private static readonly string TokeniseSystemPrompt = """
        Extract structured product attributes from metal product names.
        The user will send a numbered list of product names. Return a JSON array with one object per product, in the same order.

        Each object must have exactly these fields (null if unspecified):
        - metal: one of aluminum, steel, stainless, copper, brass, bronze, titanium, nickel, plastic -- or null
          Steel indicators: ASTM grades A36, A53, A106, A135, A179, A311, A333, A500, A512, A513, A519, A572, A992;
          terms "black pipe", "black steel", "carbon steel", "hot rolled", "cold rolled", "structural".
          Aluminum indicators: alloy series 1xxx/2xxx/3xxx/5xxx/6xxx/7xxx (e.g. 6061, 5052, 3003, 2024, 7075);
          terms "aluminum" or "aluminium".
          Stainless indicators: alloy numbers 301, 303, 304, 309, 310, 316, 317, 321, 347;
          terms "stainless", "SS", "inox".
          Copper alloys: brass (Cu+Zn), bronze (Cu+Sn), copper -- check the product name for the metal word.
          If a structural or pipe product carries an ASTM grade (e.g. A500, A513) with no other metal clue, default to steel.
        - alloy: grade or series as lowercase string ("6061", "304", "a36", "1018", "a500") -- or null
        - temper: temper designation as lowercase ("t6511", "t651", "h32", "h14") -- or null
        - shape: one of flatbar, roundbar, squarebar, hexbar, sheet, plate, angle, channel, tube_round, tube_square, tube_rect, pipe, wideflange, beam_s, coil, strip, rod, wire, expanded, grating, treadplate -- or null
          HSS (Hollow Structural Section): square cross-section -> tube_square; rectangular -> tube_rect; round -> tube_round.
          "Structural tube" follows the same rule. "SQ" or "square tube" -> tube_square.
          ASTM A500 = structural HSS (cold-formed welded/seamless, load-bearing) -> alloy=a500.
          ASTM A513 = mechanical tubing (ERW, precision tolerances, machined parts, not structural) -> alloy=a513.
          A500 and A513 are mutually exclusive; use whichever appears in the product name as the alloy token.
        - dims: decimal inches, shape-specific rules below -- or null
        - conditions: array of applicable terms from the list below -- empty [] if none
          Valid conditions: hot_rolled, cold_rolled, cold_drawn, stress_proof, galvanized, anodized, seamless, welded,
          dom, polished, drawn, extruded, key_stock, perforated, sch5, sch10, sch40, sch80, sch160
          ERW (Electric Resistance Welded) and HFW (High Frequency Welded) map to the welded condition.
          For pipe: always include exactly one schedule condition (sch5, sch10, sch40, sch80, or sch160).

        Dimensional ordering convention (industry standard):
          For angles, square tubes, rectangular tubes, tees, and zees: list dimensions from LARGEST to SMALLEST;
          the LAST dimension is always the wall thickness or material thickness.
          Example: "6x4x3/8 angle" = 6" leg x 4" leg x 0.375" wall -> "6.000x4.000x0.375"
          Example: "4x2x.188 rect tube" = 4" width x 2" height x 0.188" wall -> "4.000x2.000x0.188"
          EXCEPTION -- flatbar: thickness listed FIRST, then width.
          Example: "1/4 x 1-1/2 flat bar" -> "0.250x1.500"

        Dims rules -- use decimal inches; always ignore cut length and panel size dimensions:
          - sheet, plate, coil, strip: thickness only -> "0.050"  (ignore panel/cut size such as 48x120)
          - flatbar: thickness x width -> "0.250x1.500"  (thickness first -- exception to the ordering convention)
          - roundbar, rod, wire: diameter -> "0.500"
          - squarebar, hexbar: size -> "1.000"
          - angle: larger_leg x smaller_leg x wall -> "6.000x4.000x0.375"  (larger leg first, wall last)
          - channel: depth x weight-per-foot -> "6x13"  (e.g. C6x13, MC6x12, MC6x12#)
          - tube_square: size x size x wall -> "4.000x4.000x0.188"
          - tube_rect: larger_dim x smaller_dim x wall -> "4.000x2.000x0.188"  (larger first, wall last)
          - tube_round: OD x wall -> "2.000x0.120"
          - wideflange, beam_s: depth x weight-per-foot -> "8x15"  (e.g. W8x15, W8x15#, S12x40)
          - pipe: nominal ID only -- do NOT include OD, wall, or length -> "1.0"
            Always include one schedule condition. Resolve schedule abbreviations:
              Sch / SCH / Sched = Schedule
              STD or Standard = sch40 (for most nominal sizes up to 8")
              XH or Extra Heavy = sch80
              XXH or Double Extra Heavy = sch160
            If wall thickness is given instead of explicit schedule, use the NPS lookup table (tolerance +/-0.005"):
              Nominal | Sch 5  | Sch 10 | Sch 40/STD | Sch 80/XH | Sch 160
               0.500  |   --   |   --   |   0.109    |   0.147   |  0.187
               0.750  |   --   |   --   |   0.113    |   0.154   |  0.219
               1.000  |   --   |   --   |   0.133    |   0.179   |  0.250
               1.250  |   --   |   --   |   0.140    |   0.191   |  0.250
               1.500  |   --   |   --   |   0.145    |   0.200   |  0.281
               2.000  |   --   |   --   |   0.154    |   0.218   |  0.344
               2.500  |   --   |   --   |   0.203    |   0.276   |  0.375
               3.000  |   --   |   --   |   0.216    |   0.300   |  0.438
               4.000  |   --   | 0.120  |   0.237    |   0.337   |  0.531
               6.000  | 0.109  | 0.134  |   0.280    |   0.432   |  0.719
               8.000  | 0.109  | 0.148  |   0.322    |   0.500   |  0.906
              10.000  | 0.134  | 0.165  |   0.365    |   0.594   |  1.125
              12.000  | 0.156  | 0.180  |   0.406    |   0.688   |  1.312
            If OD and wall are given but no nominal size, reverse-lookup nominal from the NPS table.
            If wall thickness matches no known schedule, omit the schedule condition rather than guessing.

        Gauge to decimal conversion — select the correct table for the metal and shape context:
          Steel sheet (uncoated) — US Standard / Manufacturer's Standard gauge:
            3ga=0.239  4ga=0.224  5ga=0.209  6ga=0.194  7ga=0.179  8ga=0.164  9ga=0.150
            10ga=0.135  11ga=0.120  12ga=0.105  13ga=0.090  14ga=0.075  15ga=0.067  16ga=0.060
            17ga=0.054  18ga=0.048  19ga=0.042  20ga=0.036  22ga=0.030  24ga=0.024  26ga=0.018
          Galvanized steel sheet — Galvanized Sheet gauge (zinc coating adds ~0.003-0.005" vs bare steel):
            8ga=0.168  10ga=0.138  11ga=0.123  12ga=0.108  13ga=0.093  14ga=0.079
            16ga=0.064  18ga=0.052  20ga=0.040  22ga=0.034  24ga=0.028  26ga=0.022
          Stainless steel sheet — stainless gauge (NOT the same as carbon steel gauge; different thicknesses):
            7ga=0.188  8ga=0.172  9ga=0.156  10ga=0.141  11ga=0.125  12ga=0.109  13ga=0.094
            14ga=0.078  16ga=0.063  18ga=0.050  20ga=0.038  22ga=0.031  24ga=0.025  26ga=0.019
          Aluminum sheet — B&S (Brown and Sharpe) gauge:
            4ga=0.204  6ga=0.162  8ga=0.129  10ga=0.102  11ga=0.091  12ga=0.081  14ga=0.064
            16ga=0.051  18ga=0.040  20ga=0.032  22ga=0.025  24ga=0.020  26ga=0.016
          Steel tube wall — Birmingham Wire Gauge (BWG), used for mechanical tube (A513) wall thickness:
            7ga=0.180  8ga=0.165  9ga=0.148  10ga=0.134  11ga=0.120  12ga=0.109
            13ga=0.095  14ga=0.083  15ga=0.072  16ga=0.065  18ga=0.049  20ga=0.035

        Weight per foot as a metal check for pipe (when no metal is stated):
          Steel density ~0.283 lb/in3; aluminum ~0.098 lb/in3 (roughly 1/3 of steel); stainless ~0.289 lb/in3.
          If lbs/ft is provided alongside a known nominal pipe size, compare to expected steel weight.
          If actual lbs/ft is close to steel weight -> metal=steel. If roughly 1/3 of steel weight -> metal=aluminum.
          Example: 4" Sch 10 steel pipe weighs ~5.6 lb/ft; aluminum of same dims weighs ~1.9 lb/ft.

        Reference examples:
          "1045 Ground Shafting 1.000" -> {metal:steel, alloy:1045, shape:roundbar, dims:"1.000", conditions:[cold_drawn]}
          "Aluminum Sheet 5052-H32 0.063 x 48 x 144" -> {metal:aluminum, alloy:5052, temper:h32, shape:sheet, dims:"0.063"}
          "304 SS Angle 2x2x3/16" -> {metal:stainless, alloy:304, shape:angle, dims:"2.000x2.000x0.188"}
          "Hot Rolled Flat Bar A36 3/8 x 1-1/2" -> {metal:steel, alloy:a36, shape:flatbar, dims:"0.375x1.500"}
          "A135 ERW Pipe 4 Sch 10 x 21ft" -> {metal:steel, alloy:a135, shape:pipe, dims:"4.0", conditions:[welded,sch10]}
          "4 .120 Domestic A135 ERW Pipe 21ft 5.62lb/ft" -> {metal:steel, alloy:a135, shape:pipe, dims:"4.0", conditions:[welded,sch10]}
          (0.120" wall at 4" nominal = Sch 10 per NPS table; 5.62 lb/ft confirms steel)
          "HSS 4x2x.188 A500" -> {metal:steel, alloy:a500, shape:tube_rect, dims:"4.000x2.000x0.188", conditions:[]}
          "2 Sq Tube 11ga ERW A513" -> {metal:steel, alloy:a513, shape:tube_square, dims:"2.000x2.000x0.120", conditions:[welded]}
          (11ga BWG tube wall = 0.120"; A513 = mechanical tubing)
          "304 SS Sheet 11ga x 48 x 96" -> {metal:stainless, alloy:304, shape:sheet, dims:"0.125"}
          (11ga stainless = 0.125" -- different from 11ga carbon steel 0.120")
          "16ga Galvanized Sheet 4x8" -> {metal:steel, shape:sheet, dims:"0.064", conditions:[galvanized]}
          (16ga galvanized = 0.064" -- different from 16ga bare steel 0.060")
          "W8x15 Wide Flange A992" -> {metal:steel, alloy:a992, shape:wideflange, dims:"8x15"}
          "316L SS Seamless Tube 1.5 OD x .109 Wall" -> {metal:stainless, alloy:316, shape:tube_round, dims:"1.500x0.109", conditions:[seamless]}
          "6063-T52 Aluminum Pipe 1.25 Sch 40" -> {metal:aluminum, alloy:6063, temper:t52, shape:pipe, dims:"1.25", conditions:[sch40]}
          "Galvanized Pipe 1 STD" -> {metal:steel, shape:pipe, dims:"1.0", conditions:[galvanized,sch40]}
          "C6x13 Structural Channel" -> {metal:steel, shape:channel, dims:"6x13"}
          "Polycarbonate Lexan Clear Sheet 0.500 x 48 x 96" -> {metal:plastic, shape:sheet, dims:"0.500"}
          "Brass Round Bar 360 1.500" -> {metal:brass, alloy:360, shape:roundbar, dims:"1.500"}
          "6x4x3/8 Structural Angle A36" -> {metal:steel, alloy:a36, shape:angle, dims:"6.000x4.000x0.375"}
          "Hot Rolled Pipe 4.000 Schedule 10 (OD 4.500 - Wall 0.120)" -> {metal:steel, shape:pipe, dims:"4.0", conditions:[hot_rolled,sch10]}

        Respond with only the JSON array, no explanation or markdown fences.
        """;

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
        // Prompt-caching beta: static system prompt is cached across batches.
        // Activates at 1024 tokens for Sonnet, 2048 for Haiku.
        _http.DefaultRequestHeaders.Add("anthropic-beta", "prompt-caching-2024-07-31");
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

    /// <summary>Snapshot ProductCatalogService in-memory cache -> catalog.json</summary>
    public object FetchCatalog()
    {
        var entries = _catalog.CachedEntries
            .Select(e => new CatalogEntry(e.Name, e.SearchKey))
            .ToList();

        EnsureDir();
        File.WriteAllText(CatalogPath,
            JsonSerializer.Serialize(entries, Json), Encoding.UTF8);

        _log.LogInformation("[Analysis] catalog.json written -- {Count} entries", entries.Count);
        return new { count = entries.Count, path = CatalogPath };
    }

    /// <summary>Snapshot SliCacheService rows with a known ProductSearchKey -> sli-sample.json</summary>
    public async Task<object> FetchSliAsync(CancellationToken ct)
    {
        // Ensure cache is warm -- will populate from disk or SP
        if (_sli.TryGet() is null)
            await _sli.PopulateAsync(force: false, ct);

        var rows = _sli.TryGet();
        if (rows is null)
            return new { error = "SLI cache is empty -- proxy may still be starting up" };

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

        _log.LogInformation("[Analysis] sli-sample.json written -- {Count} rows (from {Total} total)",
            sample.Count, rows.Count);
        return new { count = sample.Count, totalRows = rows.Count, path = SliPath };
    }

    // ── Tokenise ──────────────────────────────────────────────────────────────

    /// <summary>
    /// AI-tokenise the catalog or SLI source file.  Resumable -- entries already
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
            return new { error = $"{label}.json not found -- run fetch first" };

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
            return new { error = "catalog-tokens.json not found -- run catalog/tokenize first" };
        if (!File.Exists(SliTokensPath))
            return new { error = "sli-tokens.json not found -- run sli/tokenize first" };

        var catJson  = await File.ReadAllTextAsync(CatalogTokensPath, ct);
        var sliJson  = await File.ReadAllTextAsync(SliTokensPath, ct);

        var catalog   = JsonSerializer.Deserialize<List<ProductTokens>>(catJson, Json)!
                            .Where(t => !t.TokenizationFailed).ToList();
        var sliTokens = JsonSerializer.Deserialize<List<ProductTokens>>(sliJson, Json)!
                            .Where(t => !t.TokenizationFailed).ToList();

        // Load sli-sample.json once -- keyed by name (first entry wins for duplicates)
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

    // Conditions that are "exclusive" -- if the catalog entry carries one of these,
    // the supplier description must also mention it.  Prevents a plain sheet matching
    // a perforated sheet, a plain tube matching a DOM tube, or Sch 10 matching Sch 80, etc.
    private static readonly HashSet<string> ExclusiveConditions =
        new(StringComparer.OrdinalIgnoreCase)
        { "perforated", "expanded", "dom", "grating", "treadplate",
          "sch5", "sch10", "sch40", "sch80", "sch160" };

    private static (ProductTokens? match, double score, string? failReason)
        FindBestMatch(ProductTokens supplier, List<ProductTokens> catalog)
    {
        ProductTokens? best            = null;
        double         bestScore       = -1;
        string?        noMatchReason   = "no candidates passed gates";

        foreach (var cat in catalog)
        {
            // Asymmetric metal gate: if supplier names a metal, catalog must name one too.
            // Prevents non-metal catalog entries (plastic, etc.) from matching metal products
            // via dims alone (catalog TkMetal=null passes the symmetric gate but is wrong).
            if (supplier.TkMetal != null && cat.TkMetal == null)
            {
                noMatchReason = $"supplier is '{supplier.TkMetal}' but catalog metal unspecified";
                continue;
            }

            // Hard gates: if both sides specify a field and they differ -> skip
            if (!GatePasses(supplier.TkMetal,  cat.TkMetal,  out var r1)) { noMatchReason = r1; continue; }
            if (!GatePasses(supplier.TkShape,  cat.TkShape,  out var r2)) { noMatchReason = r2; continue; }
            if (!GatePasses(supplier.TkAlloy,  cat.TkAlloy,  out var r3)) { noMatchReason = r3; continue; }
            if (!GatePasses(supplier.TkTemper, cat.TkTemper, out var r4)) { noMatchReason = r4; continue; }

            // Exclusive-condition gate: catalog carries an exclusive tag the supplier
            // didn't mention -> not this product
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
        if (a is null || b is null) return true; // one side unspecified -> pass
        if (string.Equals(a, b, StringComparison.OrdinalIgnoreCase)) return true;
        reason = $"{a} vs {b}";
        return false;
    }

    private static bool DimsMatch(string? a, string? b, double tolerance = 0.05)
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

        // Static instructions are in the system message with cache_control so the Anthropic
        // API can cache them across batches (avoids re-sending ~1500 tokens every call).
        var body = JsonSerializer.Serialize(new
        {
            model,
            max_tokens = 2048,
            system = new[]
            {
                new
                {
                    type         = "text",
                    text         = TokeniseSystemPrompt,
                    cache_control = new { type = "ephemeral" }
                }
            },
            messages = new[]
            {
                new { role = "user", content = $"Products:\n{numbered}" }
            }
        }, Json);

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
