using System.Diagnostics;
using System.Net.Http.Headers;
using System.Security.Cryptography;
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
    private readonly IConfiguration                          _config;
    private readonly ILogger<CatalogAnalysisService>         _log;
    private readonly ProductCatalogService                   _catalog;
    private readonly SliCacheService                         _sli;
    private readonly SharePointService                       _sp;
    private readonly SupplierProductMappingsCacheService     _mappingsCache;
    private readonly HttpClient                              _http;

    // Live-tokenisation rate limit: at most 3 concurrent calls at any time.
    private readonly SemaphoreSlim _liveSemaphore = new(3, 3);
    // Round-robin counter for Claude Haiku / Gemini Flash alternation.
    private int _rrCounter;

    // Live tokenization progress (updated each batch; reset when a new run starts).
    private volatile string _tokenizeTarget = "";
    private int _tokenizeDone;
    private int _tokenizeTotal;

    // In-memory catalog token cache for MatchProductAsync — loaded once from disk and reused.
    // Null until first call; refreshed if the on-disk file is newer than the snapshot.
    private List<ProductTokens>? _catalogTokenCache;
    private DateTime             _catalogTokenCacheAt;

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
          Dual-certified stainless grades: "304/304L" or "304L/304" -> alloy=304 (not "304l").
          "316/316L" or "316L/316" -> alloy=316. The L variant is interchangeable; use the base grade.
          Similarly "317L" -> alloy=317, "321H" -> alloy=321.
        - temper: temper designation as lowercase ("t6511", "t651", "h32", "h14") -- or null
        - shape: one of flatbar, roundbar, squarebar, hexbar, sheet, plate, angle, channel, tube_round, tube_square, tube_rect, pipe, wideflange, beam_s, coil, strip, rod, wire, expanded, grating, treadplate -- or null
          CRITICAL: tread plate / diamond plate / checker plate / lug plate (any plate with a raised surface pattern) MUST use shape=treadplate, NOT shape=plate or shape=sheet. The word "tread", "diamond plate", or "checker" in the product name always means shape=treadplate.
          wideflange = W-shape steel beams only (W4x13, W8x31 etc.). Use beam_s for: (a) steel
          S-shape / American Standard I-beams (S4x7.7, S6x12.5 etc.) AND (b) aluminum I-beams
          / structural I-beams / wide flange I-beams in 6061 or 6063 (these follow S-shape dims).
          HSS (Hollow Structural Section): square cross-section -> tube_square; rectangular -> tube_rect; round -> tube_round.
          "Structural tube" follows the same rule. "SQ" or "square tube" -> tube_square.
          ASTM A500 = structural HSS (cold-formed welded/seamless, load-bearing) -> alloy=a500.
          ASTM A513 = mechanical tubing (ERW, precision tolerances, machined parts, not structural) -> alloy=a513.
          EXCEPTION: if both A500 and A513 appear together (e.g. "A500/A513") -> alloy=null.
          Dual listing means either standard is acceptable; alloy=null lets the matcher accept either.
        - dims: decimal inches, shape-specific rules below -- or null
        - conditions: array of applicable terms from the list below -- empty [] if none
          Valid conditions: hot_rolled, cold_rolled, cold_drawn, stress_proof, galvanized, galvanneal,
          anodized, seamless, welded, dom, polished, drawn, extruded, key_stock, perforated,
          bright_annealed, ar_400, ar_500, ar_600, weathering_steel, painted,
          brushed, mirror_finish, ornamental_180,
          round_corner, pno, tgp,
          sch5, sch10, sch40, sch80, sch160
          pno: pickled and oiled surface treatment (Pickled & Oiled, P&O, PNO, HPNO). Use ONLY when
            the product explicitly specifies pickling/oiling. Plain hot-rolled plate/sheet with no
            surface treatment specified is NOT pno.
          ERW (Electric Resistance Welded) and HFW (High Frequency Welded) map to the welded condition.
          galvanized: hot-dip zinc coating (G60/G90/HDG/HDGI); spangled silver surface.
          galvanneal: galvannealed — hot-dip zinc then annealed to create iron-zinc alloy coating; matte gray
            surface; product names often include "Galvanneal", "GA", "HDGA", "A60" coating. Do NOT use
            galvanized for galvanneal — they are different products.
          bright_annealed: BA finish on stainless tube/pipe; smooth reflective interior; product names
            include "Bright Annealed", "BA", "Bright Anneal". For stainless sheet/plate, BA refers to the
            annealed mill surface — only use bright_annealed for tube/pipe.
          ar_400: abrasion-resistant plate (Brinell ~400); product names include "AR400", "AR 400",
            "Hardox 400", "Quend 400".
          ar_500: abrasion-resistant plate (Brinell ~500); product names include "AR500", "AR 500",
            "Hardox 500", "Quend 500".
          ar_600: abrasion-resistant plate (Brinell ~600); product names include "AR600", "AR 600",
            "Hardox 600", "Quend 600".
          weathering_steel: atmospheric-corrosion-resistant steel; product names include "A606 Type 4",
            "A588", "Corten", "Cor-Ten", "Weathering Steel". ASTM A606 = sheet; A588 = plate and structural.
          painted: factory-applied paint or coating; product names include "Painted", "Coated", "Pre-painted".
          brushed: #4 brushed/satin stainless surface; product names include "#4", "#4 Finish",
            "#4 Brushed", "Brushed Finish", "Satin Finish".
          mirror_finish: #8 bright mirror-polished stainless surface; product names include "#8",
            "#8 Mirror", "Mirror Finish", "Mirror Polish", "Bright Polish".
          ornamental_180: 180-grit ornamental-polished stainless tube; product names include
            "Ornamental", "180 Grit", "Ornamental 180 Grit". Use this for stainless tube/pipe
            with an ornamental polish. Use polished for other polished stainless products.
          polished: generic polished/smooth surface on stainless or other metals; use when "polished"
            or "polished finish" appears and no more specific condition (brushed/mirror_finish/
            ornamental_180) applies.
          round_corner: aluminum tube (square or rectangular) with rounded outer corners, as opposed
            to the standard sharp-corner extrusion; product names include "Round Corner", "RC",
            "Rounded Corner". Do NOT use round_corner for standard/sharp-corner tube.
          tgp: precision ground and polished round bar — "Turned, Ground and Polished", "Ground
            Shafting", "TGP", "T.G.P.", "T.G.&P.", "Ground Shaft", "Bearing Quality Ground".
            Use for round bar that has been precision-ground to a tight tolerance with a smooth finish.
            Do NOT use tgp for plain cold-drawn or hot-rolled round bar that is not ground/polished.
          IMPORTANT — tread plate is NOT perforated: tread plate (diamond plate, checker plate)
            has a raised diamond or lug pattern embossed on the surface. It has NO holes.
            Never use the perforated condition for tread plate products. Use perforated only for
            products with actual holes punched through the material.
          For pipe: always include exactly one schedule condition (sch5, sch10, sch40, sch80, or sch160).

        Dimensional ordering convention (industry standard):
          For angles, square tubes, rectangular tubes, tees, and zees: list dimensions from LARGEST to SMALLEST;
          the LAST dimension is always the wall thickness or material thickness.
          Example: "6x4x3/8 angle" = 6" leg x 4" leg x 0.375" wall -> "6.000x4.000x0.375"
          Example: "4x2x.188 rect tube" = 4" width x 2" height x 0.188" wall -> "4.000x2.000x0.188"
          EXCEPTION -- flatbar: thickness listed FIRST, then width.
          Example: "1/4 x 1-1/2 flat bar" -> "0.250x1.500"

        Dims rules -- use decimal inches; always ignore cut length and panel size dimensions:
          - sheet, plate, coil: thickness only -> "0.050"  (ignore panel/cut size such as 48x120)
          - flatbar, strip: thickness x width -> "0.250x1.500"  (thickness first -- exception to the ordering convention)
            strip examples: "HR Strip 1/8 x 1-1/2" -> "0.125x1.500"; "Galv Strip 1/8 x 1-1/4" -> "0.125x1.250"
          - roundbar: diameter -> "0.500"  (smooth round bar / round rod: "round bar", "round rod", "ground shafting")
          - rod: diameter -> "0.500"  (deformed rebar ONLY: "rebar", "re-bar", "#4 rebar" — do NOT use rod for smooth round bar)
          - wire: diameter -> "0.500"
          - squarebar, hexbar: size -> "1.000"
          - angle: larger_leg x smaller_leg x wall -> "6.000x4.000x0.375"  (larger leg first, wall last)
          - channel: depth x weight-per-foot -> "6x13"  (e.g. C6x13, MC6x12, MC6x12#)
          - tube_square: size x size x wall -> "4.000x4.000x0.188"
          - tube_rect: larger_dim x smaller_dim x wall -> "4.000x2.000x0.188"  (larger first, wall last)
          - tube_round: OD x wall -> "2.000x0.120"
          - wideflange, beam_s: depth x weight-per-foot -> "8x15"  (e.g. W8x15, W8x15#, S12x40)
          - pipe: nominal size ONLY — NEVER use OD, wall thickness, or length as dims -> "1.0"
            CRITICAL: if the product says "2.375" OD" that is the OD of nominal 2" pipe. Use dims="2.0", not "2.375".
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
          "Steel Rect Tube A500/A513 4x2x.188" -> {metal:steel, alloy:null, shape:tube_rect, dims:"4.000x2.000x0.188", conditions:[welded]}
          (A500/A513 dual standard -> alloy=null)
          "2 Sq Tube 11ga ERW A513" -> {metal:steel, alloy:a513, shape:tube_square, dims:"2.000x2.000x0.120", conditions:[welded]}
          (11ga BWG tube wall = 0.120"; A513 = mechanical tubing)
          "A513 Square Tube 2x2 14ga x 24ft" -> {metal:steel, alloy:a513, shape:tube_square, dims:"2.000x2.000x0.083", conditions:[welded]}
          (14ga BWG tube wall = 0.083" -- NOT 0.120")
          "A513 Rect Tube 2x1 16ga x 24ft" -> {metal:steel, alloy:a513, shape:tube_rect, dims:"2.000x1.000x0.065", conditions:[welded]}
          (16ga BWG tube wall = 0.065" -- NOT 0.120")
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
          "Hot Rolled Sheet Pickled & Oiled 11 Ga x 48 x 96" -> {metal:steel, shape:sheet, dims:"0.120", conditions:[hot_rolled,pno]}
          "HR Plate P&O 0.250 x 48 x 120" -> {metal:steel, shape:plate, dims:"0.250", conditions:[hot_rolled,pno]}
          "1045 Ground Shafting 1.000" -> {metal:steel, alloy:1045, shape:roundbar, dims:"1.000", conditions:[cold_drawn,tgp]}
          "303 SS TGP Round Bar 0.500" -> {metal:stainless, alloy:303, shape:roundbar, dims:"0.500", conditions:[cold_drawn,tgp]}
          "3003-H22 Aluminum Diamond Tread Plate 0.125 x 48 x 96" -> {metal:aluminum, alloy:3003, temper:h22, shape:treadplate, dims:"0.125"}
          "6061-T6 Aluminum Tread Plate 0.250 x 48 x 96" -> {metal:aluminum, alloy:6061, temper:t6, shape:treadplate, dims:"0.250"}

        Respond with only the JSON array, no explanation or markdown fences.
        """;

    public CatalogAnalysisService(
        IConfiguration config,
        ILogger<CatalogAnalysisService> log,
        ProductCatalogService catalog,
        SliCacheService sli,
        SharePointService sp,
        SupplierProductMappingsCacheService mappingsCache)
    {
        _config        = config;
        _log           = log;
        _catalog       = catalog;
        _sli           = sli;
        _sp            = sp;
        _mappingsCache = mappingsCache;

        _http = new HttpClient { Timeout = TimeSpan.FromSeconds(60) };
        _http.DefaultRequestHeaders.Accept.Add(
            new MediaTypeWithQualityHeaderValue("application/json"));
        // Prompt-caching beta: static system prompt is cached across batches.
        // Activates at 1024 tokens for Sonnet, 2048 for Haiku.
        _http.DefaultRequestHeaders.Add("anthropic-beta", "prompt-caching-2024-07-31");
    }

    // ── Industry Dictionary ───────────────────────────────────────────────────

    private static readonly string DictionarySystemPrompt = """
        You are a metals industry expert. For each numbered term, write a concise plain-English definition
        and 1-2 short product-name examples showing how the term appears in practice.

        Term types for context:
        - alloy_series: a material composition identifier (e.g. 6061, 304, 1018, a500)
        - standard: an ASTM or dimensional standard (e.g. A500, A36, sch40)
        - condition: a processing or surface condition (e.g. hot_rolled, galvanized, welded)
        - temper: a heat treatment or work-hardening designation (e.g. T6511, H32, O)

        Return a JSON array with one object per input term, in the same order:
        [{"term":"...","definition":"One sentence max.","examples":"Product Name 1; Product Name 2"},...]

        Respond with only the JSON array — no markdown fences, no extra text.
        """;

    private record DictDefDto(string Term, string? Definition, string? Examples);

    private async Task<Dictionary<string, DictDefDto>> GenerateDefinitionsAsync(
        List<IndustryDictionaryEntry> entries, string apiKey, string model, CancellationToken ct)
    {
        var result = new Dictionary<string, DictDefDto>(StringComparer.OrdinalIgnoreCase);
        const int batchSize = 20;

        for (int i = 0; i < entries.Count; i += batchSize)
        {
            if (ct.IsCancellationRequested) break;
            var batch = entries.Skip(i).Take(batchSize).ToList();

            var numbered = string.Join("\n", batch.Select((e, idx) =>
                $"{idx + 1}. {e.Term} [{e.TermType}" +
                (e.AppliesTo is { Length: > 0 } a ? $", {a}" : "") + "]"));

            var body = JsonSerializer.Serialize(new
            {
                model,
                max_tokens = 2048,
                system = new[]
                {
                    new { type = "text", text = DictionarySystemPrompt,
                          cache_control = new { type = "ephemeral" } }
                },
                messages = new[] { new { role = "user", content = numbered } }
            }, Json);

            using var req = new HttpRequestMessage(HttpMethod.Post, ApiUrl);
            req.Headers.Add("x-api-key", apiKey);
            req.Headers.Add("anthropic-version", ApiVersion);
            req.Content = new StringContent(body, Encoding.UTF8, "application/json");

            HttpResponseMessage resp;
            try { resp = await _http.SendAsync(req, ct); }
            catch (Exception ex)
            {
                _log.LogWarning(ex, "[Analysis] Dict-def batch {I} failed", i / batchSize + 1);
                continue;
            }

            var raw = await resp.Content.ReadAsStringAsync(ct);
            if (!resp.IsSuccessStatusCode)
            {
                _log.LogWarning("[Analysis] Dict-def Claude {Status}: {Snippet}",
                    resp.StatusCode, raw[..Math.Min(200, raw.Length)]);
                continue;
            }

            try
            {
                using var doc = JsonDocument.Parse(raw);
                var text = doc.RootElement.GetProperty("content")[0]
                              .GetProperty("text").GetString() ?? "[]";
                text = text.Trim();
                if (text.StartsWith("```")) text = text[(text.IndexOf('\n') + 1)..];
                if (text.EndsWith("```"))   text = text[..text.LastIndexOf("```")].TrimEnd();

                var dtos = JsonSerializer.Deserialize<List<DictDefDto>>(text,
                    new JsonSerializerOptions { PropertyNameCaseInsensitive = true }) ?? [];

                foreach (var dto in dtos.Where(d => d.Term is { Length: > 0 }))
                    result[dto.Term] = dto;

                _log.LogInformation("[Analysis] Dict-def batch {B}: {N}/{Total} terms defined",
                    i / batchSize + 1, dtos.Count, batch.Count);
            }
            catch (Exception ex)
            {
                _log.LogWarning(ex, "[Analysis] Dict-def parse failed for batch {B}", i / batchSize + 1);
            }
        }

        return result;
    }

    /// <summary>
    /// Mine catalog-tokens.json + sli-tokens.json for unique alloy, condition, and temper values.
    /// Upserts results to the SP IndustryDictionary list (new terms inserted; existing ones updated
    /// with fresh AppliesTo + occurrence counts; Definition/Examples are never overwritten).
    /// </summary>
    public async Task<object> BuildDictionaryAsync(CancellationToken ct)
    {
        var apiKey = _config["Anthropic:ApiKey"];
        if (string.IsNullOrEmpty(apiKey))
            return new { error = "Anthropic:ApiKey not configured" };

        var model = _config["Claude:AnalysisModel"] ?? "claude-haiku-4-5-20251001";

        if (!File.Exists(CatalogTokensPath) || !File.Exists(SliTokensPath))
            return new { error = "Token files missing — run catalog/tokenize and sli/tokenize first" };

        var catTokens = JsonSerializer.Deserialize<List<ProductTokens>>(
            File.ReadAllText(CatalogTokensPath, Encoding.UTF8), Json) ?? [];
        var sliTokens = JsonSerializer.Deserialize<List<ProductTokens>>(
            File.ReadAllText(SliTokensPath, Encoding.UTF8), Json) ?? [];

        // term -> entry being built
        var entries = new Dictionary<string, IndustryDictionaryEntry>(StringComparer.OrdinalIgnoreCase);
        // term -> shapes seen (for AppliesTo)
        var shapeMap = new Dictionary<string, Dictionary<string, int>>(StringComparer.OrdinalIgnoreCase);

        IndustryDictionaryEntry GetEntry(string term, string termType, string mapsToToken)
        {
            if (!entries.TryGetValue(term, out var e))
            {
                e = new IndustryDictionaryEntry
                {
                    Term        = term,
                    TermType    = termType,
                    MapsToToken = mapsToToken,
                };
                entries[term]   = e;
                shapeMap[term]  = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);
            }
            return e;
        }

        void Track(string term, ProductTokens t, bool isCatalog)
        {
            var e = entries[term];
            if (isCatalog) e.CatalogCount++; else e.SliCount++;
            if (t.TkShape is not null)
            {
                shapeMap[term].TryGetValue(t.TkShape, out var c);
                shapeMap[term][t.TkShape] = c + 1;
            }
        }

        static string ClassifyAlloy(string alloy) =>
            System.Text.RegularExpressions.Regex.IsMatch(alloy, @"^[a-zA-Z]\d")
                ? "standard"       // ASTM grade: a36, a500, a513, a135, etc.
                : "alloy_series";  // numeric or alphanumeric series: 6061, 304, 1018, 360

        static string DisplayAlloy(string alloy) =>
            System.Text.RegularExpressions.Regex.IsMatch(alloy, @"^[a-zA-Z]\d")
                ? alloy.ToUpperInvariant()   // A36, A500, etc.
                : alloy.ToUpperInvariant();  // 6061, 304, etc.

        static string ClassifyCondition(string cond) =>
            cond.StartsWith("sch", StringComparison.OrdinalIgnoreCase) ? "standard" : "condition";

        void Mine(IEnumerable<ProductTokens> tokens, bool isCatalog)
        {
            foreach (var t in tokens.Where(t => !t.TokenizationFailed))
            {
                if (t.TkAlloy is { Length: > 0 } alloy)
                {
                    var display = DisplayAlloy(alloy);
                    GetEntry(display, ClassifyAlloy(alloy), alloy);
                    Track(display, t, isCatalog);
                }
                if (t.TkTemper is { Length: > 0 } temper)
                {
                    var display = temper.ToUpperInvariant();
                    GetEntry(display, "temper", temper);
                    Track(display, t, isCatalog);
                }
                foreach (var cond in t.TkConditions)
                {
                    GetEntry(cond, ClassifyCondition(cond), cond);
                    Track(cond, t, isCatalog);
                }
            }
        }

        Mine(catTokens, isCatalog: true);
        Mine(sliTokens, isCatalog: false);

        // Build AppliesTo: top shapes by occurrence count
        foreach (var (term, shapes) in shapeMap)
        {
            entries[term].AppliesTo = string.Join(", ",
                shapes.OrderByDescending(kv => kv.Value)
                      .Take(8)
                      .Select(kv => kv.Key));
        }

        _log.LogInformation("[Analysis] Dictionary: mined {Count} unique terms ({Cat} catalog, {Sli} SLI tokens)",
            entries.Count, catTokens.Count, sliTokens.Count);

        // AI-generate Definition + Examples for each term before writing to SP.
        // On re-runs the SP write skips Definition/Examples for existing rows, so
        // manual edits made in the SP list are never overwritten.
        var defs = await GenerateDefinitionsAsync(entries.Values.ToList(), apiKey, model, ct);
        foreach (var e in entries.Values)
        {
            if (defs.TryGetValue(e.Term, out var d))
            {
                e.Definition = d.Definition;
                e.Examples   = d.Examples;
            }
        }

        var (inserted, updated) = await _sp.WriteIndustryDictionaryEntriesAsync(
            entries.Values.ToList(), ct);

        return new
        {
            totalTerms = entries.Count,
            inserted,
            updated,
            byType = entries.Values
                .GroupBy(e => e.TermType)
                .OrderBy(g => g.Key)
                .ToDictionary(g => g.Key, g => g.OrderBy(e => e.Term).Select(e => new
                {
                    e.Term, e.MapsToToken, e.AppliesTo, e.CatalogCount, e.SliCount,
                    e.Definition, e.Examples
                })),
        };
    }

    public async Task<List<IndustryDictionaryEntry>> ReadDictionaryAsync(CancellationToken ct)
        => await _sp.ReadIndustryDictionaryAsync(ct);

    // ── Status ────────────────────────────────────────────────────────────────

    public object GetStatus()
    {
        static object FileInfo(string path)
        {
            if (!File.Exists(path)) return new { exists = false };
            var fi = new FileInfo(path);
            return new { exists = true, bytes = fi.Length, modified = fi.LastWriteTimeUtc };
        }
        var done  = _tokenizeDone;
        var total = _tokenizeTotal;
        var target = _tokenizeTarget;
        return new
        {
            catalogJson       = FileInfo(CatalogPath),
            sliJson           = FileInfo(SliPath),
            catalogTokensJson = FileInfo(CatalogTokensPath),
            sliTokensJson     = FileInfo(SliTokensPath),
            matchResultsJson  = FileInfo(MatchResultsPath),
            tokenize = total > 0 ? new
            {
                target,
                done,
                total,
                pct     = (int)(done * 100.0 / total),
                running = done < total && !string.IsNullOrEmpty(target),
            } : (object?)null,
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
        // Always force-refresh from SP so GT audit patches are immediately reflected
        await _sli.PopulateAsync(force: true, ct);

        var rows = _sli.TryGet();
        if (rows is null)
            return new { error = "SLI cache is empty -- proxy may still be starting up" };

        var sample = rows
            .Select(d => new SliEntry(
                ProductName:        GetStr(d, "ProductName") ?? "",
                ProductSearchKey:   GetStr(d, "CatalogProductSearchKey", "ProductSearchKey"),
                CatalogProductName: GetStr(d, "CatalogProductName"),
                SupplierName:       GetStr(d, "SupplierName"),
                RfqId:              GetStr(d, "JobReference"),
                SpItemId:           GetStr(d, "SpItemId")))
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

        _tokenizeTarget = label;
        Interlocked.Exchange(ref _tokenizeDone,  existing.Count);
        Interlocked.Exchange(ref _tokenizeTotal, source.Count);

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

            // For catalog target: also patch SP so tokens survive reinstall.
            if (target == "catalog")
            {
                var patches = existing.Values
                    .Where(t => !t.TokenizationFailed && t.SearchKey != null)
                    .Select(t => (SpItemId: _catalog.GetSpItemId(t.Name), Tokens: t))
                    .Where(p => p.SpItemId != null)
                    .Select(p => (p.SpItemId!, p.Tokens))
                    .ToList();
                if (patches.Count > 0)
                {
                    try { await _catalog.PatchCatalogTokensAsync(patches, ct); }
                    catch (Exception ex) { _log.LogWarning(ex, "[Analysis] SP token patch failed for batch"); }
                }
            }

            Interlocked.Exchange(ref _tokenizeDone, existing.Count);
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
                else
                    t.SearchKey = null; // not in current sli-sample (e.g. GT was cleared)
            }
        }

        // Only test entries that have a current GT assignment — entries absent from
        // sli-sample.json (cleared GT, new rows, or tokenised before this fetch) are excluded.
        var eligible = sliTokens.Where(t => t.SearchKey != null).ToList();

        var run    = new MatchTestRun { RunAt = DateTime.UtcNow };
        var sample = limit > 0 ? eligible.Take(limit).ToList() : eligible;

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

            // Attach supplier/catalog name + identity from pre-loaded lookup
            if (sliLookup.TryGetValue(supplier.Name, out var sliEntry))
            {
                mc.SupplierName        = sliEntry.SupplierName;
                mc.ExpectedCatalogName = sliEntry.CatalogProductName;
                mc.RfqId               = sliEntry.RfqId;
                mc.SpItemId            = sliEntry.SpItemId;
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
          "galvanized", "galvanneal", "bright_annealed",
          "ar_400", "ar_500", "ar_600", "weathering_steel", "painted",
          "brushed", "mirror_finish", "ornamental_180", "polished", "seamless",
          "round_corner", "tgp", "pno",
          "sch5", "sch10", "sch40", "sch80", "sch160" };
    // pno is exclusive (catalog PNO requires supplier to specify P&O) but NOT bidirectional
    // (supplier quoting P&O is still acceptable when catalog entry is plain — we'll take it).

    // Surface-treatment conditions are bidirectional: if the SUPPLIER description carries one
    // of these, the catalog entry must also carry it.  Prevents a galvanized supplier quote from
    // matching a plain HR catalog entry (or a #4-brushed stainless from matching plain 2B).
    private static readonly HashSet<string> BidirectionalConditions =
        new(StringComparer.OrdinalIgnoreCase)
        { "galvanized", "galvanneal", "bright_annealed", "painted",
          "brushed", "mirror_finish", "ornamental_180", "polished",
          "ar_400", "ar_500", "ar_600", "tgp" };

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
            // Shape/temper are normalized before comparison to handle equivalent labels
            // (rod == roundbar; t651/t6511/t652 all collapse to t6).
            if (!GatePasses(supplier.TkMetal,                 cat.TkMetal,                 out var r1)) { noMatchReason = r1; continue; }
            if (!GatePasses(NormalizeShape(supplier.TkShape),  NormalizeShape(cat.TkShape),  out var r2)) { noMatchReason = r2; continue; }
            if (!GatePasses(supplier.TkAlloy,                 cat.TkAlloy,                 out var r3)) { noMatchReason = r3; continue; }
            if (!GatePasses(NormalizeTemper(supplier.TkTemper),NormalizeTemper(cat.TkTemper),out var r4)) { noMatchReason = r4; continue; }

            // Exclusive-condition gate: catalog carries an exclusive tag the supplier
            // didn't mention -> not this product
            var catExclusive = cat.TkConditions
                .Where(c => ExclusiveConditions.Contains(c)).ToArray();
            if (catExclusive.Length > 0)
            {
                // ALL exclusive conditions must be present in supplier (not just any one).
                // E.g. catalog [galvanized, perforated] requires supplier to mention both.
                bool supplierMentionsAll = catExclusive.All(c =>
                    supplier.TkConditions.Contains(c, StringComparer.OrdinalIgnoreCase));
                if (!supplierMentionsAll)
                {
                    var missing = catExclusive.First(c =>
                        !supplier.TkConditions.Contains(c, StringComparer.OrdinalIgnoreCase));
                    noMatchReason = $"catalog requires condition '{missing}'";
                    continue;
                }
            }

            // Bidirectional surface-treatment gate: both directions must agree.
            // (1) Supplier has treatment → catalog must too (prevents galvanized quote → plain entry).
            // (2) Catalog has treatment → supplier must too (prevents plain pipe quote → GPI entry).
            var suppBidi = supplier.TkConditions
                .Where(c => BidirectionalConditions.Contains(c)).ToArray();
            if (suppBidi.Length > 0)
            {
                bool catalogHasIt = suppBidi.Any(c =>
                    cat.TkConditions.Contains(c, StringComparer.OrdinalIgnoreCase));
                if (!catalogHasIt)
                {
                    noMatchReason = $"supplier surface-treatment '{suppBidi[0]}' not in catalog";
                    continue;
                }
            }
            var catBidi = cat.TkConditions
                .Where(c => BidirectionalConditions.Contains(c)).ToArray();
            if (catBidi.Length > 0)
            {
                bool supplierHasIt = catBidi.Any(c =>
                    supplier.TkConditions.Contains(c, StringComparer.OrdinalIgnoreCase));
                if (!supplierHasIt)
                {
                    noMatchReason = $"catalog surface-treatment '{catBidi[0]}' not in supplier";
                    continue;
                }
            }

            // Scoring
            double score = 0;
            if (DimsMatch(supplier.TkDims, cat.TkDims))             score += 3;  // full match
            else if (DimsPartialMatch(supplier.TkDims, cat.TkDims)) score += 2;  // partial (fewer supplier dims)
            else if (DimsAnyDimMatch(supplier.TkDims, cat.TkDims))  score += 0.5; // at least one dim close
            if (ConditionsOverlap(supplier.TkConditions, cat.TkConditions))     score += 1;

            // Alloy specificity tie-break: when the supplier doesn't specify an alloy, prefer
            // generic catalog entries (alloy=null) over specific-alloy entries that happen to
            // have matching dims.  Avoids "HR Plate 3/8"" landing on "HR Plate A572 0.375"
            // when "HR Plate 0.375" (alloy=null) is also a candidate.
            if (supplier.TkAlloy == null && cat.TkAlloy != null) score -= 0.5;

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

            // Tiebreak: when scores are equal prefer fewer MSPC key segments.
            // Panel-specific catalog entries (SSH304/050/4/48120) have more segments than the
            // stock entry (SSH304/050/4), so stock entries win when no panel dims were supplied.
            var newWins = score > bestScore ||
                          (score == bestScore && best is not null &&
                           KeySegments(cat.SearchKey) < KeySegments(best.SearchKey));
            if (newWins)
            {
                bestScore    = score;
                best         = cat;
            }
        }

        // Only return a failReason when we have no match at all
        return (best, bestScore < 0 ? 0 : bestScore, best is null ? noMatchReason : null);
    }

    private static int KeySegments(string? key) => key is null ? 0 : key.Count(c => c == '/');

    private static string? NormalizeShape(string? s) => s?.ToLowerInvariant() switch
    {
        "strip" => "flatbar",   // HR/galv strip is the same product family as flatbar
        { } v   => v,
        null    => null
    };

    // T6511/T651/T652 are all T6-base tempers (stretch-relieved or stress-relieved variants).
    // Collapse them to "t6" so a catalog entry marked t6511 matches an SLI marked t6, and
    // vice versa. Other temper families (t4, t5, h-series, etc.) remain distinct.
    private static string? NormalizeTemper(string? t) =>
        t?.ToLowerInvariant() switch
        {
            "t651" or "t6511" or "t652" => "t6",
            { } s                        => s,
            null                         => null
        };

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

    /// <summary>
    /// Partial dims match: one side has fewer dimensions (e.g. supplier gives HxF but catalog has HxFxT).
    /// Checks that the shorter set of dims matches the leading (largest) dims of the longer set.
    /// Returns false when lengths are equal — use DimsMatch for that case.
    /// </summary>
    private static bool DimsPartialMatch(string? a, string? b, double tolerance = 0.05)
    {
        var da = ParseDims(a);
        var db = ParseDims(b);
        if (da is null || db is null || da.Length == db.Length) return false;
        double[] shorter = da.Length < db.Length ? da : db;
        double[] longer  = da.Length < db.Length ? db : da;
        // Sort both descending so largest dims align at index 0
        Array.Sort(shorter, (x, y) => y.CompareTo(x));
        Array.Sort(longer,  (x, y) => y.CompareTo(x));
        for (int i = 0; i < shorter.Length; i++)
        {
            double avg = (shorter[i] + longer[i]) / 2.0;
            if (avg < 0.001) continue;
            if (Math.Abs(shorter[i] - longer[i]) / avg > tolerance) return false;
        }
        return true;
    }

    /// <summary>
    /// Returns true when at least one dimension pair at the same ordinal position (after
    /// sorting both arrays ascending) is within tolerance.  Positional matching prevents
    /// cross-position false positives, e.g. flat bar 0.375×1.375 vs catalog 0.125×0.375
    /// where 0.375 appears at different positions (width vs thickness) and should NOT score.
    /// Legitimate case: 0.375×2.5 vs 0.375×4.0 — position-0 (thickness) matches in both.
    /// </summary>
    private static bool DimsAnyDimMatch(string? a, string? b, double tolerance = 0.05)
    {
        var da = ParseDims(a);
        var db = ParseDims(b);
        if (da is null || db is null) return false;
        var sa  = da.OrderBy(x => x).ToArray();
        var sb  = db.OrderBy(x => x).ToArray();
        int len = Math.Min(sa.Length, sb.Length);
        for (int i = 0; i < len; i++)
        {
            double avg = (sa[i] + sb[i]) / 2.0;
            if (avg >= 0.001 && Math.Abs(sa[i] - sb[i]) / avg <= tolerance) return true;
        }
        return false;
    }

    private static double[]? ParseDims(string? dims)
    {
        if (string.IsNullOrWhiteSpace(dims)) return null;
        // Strip literal "null" tokens that Claude may emit when a dim is absent.
        var parts = dims.Split('x', StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries)
            .Where(p => !string.Equals(p, "null", StringComparison.OrdinalIgnoreCase))
            .ToArray();
        if (parts.Length == 0) return null;
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

    // ── Live product matching (used during extraction pipeline) ───────────────

    /// <summary>
    /// Matches one supplier product name against the catalog using the token scorer.
    /// Priority: (1) user mapping cache, (2) token scorer with round-robin Claude Haiku / Gemini Flash.
    /// Returns null Source when no confident match is found (score below threshold).
    /// </summary>
    public async Task<TokenMatchResult> MatchProductAsync(
        string productName, string? supplierName = null, CancellationToken ct = default)
    {
        // 1. Check user-confirmed mappings first — highest priority.
        if (supplierName is not null)
        {
            var mapping = _mappingsCache.TryGetMapping(supplierName, productName);
            if (mapping?.ProductSearchKey is not null)
            {
                _log.LogInformation(
                    "[Match] User mapping: '{Name}' -> {Key}", productName, mapping.ProductSearchKey);
                return new TokenMatchResult
                {
                    SearchKey   = mapping.ProductSearchKey,
                    CatalogName = mapping.CatalogProductName,
                    Score       = 1.0,
                    Source      = "user_mapping",
                };
            }
        }

        // 2. Tokenise the supplier product name via round-robin AI.
        var tokens = await TokeniseSingleLiveAsync(productName, ct);
        if (tokens is null || tokens.TokenizationFailed)
        {
            _log.LogWarning("[Match] Tokenisation failed for '{Name}'", productName);
            return new TokenMatchResult { Source = null };
        }

        // 3. Score against catalog tokens.
        var catalog = await GetCatalogTokensAsync();
        if (catalog.Count == 0)
        {
            _log.LogWarning("[Match] Catalog token cache is empty — run /api/analysis/catalog/tokenize first");
            return new TokenMatchResult { Source = null };
        }

        // Collect all scored candidates above the metal/shape gate, sorted by score desc.
        var scored = new List<(ProductTokens Cat, double Score)>();
        foreach (var cat in catalog)
        {
            var (match, score, _) = FindBestMatchSingle(tokens, cat);
            if (match is not null) scored.Add((cat, score));
        }
        scored.Sort((a, b) => b.Score.CompareTo(a.Score));

        const double MinScore = 1.5; // require at least metal + shape match
        var top = scored.Take(3).Select(s => new TokenMatchCandidate
        {
            SearchKey   = s.Cat.SearchKey,
            CatalogName = s.Cat.Name,
            Score       = Math.Round(s.Score, 2),
        }).ToList();

        if (scored.Count == 0 || scored[0].Score < MinScore)
        {
            _log.LogInformation(
                "[Match] No confident match for '{Name}' (best={Best:F2})",
                productName, scored.FirstOrDefault().Score);
            return new TokenMatchResult { TopCandidates = top, Source = null };
        }

        var best = scored[0];
        _log.LogInformation(
            "[Match] Token match: '{Name}' -> {Key} ({Score:F2})",
            productName, best.Cat.SearchKey, best.Score);

        return new TokenMatchResult
        {
            SearchKey      = best.Cat.SearchKey,
            CatalogName    = best.Cat.Name,
            Score          = Math.Round(best.Score, 2),
            Source         = "token_scorer",
            TopCandidates  = top,
        };
    }

    // ── Custom ID creation ────────────────────────────────────────────────────

    /// <summary>
    /// Derives a deterministic CUSTOM_ID from a product's token bag.
    /// Same tokens on any proxy instance produce the same ID, preventing collisions
    /// without any distributed counter or coordination.
    /// Format: "CUST_" + first 8 hex chars of SHA-256(canonical token string).
    /// </summary>
    public static string ComputeCustomId(ProductTokens tokens)
    {
        var conds    = string.Join(",",
            (tokens.TkConditions ?? []).OrderBy(c => c, StringComparer.Ordinal));
        var canonical = $"{tokens.TkMetal ?? ""}|{tokens.TkShape ?? ""}|{tokens.TkAlloy ?? ""}|{tokens.TkTemper ?? ""}|{conds}|{tokens.TkDims ?? ""}";
        var hash     = SHA256.HashData(Encoding.UTF8.GetBytes(canonical));
        return "CUST_" + Convert.ToHexString(hash)[..8];
    }

    /// <summary>
    /// Tokenises <paramref name="term"/>, computes a deterministic CUSTOM_ID, and
    /// creates (or returns an existing) entry in the Product Catalog SP list.
    /// Called for SLI rows with no catalog match and for RLI rows with no MSPC.
    /// </summary>
    public async Task<(string CustomId, string DisplayName)> GetOrCreateCustomIdAsync(
        string term, string source, CancellationToken ct = default)
    {
        var tokens = await TokeniseSingleLiveAsync(term, ct);

        string customId;
        if (tokens is null || tokens.TokenizationFailed)
        {
            // Fallback: hash the raw term so we still get a stable, unique ID.
            var rawHash = SHA256.HashData(Encoding.UTF8.GetBytes(term.Trim().ToLowerInvariant()));
            customId    = "CUST_" + Convert.ToHexString(rawHash)[..8];
            _log.LogWarning("[Match] Tokenisation failed for '{Term}' — custom ID via raw hash: {Id}", term, customId);
            tokens      = new ProductTokens { Name = term };
        }
        else
        {
            tokens.Name = term;
            customId    = ComputeCustomId(tokens);
            _log.LogInformation("[Match] Custom ID for '{Term}': {Id} (metal={M} shape={S} alloy={A})",
                term, customId, tokens.TkMetal, tokens.TkShape, tokens.TkAlloy);
        }

        var (resultId, resultName) = await _catalog.GetOrCreateCustomEntryAsync(customId, term, tokens, source);
        return (resultId, resultName);
    }

    private static (ProductTokens? match, double score, string? failReason)
        FindBestMatchSingle(ProductTokens supplier, ProductTokens cat)
    {
        if (supplier.TkMetal != null && cat.TkMetal == null)
            return (null, -1, null);
        if (!GatePasses(supplier.TkMetal,                  cat.TkMetal,                  out _)) return (null, -1, null);
        if (!GatePasses(NormalizeShape(supplier.TkShape),   NormalizeShape(cat.TkShape),   out _)) return (null, -1, null);
        if (!GatePasses(supplier.TkAlloy,                  cat.TkAlloy,                  out _)) return (null, -1, null);
        if (!GatePasses(NormalizeTemper(supplier.TkTemper), NormalizeTemper(cat.TkTemper), out _)) return (null, -1, null);
        if (cat.TkConditions.Length > 0 && cat.TkConditions.Any(c => ExclusiveConditions.Contains(c))
            && !supplier.TkConditions.Any(c => cat.TkConditions.Contains(c, StringComparer.OrdinalIgnoreCase)))
            return (null, -1, null);

        double score = 0;
        if (supplier.TkMetal  is not null && supplier.TkMetal  == cat.TkMetal)  score += 1.0;
        if (supplier.TkShape  is not null && NormalizeShape(supplier.TkShape) == NormalizeShape(cat.TkShape))  score += 1.0;
        if (supplier.TkAlloy  is not null && supplier.TkAlloy  == cat.TkAlloy)  score += 1.0;
        if (supplier.TkTemper is not null && NormalizeTemper(supplier.TkTemper) == NormalizeTemper(cat.TkTemper)) score += 0.5;
        if (supplier.TkDims   is not null && cat.TkDims is not null)
        {
            if (DimsMatch(supplier.TkDims, cat.TkDims))         score += 1.0;
            else if (DimsPartialMatch(supplier.TkDims, cat.TkDims)) score += 0.5;
            else if (DimsAnyDimMatch(supplier.TkDims, cat.TkDims))  score += 0.5;
        }
        if (supplier.TkConditions.Length > 0 && ConditionsOverlap(supplier.TkConditions, cat.TkConditions)) score += 0.5;

        return (cat, score, null);
    }

    private async Task<List<ProductTokens>> GetCatalogTokensAsync()
    {
        if (!File.Exists(CatalogTokensPath)) return [];
        var fileTime = File.GetLastWriteTimeUtc(CatalogTokensPath);
        if (_catalogTokenCache is not null && fileTime <= _catalogTokenCacheAt)
            return _catalogTokenCache;

        try
        {
            var json = await File.ReadAllTextAsync(CatalogTokensPath);
            _catalogTokenCache  = JsonSerializer.Deserialize<List<ProductTokens>>(json, Json) ?? [];
            _catalogTokenCacheAt = fileTime;
            return _catalogTokenCache;
        }
        catch (Exception ex)
        {
            _log.LogWarning(ex, "[Match] Failed to load catalog token cache");
            return [];
        }
    }

    public async Task<Dictionary<string, ProductTokens>> GetCatalogTokensByKeyAsync(
        CancellationToken ct = default)
    {
        // Prefer SP-backed tokens (populated by TokenizeAsync / PatchCatalogTokensAsync).
        // Falls back to the file cache for migration path and fresh installs before first tokenize.
        var fromSp = _catalog.GetTokensByKey();
        if (fromSp.Count > 0) return fromSp;

        var tokens = await GetCatalogTokensAsync();
        return tokens
            .Where(t => !string.IsNullOrWhiteSpace(t.SearchKey) && !t.TokenizationFailed)
            .ToDictionary(t => t.SearchKey!, StringComparer.OrdinalIgnoreCase);
    }

    /// <summary>
    /// For each CUST_* entry, finds the best-matching real (non-CUST_) catalog MSPC using two methods:
    ///   1. Prefix match — catalog product name starts with (or is a prefix of) the CUST_ OriginalTerm.
    ///   2. Token score >= 3.5 via <see cref="FindBestMatchSingle"/> (when tokens are available).
    /// Returns one candidate per CUST_ key (highest score wins). The caller shows results to the
    /// user before committing via <c>POST /api/catalog/commit-promotions</c>.
    /// </summary>
    public Task<List<PromotionCandidate>> FindPromotionCandidatesAsync(
        IReadOnlyList<CustomCatalogEntry> custEntries, CancellationToken ct = default)
    {
        // Non-CUST_ catalog entries for name-based matching (includes newly imported rows)
        var allEntries = _catalog.CachedEntries
            .Where(e => e.SearchKey != null &&
                        !e.SearchKey.StartsWith("CUST_", StringComparison.OrdinalIgnoreCase))
            .ToList();

        // Non-CUST_ catalog tokens for token-based matching (may be empty for freshly imported rows)
        var allTokens = _catalog.GetTokensByKey()
            .Where(kv => !kv.Key.StartsWith("CUST_", StringComparison.OrdinalIgnoreCase))
            .ToDictionary(kv => kv.Key, kv => kv.Value, StringComparer.OrdinalIgnoreCase);

        var candidates = new List<PromotionCandidate>();

        foreach (var cust in custEntries)
        {
            string? bestMspc      = null;
            string? bestName      = null;
            double  bestScore     = -1;
            string  bestMatchType = "";

            // Method 1 — prefix match on OriginalTerm vs catalog names (no tokens needed)
            if (cust.OriginalTerm is { Length: > 0 })
            {
                foreach (var (catName, catKey) in allEntries)
                {
                    if (catKey is null) continue;
                    bool fwd = catName.StartsWith(cust.OriginalTerm, StringComparison.OrdinalIgnoreCase);
                    bool rev = cust.OriginalTerm.StartsWith(catName, StringComparison.OrdinalIgnoreCase);
                    if (!fwd && !rev) continue;

                    // Score = shorter/longer ratio; 1.0 when lengths are equal
                    double ratio = (double)Math.Min(cust.OriginalTerm.Length, catName.Length)
                                         / Math.Max(cust.OriginalTerm.Length, catName.Length);
                    if (ratio > bestScore)
                    {
                        bestScore     = ratio;
                        bestMspc      = catKey;
                        bestName      = catName;
                        bestMatchType = "prefix";
                    }
                }
            }

            // Method 2 — token score >= 3.5 (requires AI tokens on both sides)
            if (cust.TkMetal is not null || cust.TkShape is not null)
            {
                var custTokens = new OutlookShredder.Proxy.Models.ProductTokens
                {
                    Name         = cust.OriginalTerm ?? cust.ProductName ?? cust.SearchKey,
                    SearchKey    = cust.SearchKey,
                    TkMetal      = cust.TkMetal,
                    TkShape      = cust.TkShape,
                    TkAlloy      = cust.TkAlloy,
                    TkTemper     = cust.TkTemper,
                    TkConditions = cust.TkConditions ?? [],
                };

                foreach (var (mspc, catTokens) in allTokens)
                {
                    var (match, score, _) = FindBestMatchSingle(custTokens, catTokens);
                    if (match is not null && score >= 3.5 && score > bestScore)
                    {
                        bestScore     = score;
                        bestMspc      = mspc;
                        bestName      = catTokens.Name;
                        bestMatchType = "token";
                    }
                }
            }

            if (bestMspc is not null)
            {
                candidates.Add(new PromotionCandidate(
                    CustSearchKey:     cust.SearchKey,
                    CustSpItemId:      cust.SpItemId,
                    OriginalTerm:      cust.OriginalTerm,
                    ProductName:       cust.ProductName,
                    TargetMspc:        bestMspc,
                    TargetProductName: bestName ?? "",
                    Score:             bestScore,
                    MatchType:         bestMatchType));
            }
        }

        // One candidate per CUST_ key — take highest score
        var result = candidates
            .GroupBy(c => c.CustSearchKey, StringComparer.OrdinalIgnoreCase)
            .Select(g => g.OrderByDescending(c => c.Score).First())
            .ToList();

        _log.LogInformation("[Catalog] Promotion scan: {Cust} CUST_ entries, {Candidates} candidates",
            custEntries.Count, result.Count);

        return Task.FromResult(result);
    }

    private async Task<ProductTokens?> TokeniseSingleLiveAsync(string productName, CancellationToken ct)
    {
        await _liveSemaphore.WaitAsync(ct);
        try
        {
            var claudeKey  = _config["Anthropic:ApiKey"]    ?? "";
            var geminiKey  = _config["Google:ApiKey"]       ?? "";
            var claudeModel = _config["Claude:AnalysisModel"]  ?? "claude-haiku-4-5-20251001";
            var geminiModel = _config["Gemini:AnalysisModel"]  ?? "gemini-2.5-flash";

            // Alternate between Claude and Gemini per call; fall back immediately on failure.
            var counter = Interlocked.Increment(ref _rrCounter);
            bool claudeFirst = (counter % 2 == 1);

            List<ProductTokens?> result;
            if (claudeFirst && !string.IsNullOrEmpty(claudeKey))
            {
                result = await TokeniseBatchAsync([productName], claudeKey, claudeModel, ct);
                if (result[0] is null && !string.IsNullOrEmpty(geminiKey))
                    result = await TokeniseBatchGeminiAsync([productName], geminiKey, geminiModel, ct);
            }
            else if (!string.IsNullOrEmpty(geminiKey))
            {
                result = await TokeniseBatchGeminiAsync([productName], geminiKey, geminiModel, ct);
                if (result[0] is null && !string.IsNullOrEmpty(claudeKey))
                    result = await TokeniseBatchAsync([productName], claudeKey, claudeModel, ct);
            }
            else if (!string.IsNullOrEmpty(claudeKey))
            {
                result = await TokeniseBatchAsync([productName], claudeKey, claudeModel, ct);
            }
            else
            {
                _log.LogWarning("[Match] No AI key configured for live tokenisation");
                return null;
            }

            var tok = result[0];
            if (tok is not null) tok.Name = productName;
            return tok;
        }
        finally
        {
            _liveSemaphore.Release();
        }
    }

    private async Task<List<ProductTokens?>> TokeniseBatchGeminiAsync(
        List<string> names, string apiKey, string model, CancellationToken ct)
    {
        var numbered = string.Join("\n", names.Select((n, i) => $"{i + 1}. {n}"));
        var url = $"https://generativelanguage.googleapis.com/v1beta/models/{model}:generateContent?key={apiKey}";

        var body = JsonSerializer.Serialize(new
        {
            system_instruction = new { parts = new[] { new { text = TokeniseSystemPrompt } } },
            contents = new[] { new { role = "user", parts = new[] { new { text = $"Products:\n{numbered}" } } } },
            generationConfig = new { responseMimeType = "application/json" },
        }, Json);

        using var req = new HttpRequestMessage(HttpMethod.Post, url);
        req.Content = new StringContent(body, Encoding.UTF8, "application/json");

        HttpResponseMessage resp;
        try { resp = await _http.SendAsync(req, ct); }
        catch (Exception ex)
        {
            _log.LogWarning(ex, "[Analysis] Gemini call failed for batch");
            return Enumerable.Repeat<ProductTokens?>(null, names.Count).ToList();
        }

        var raw = await resp.Content.ReadAsStringAsync(ct);
        if (!resp.IsSuccessStatusCode)
        {
            _log.LogWarning("[Analysis] Gemini returned {Status}: {Body}",
                resp.StatusCode, raw[..Math.Min(200, raw.Length)]);
            return Enumerable.Repeat<ProductTokens?>(null, names.Count).ToList();
        }

        try
        {
            using var doc  = JsonDocument.Parse(raw);
            var text = doc.RootElement
                .GetProperty("candidates")[0]
                .GetProperty("content")
                .GetProperty("parts")[0]
                .GetProperty("text")
                .GetString() ?? "[]";

            text = text.Trim();
            if (text.StartsWith("```")) text = text[(text.IndexOf('\n') + 1)..];
            if (text.EndsWith("```"))   text = text[..text.LastIndexOf("```")].TrimEnd();

            var tokenDtos = JsonSerializer.Deserialize<List<ProductTokens>>(text, Json) ?? [];
            while (tokenDtos.Count < names.Count) tokenDtos.Add(null!);
            return tokenDtos.Take(names.Count).Select(t => (ProductTokens?)t).ToList();
        }
        catch (Exception ex)
        {
            _log.LogWarning(ex, "[Analysis] Failed to parse Gemini batch response");
            return Enumerable.Repeat<ProductTokens?>(null, names.Count).ToList();
        }
    }

    // ── GT Audit ─────────────────────────────────────────────────────────────

    // Prefixes whose GT assignments are confirmed substitutions or wrong product type.
    // All misses whose expectedSearchKey prefix is in this set will be CLEARED (null out PSK).
    private static readonly HashSet<string> GtClearPrefixes = new(StringComparer.OrdinalIgnoreCase)
        { "HCSC", "HBSB", "AA6063" }; // HA moved to GtUpdatePrefixes (remaining HA misses are dim-only errors)

    // Prefixes where the algorithm's match is confirmed correct and GT is wrong.
    // Misses in this set will be UPDATED to the algorithm's actualSearchKey.
    private static readonly HashSet<string> GtUpdatePrefixes = new(StringComparer.OrdinalIgnoreCase)
    {
        // Pass 1 (applied 2026-05-14)
        "GPI", "HPI", "HSH", "CSHCQ", "CTR", "HF", "SSH304", "CTSQ", "AR6061T6", "GSHG90",
        // Pass 2 — residual GT errors identified after pass 1
        "SPI304S",   // carbon A106 pipe labeled as stainless 304S pipe
        "HPA588",    // plain HR plate labeled as A588 weathering steel
        "HTSQ",      // hot square tube with wrong wall/dims in GT
        "CTRT",      // cold rect tube with wrong gauge (11ga recorded as 14ga)
        "AC6063SHC", // aluminum channel with wrong dimensions in GT
        "STSQ304MF", // stainless square tube with wrong wall (0.250 recorded for 0.125 product)
        "AF6061",    // aluminum flat bar GT for products that are actually tubes/bars
        "GTSQCQ",    // galvanized square tube — GT wrong (plain tubes or wrong dims)
        // Pass 3 — identified after pass 2
        "AA6061",    // flat bars mislabeled with angle prefix; angles with wrong dims
        "AH6061",    // "AH" (hex bar?) prefix for products that are flat bars
        "HPA572G50", // A572 GR50 plate with wrong thickness in GT (0.375 for 1.5" products)
        "SA304",     // brushed stainless angle labeled as non-brushed; rect tube labeled as angle
        "AP6061T0",  // generic aluminum plates bulk-assigned to nonsensical AP6061T0/1750 key
        "SF303TB",   // 304/316 flat bar products labeled as 303 True Bar in GT
        "STRT304MF", // ornamental rect tube labeled as Mill Finish (two ornamental rows + one dim-swap)
        "STSQ304",   // stainless sq tube label on round tube; steel tube labeled as stainless
        "HP",        // "Galv. Plate" products labeled as hot rolled plate HP (should be GP)
        // Pass 4 — identified after pass 3
        "HA",        // hot rolled angle with wrong dims in GT; moved from ClearPrefixes (remaining misses are dim-only)
        "GACQ",      // galvanized angle with wrong dims in GT
        "SR304",     // stainless round bar code used for flat bar product with 2 dimensions
        "AP3003TP",  // tread plate with wrong thickness in GT (0.375 for 0.190 product)
        "SF304SE",   // 304 alloy code on 316 alloy product
        "HR",        // rebar labeled as round bar; galvanized round bar with wrong GT prefix
        "GP",        // galvanized sheet/grating products labeled as galvanized plate GP/250
        "ATSQ6063",  // aluminum square tube with wrong wall thickness in GT (.065 for explicit .125-wall product)
        // Pass 5 — identified after pass 4
        "ASH5052",   // 5052 aluminum sheet with wrong thickness in GT (0.032 for .090/.125 products)
        "AP3003",    // 3003 alloy plate code used for explicit 5052 products
        "GTRTCQ",    // galvanized rect tube with wrong dims in GT
        "SP304",     // stainless plate 304 plain code for #4 polished product
        // Pass 6 — identified after pass 5
        "GRCQ",      // galvanized flat bar mislabeled with round bar (GRCQ) prefix
        "HTR",       // steel round tube with wrong OD in GT; or square tube labeled as round tube
        "HBWB",      // wide flange beams with wrong weight/section code in GT
        "HBHP572G50",// HP beam prefix used for standard WF beam product
    };

    /// <summary>
    /// Preview or apply GT audit: reads match-test-results.json, classifies each miss,
    /// and (when dryRun=false) patches SP CatalogProductSearchKey via PatchSliProductKeyAsync.
    /// SpItemId is resolved from the sli-sample.json SpItemId field when available, then falls
    /// back to a live SLI cache lookup keyed by (productName, rfqId, supplierName).
    /// </summary>
    public async Task<object> ApplyGtAuditAsync(bool dryRun, string actionFilter, CancellationToken ct)
    {
        if (!File.Exists(MatchResultsPath))
            return new { error = "match-test-results.json not found -- run match-test first" };

        var matchJson = await File.ReadAllTextAsync(MatchResultsPath, ct);
        var run = JsonSerializer.Deserialize<MatchTestRun>(matchJson, Json)!;
        var misses = run.Cases.Where(c => !c.IsHit && !c.IsNoMatch).ToList();

        // Build SpItemId lookup from sli-sample.json (productName -> spItemId)
        // Falls back to live SLI cache when sli-sample.json has no SpItemId.
        var sliSampleById = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase); // productName -> spItemId
        var sliSampleByKey = new Dictionary<string, SliEntry>(StringComparer.OrdinalIgnoreCase); // productName -> SliEntry
        if (File.Exists(SliPath))
        {
            var rawSli = await File.ReadAllTextAsync(SliPath, ct);
            var sliEntries = JsonSerializer.Deserialize<List<SliEntry>>(rawSli, Json)!;
            foreach (var e in sliEntries)
            {
                sliSampleByKey.TryAdd(e.ProductName, e);
                if (!string.IsNullOrWhiteSpace(e.SpItemId))
                    sliSampleById.TryAdd(e.ProductName, e.SpItemId);
            }
        }

        // Build live cache lookup as fallback: (productName|rfqId|supplierName) -> List<spItemId>
        var liveLookup = new Dictionary<string, List<string>>(StringComparer.OrdinalIgnoreCase);
        var cacheRows = _sli.TryGet();
        if (cacheRows is not null)
        {
            foreach (var row in cacheRows)
            {
                var pName = GetStr(row, "ProductName");
                var rfqId = GetStr(row, "JobReference");
                var sup   = GetStr(row, "SupplierName");
                var spId  = GetStr(row, "SpItemId");
                if (string.IsNullOrWhiteSpace(pName) || string.IsNullOrWhiteSpace(spId)) continue;
                var k = $"{pName.ToLowerInvariant()}|{(rfqId ?? "").ToLowerInvariant()}|{(sup ?? "").ToLowerInvariant()}";
                if (!liveLookup.TryGetValue(k, out var ids)) liveLookup[k] = ids = [];
                ids.Add(spId);
            }
        }

        var actions = new List<GtAuditAction>();

        foreach (var mc in misses)
        {
            var prefix = mc.ExpectedSearchKey?.Split('/')[0] ?? "";
            string action;
            if (GtClearPrefixes.Contains(prefix))
                action = "clear";
            else if (GtUpdatePrefixes.Contains(prefix) && !string.IsNullOrWhiteSpace(mc.ActualSearchKey))
                action = "update";
            else
                action = "review";

            bool include = actionFilter switch
            {
                "both"   => action is "clear" or "update",
                "clear"  => action == "clear",
                "update" => action == "update",
                "review" => action == "review",
                _        => true // "all"
            };
            if (!include) continue;

            // Resolve SpItemId(s) — sli-sample first, then live cache
            var spIds = new List<string>();
            if (!string.IsNullOrWhiteSpace(mc.SpItemId))
            {
                spIds.Add(mc.SpItemId);
            }
            else if (sliSampleById.TryGetValue(mc.ProductName, out var sampleId))
            {
                spIds.Add(sampleId);
            }
            else
            {
                sliSampleByKey.TryGetValue(mc.ProductName, out var sliEntry);
                var rfq = mc.RfqId ?? sliEntry?.RfqId ?? "";
                var sup = mc.SupplierName ?? sliEntry?.SupplierName ?? "";
                var k = $"{mc.ProductName.ToLowerInvariant()}|{rfq.ToLowerInvariant()}|{sup.ToLowerInvariant()}";
                if (liveLookup.TryGetValue(k, out var ids)) spIds.AddRange(ids);
            }

            string? newKey = action == "update" ? mc.ActualSearchKey : null;
            string? newCatName = null;
            if (action == "update" && newKey != null)
                newCatName = _catalog.FindBySearchKey(newKey)?.Name;

            // Emit one action per resolved SpItemId (could be multiple rows for same product)
            if (spIds.Count == 0)
            {
                actions.Add(new GtAuditAction
                {
                    ProductName    = mc.ProductName,
                    SupplierName   = mc.SupplierName,
                    RfqId          = mc.RfqId,
                    Action         = action,
                    OldSearchKey   = mc.ExpectedSearchKey,
                    NewSearchKey   = newKey,
                    NewCatalogName = newCatName,
                    Note           = "SpItemId not found in sli-sample.json or live cache"
                });
                continue;
            }

            foreach (var spId in spIds)
            {
                var a = new GtAuditAction
                {
                    ProductName    = mc.ProductName,
                    SupplierName   = mc.SupplierName,
                    RfqId          = mc.RfqId,
                    SpItemId       = spId,
                    Action         = action,
                    OldSearchKey   = mc.ExpectedSearchKey,
                    NewSearchKey   = newKey,
                    NewCatalogName = newCatName,
                };

                if (!dryRun && action != "review")
                {
                    try
                    {
                        await _sp.PatchSliProductKeyAsync(spId, newKey, newCatName);
                        a.Applied = true;
                    }
                    catch (Exception ex)
                    {
                        a.Error = ex.Message;
                        _log.LogWarning(ex, "[Analysis] GT audit patch failed for SLI {SpId}", spId);
                    }
                }

                actions.Add(a);
            }
        }

        var cleared  = actions.Count(a => a.Action == "clear");
        var updated  = actions.Count(a => a.Action == "update");
        var reviewed = actions.Count(a => a.Action == "review");
        var noId     = actions.Count(a => a.SpItemId == null && a.Action != "review");
        var applied  = actions.Count(a => a.Applied);
        var errored  = actions.Count(a => a.Error != null);

        _log.LogInformation("[Analysis] GT audit {Mode}: clear={C} update={U} review={R} noId={N} applied={A} errors={E}",
            dryRun ? "dry-run" : "APPLY", cleared, updated, reviewed, noId, applied, errored);

        return new
        {
            dryRun,
            summary = new { cleared, updated, reviewed, noId, applied, errored },
            actions
        };
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

/// <summary>A CUST_* → real MSPC promotion candidate returned by FindPromotionCandidatesAsync.</summary>
public record PromotionCandidate(
    string  CustSearchKey,
    string  CustSpItemId,
    string? OriginalTerm,
    string? ProductName,
    string  TargetMspc,
    string  TargetProductName,
    double  Score,
    string  MatchType);
