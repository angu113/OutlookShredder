using System.Text;
using System.Text.Json;
using OutlookShredder.Proxy.Models;

namespace OutlookShredder.Proxy.Services;

/// <summary>
/// Produces a per-day pricing analysis over supplier quote SLI rows.
/// Raw SLI data is cached to disk per date so repeat calls are fast and
/// data accumulates day-by-day without re-fetching SharePoint.
/// </summary>
public class PricingAnalysisService
{
    private readonly SliCacheService                    _sliCache;
    private readonly IConfiguration                    _config;
    private readonly IHttpClientFactory                _http;
    private readonly ILogger<PricingAnalysisService>   _log;
    private readonly string                            _cacheDir;

    // Checked in order; more-specific phrases first to avoid prefix collisions.
    private static readonly string[] KnownMetals =
    [
        "Stainless Steel", "Alloy Steel", "Aluminum", "Steel",
        "Copper", "Brass", "Bronze", "Titanium", "Nickel"
    ];

    // Checked in order; longer / more specific phrases first.
    private static readonly string[] KnownShapes =
    [
        "Tread Plate", "Structural Tube", "DOM Tube", "Round Tube", "Square Tube",
        "Rectangular Tube", "Flat Bar", "Round Bar", "Square Bar", "Hex Bar",
        "Wide Flange", "I-Beam", "I Beam", "H-Beam", "H Beam", "Angle", "Channel", "Beam",
        "Expanded Metal", "Perforated", "Grating",
        "Sheet", "Plate", "Coil", "Strip", "Rod", "Wire", "Pipe"
    ];

    // Secondary indicators: product-name tokens that imply a metal when KnownMetals doesn't match.
    // Each entry: (indicator, canonical metal). Checked against lower-cased name.
    private static readonly (string Indicator, string Metal)[] MetalIndicators =
    [
        // Stainless via alloy grade (before steel indicators to avoid "Steel" grabbing 304/316 rows)
        ("stainless",      "Stainless Steel"),
        ("304",            "Stainless Steel"),
        ("316",            "Stainless Steel"),
        ("310",            "Stainless Steel"),
        ("321",            "Stainless Steel"),
        // Aluminum via alloy series
        ("6061",           "Aluminum"),
        ("6063",           "Aluminum"),
        ("6082",           "Aluminum"),
        ("5052",           "Aluminum"),
        ("5083",           "Aluminum"),
        ("3003",           "Aluminum"),
        ("2024",           "Aluminum"),
        ("7075",           "Aluminum"),
        // Steel via processing / grade descriptors
        ("hot rolled",     "Steel"),
        ("cold rolled",    "Steel"),
        ("cold drawn",     "Steel"),
        ("galvanized",     "Steel"),
        ("hr ",            "Steel"),   // "HR Standard Beam", "HR Flat Bar" etc.
        (" hr",            "Steel"),   // trailing "... 20' HR"
        ("a36",            "Steel"),
        ("a572",           "Steel"),
        ("a513",           "Steel"),
        ("a500",           "Steel"),
        ("a106",           "Steel"),
        ("a53",            "Steel"),
        ("1018",           "Steel"),
        ("1020",           "Steel"),
        ("1045",           "Steel"),
        ("12l14",          "Steel"),
    ];

    private const string ClassifySystemPrompt = """
        You are classifying supplier metal product quote line items for a pricing analysis system.

        For each product, identify:
        1. "special_conditions": array of tags from this list only — use empty [] for plain material:
           polished, anodized, tread_plate, dom, painted, hot_rolled, cold_rolled, stress_proof,
           chrome_plated, galvanized, extruded, drawn, seamless, welded, hex, key_stock,
           perforated, expanded, grating, mirror_finish, brushed, pvc_coated, tig_welded

        2. "is_service": true ONLY when the line item is pure processing work (bending, flame cutting,
           saw cutting, drilling, threading, machining, tapping, rolling, punching, deburring,
           anodizing-as-a-service, painting-as-a-service) — NOT for coated/finished material sold by weight.

        3. "confidence_note": one short sentence if there is genuine ambiguity, else "".

        Respond ONLY via the classify_products tool.
        """;

    private const string ClassifyToolSchema = """
        {
          "name": "classify_products",
          "description": "Return classification for each metal product description",
          "input_schema": {
            "type": "object",
            "properties": {
              "results": {
                "type": "array",
                "items": {
                  "type": "object",
                  "properties": {
                    "id":                 { "type": "string" },
                    "special_conditions": { "type": "array", "items": { "type": "string" } },
                    "is_service":         { "type": "boolean" },
                    "confidence_note":    { "type": "string" }
                  },
                  "required": ["id", "special_conditions", "is_service", "confidence_note"]
                }
              }
            },
            "required": ["results"]
          }
        }
        """;

    public PricingAnalysisService(
        SliCacheService                  sliCache,
        IConfiguration                   config,
        IHttpClientFactory               http,
        ILogger<PricingAnalysisService>  log)
    {
        _sliCache = sliCache;
        _config   = config;
        _http     = http;
        _log      = log;
        _cacheDir = Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
            "Shredder", "Proxy", "cache", "pricing");
        Directory.CreateDirectory(_cacheDir);
    }

    // ── Public entry point ───────────────────────────────────────────────────

    public async Task<PricingReport> GetReportAsync(DateOnly date, CancellationToken ct)
    {
        var (rawItems, fromCache) = await LoadOrFetchRawAsync(date, ct);

        var totalRows = rawItems.Count;
        var regretExcluded = 0;
        var processable = new List<PricingReportItem>(rawItems.Count);

        foreach (var sli in rawItems)
        {
            if (GetBool(sli, "IsRegret") || GetBool(sli, "IsDeleted"))
            { regretExcluded++; continue; }

            var (price, source, confidence, confNote) = DerivePrice(sli);
            var catalogName = GetStr(sli, "CatalogProductName");
            var productName = GetStr(sli, "ProductName") ?? "";
            var (metal, shape) = ParseMetalShape(catalogName, productName);

            processable.Add(new PricingReportItem
            {
                SpItemId           = GetStr(sli, "SpItemId") ?? GetStr(sli, "id") ?? "",
                RfqId              = GetStr(sli, "JobReference") ?? GetStr(sli, "RFQ_ID") ?? "",
                SupplierName       = GetStr(sli, "SupplierName") ?? "",
                ProductName        = productName,
                CatalogProductName = catalogName,
                Metal              = metal,
                Shape              = shape,
                PricePerPound      = price,
                PriceSource        = source,
                Confidence         = confidence,
                ConfidenceNote     = confNote,
                ReceivedAt         = GetDate(sli, "ReceivedAt") ?? DateTime.UtcNow,
                SpecialConditions  = [],
                IsService          = false
            });
        }

        _log.LogInformation("[Pricing] {Date}: {Total} rows, {Regret} regrets → {Proc} to classify",
            date, totalRows, regretExcluded, processable.Count);

        await BatchClassifyAsync(processable, ct);

        var report = BuildReport(date, fromCache, totalRows, regretExcluded, processable);
        LogLowConfidence(date, report.LowMediumConfidenceItems);
        return report;
    }

    // ── Raw data fetch / cache ───────────────────────────────────────────────

    private async Task<(List<Dictionary<string, object?>> Items, bool FromCache)>
        LoadOrFetchRawAsync(DateOnly date, CancellationToken ct)
    {
        var rawPath = Path.Combine(_cacheDir, $"pricing-raw-{date:yyyy-MM-dd}.json");

        if (File.Exists(rawPath))
        {
            try
            {
                var json = await File.ReadAllTextAsync(rawPath, ct);
                var loaded = JsonSerializer.Deserialize<List<Dictionary<string, JsonElement>>>(json);
                if (loaded is not null)
                {
                    var items = loaded
                        .Select(d => d.ToDictionary(kv => kv.Key, kv => (object?)kv.Value))
                        .ToList();
                    _log.LogInformation("[Pricing] {Date}: loaded {Count} rows from disk cache", date, items.Count);
                    return (items, true);
                }
            }
            catch (Exception ex)
            {
                _log.LogWarning(ex, "[Pricing] {Date}: disk cache corrupt — re-fetching", date);
            }
        }

        var startUtc = date.ToDateTime(TimeOnly.MinValue, DateTimeKind.Utc);
        var endUtc   = startUtc.AddDays(1);

        var allSli = _sliCache.TryGet();
        if (allSli is null)
        {
            _log.LogInformation("[Pricing] {Date}: SLI cache stale — populating", date);
            await _sliCache.PopulateAsync(force: false, ct);
            allSli = _sliCache.TryGet() ?? [];
        }

        var raw = allSli
            .Where(item =>
            {
                var recv = GetDate(item, "ReceivedAt");
                return recv.HasValue && recv.Value >= startUtc && recv.Value < endUtc;
            })
            .ToList();

        _log.LogInformation("[Pricing] {Date}: filtered {Count} rows from SLI cache", date, raw.Count);

        await SaveRawCacheAsync(rawPath, raw, ct);
        return (raw, false);
    }

    private async Task SaveRawCacheAsync(
        string rawPath,
        List<Dictionary<string, object?>> items,
        CancellationToken ct)
    {
        try
        {
            var opts = new JsonSerializerOptions { WriteIndented = false };
            var json = JsonSerializer.Serialize(items, opts);
            await File.WriteAllTextAsync(rawPath, json, Encoding.UTF8, ct);
            _log.LogInformation("[Pricing] Saved {Count} rows to {Path}", items.Count, rawPath);
        }
        catch (Exception ex)
        {
            _log.LogWarning(ex, "[Pricing] Failed to save raw cache to {Path}", rawPath);
        }
    }

    // ── Price derivation ─────────────────────────────────────────────────────

    private static (double? Price, string Source, string Confidence, string? Note)
        DerivePrice(Dictionary<string, object?> row)
    {
        var ppp = GetNum(row, "PricePerPound");
        if (ppp is > 0)
            return (ppp, "direct", "high", null);

        var total  = GetNum(row, "TotalPrice");
        var qty    = GetNum(row, "UnitsQuoted") ?? GetNum(row, "UnitsRequested");
        var weight = GetNum(row, "WeightPerUnit");

        if (total is > 0 && qty is > 0 && weight is > 0)
        {
            var unit    = GetStr(row, "WeightUnit");
            var totalLb = qty.Value * ToPounds(weight.Value, unit);
            if (totalLb > 0)
                return (total.Value / totalLb, "total_weight", "high", null);
        }

        // Medium: $/ft + length + weight → $/lb (extra conversion step, density not assumed)
        var ppf    = GetNum(row, "PricePerFoot");
        var length = GetNum(row, "LengthPerUnit");
        if (ppf is > 0 && length is > 0 && weight is > 0)
        {
            var lengthFt = ToFeet(length.Value, GetStr(row, "LengthUnit"));
            var lbs      = ToPounds(weight.Value, GetStr(row, "WeightUnit"));
            if (lbs > 0)
                return (ppf.Value * lengthFt / lbs, "foot_weight", "medium",
                    "Derived from $/ft × length / weight");
        }

        if (total is > 0 && qty is > 0)
            return (null, "total_only", "low",
                "TotalPrice present but no weight — $/lb uncomputable");

        return (null, "uncomputable", "low", "No usable price fields");
    }

    // ── Metal / shape parsing ────────────────────────────────────────────────

    // Metal: try product name first (supplier's own words are most accurate), then catalog name.
    // This avoids wrong-catalog-match cases where e.g. "Copper Sheet" is matched to
    // "Hot Rolled Sheet" in the catalog, which would otherwise classify it as Steel.
    private static (string? Metal, string? Shape) ParseMetalShape(string? catalogName, string productName)
    {
        var metal = ExtractMetal(productName)
                 ?? (catalogName is not null ? ExtractMetal(catalogName) : null);

        // Shape: prefer catalog name (more structured/standardized), fall back to product name.
        var shape = (catalogName is not null ? ExtractShape(catalogName) : null)
                 ?? ExtractShape(productName);

        return (metal, shape);
    }

    private static string? ExtractMetal(string name)
    {
        if (string.IsNullOrWhiteSpace(name)) return null;
        var lower = name.ToLowerInvariant();
        foreach (var m in KnownMetals)
            if (lower.Contains(m.ToLowerInvariant())) return m;
        foreach (var (indicator, implied) in MetalIndicators)
            if (lower.Contains(indicator)) return implied;
        return null;
    }

    private static string? ExtractShape(string name)
    {
        if (string.IsNullOrWhiteSpace(name)) return null;
        var lower = name.ToLowerInvariant();
        foreach (var s in KnownShapes)
            if (lower.Contains(s.ToLowerInvariant())) return s;
        return null;
    }

    // ── AI batch classification ──────────────────────────────────────────────

    private async Task BatchClassifyAsync(List<PricingReportItem> items, CancellationToken ct)
    {
        var apiKey = _config["Anthropic:ApiKey"];
        if (string.IsNullOrEmpty(apiKey))
        {
            _log.LogWarning("[Pricing] Anthropic:ApiKey not configured — skipping AI classification");
            return;
        }

        var model     = _config["Claude:Model"] ?? "claude-sonnet-4-6";
        var toolEl    = JsonDocument.Parse(ClassifyToolSchema).RootElement.Clone();
        const int batchSize = 15;
        var totalBatches = (items.Count + batchSize - 1) / batchSize;

        for (int i = 0; i < items.Count && !ct.IsCancellationRequested; i += batchSize)
        {
            var batch     = items.GetRange(i, Math.Min(batchSize, items.Count - i));
            var batchNum  = i / batchSize + 1;
            _log.LogInformation("[Pricing] Classifying batch {Batch}/{Total} ({Count} items)",
                batchNum, totalBatches, batch.Count);
            await ClassifyBatchAsync(batch, model, apiKey, toolEl, ct);
        }
    }

    private async Task ClassifyBatchAsync(
        List<PricingReportItem> batch,
        string model,
        string apiKey,
        JsonElement toolEl,
        CancellationToken ct)
    {
        var payload = batch.Select((item, idx) => new
        {
            id           = idx.ToString(),
            product_name = item.CatalogProductName ?? item.ProductName
        });

        var userText = $"Classify these {batch.Count} metal product descriptions:\n"
                     + JsonSerializer.Serialize(payload);

        var body = JsonSerializer.Serialize(new
        {
            model,
            max_tokens  = 1024,
            system      = ClassifySystemPrompt,
            tools       = new[] { toolEl },
            tool_choice = new { type = "tool", name = "classify_products" },
            messages    = new[] { new { role = "user", content = userText } }
        });

        using var http = _http.CreateClient();
        http.DefaultRequestHeaders.Add("x-api-key", apiKey);
        http.DefaultRequestHeaders.Add("anthropic-version", "2023-06-01");

        HttpResponseMessage resp;
        try
        {
            resp = await http.PostAsync(
                "https://api.anthropic.com/v1/messages",
                new StringContent(body, Encoding.UTF8, "application/json"),
                ct);
        }
        catch (Exception ex)
        {
            _log.LogWarning(ex, "[Pricing] Claude classification HTTP error");
            return;
        }

        if (!resp.IsSuccessStatusCode)
        {
            _log.LogWarning("[Pricing] Claude classification returned {Status}", resp.StatusCode);
            return;
        }

        var respJson = await resp.Content.ReadAsStringAsync(ct);
        ApplyClassificationResults(batch, respJson);
    }

    private void ApplyClassificationResults(List<PricingReportItem> batch, string respJson)
    {
        try
        {
            using var doc = JsonDocument.Parse(respJson);
            var content   = doc.RootElement.GetProperty("content");

            foreach (var block in content.EnumerateArray())
            {
                if (!block.TryGetProperty("type", out var t) || t.GetString() != "tool_use") continue;
                if (!block.TryGetProperty("input", out var input))                           continue;
                if (!input.TryGetProperty("results", out var results))                       continue;

                foreach (var result in results.EnumerateArray())
                {
                    if (!result.TryGetProperty("id", out var idEl)) continue;
                    if (!int.TryParse(idEl.GetString(), out var idx) || idx >= batch.Count) continue;

                    var conditions = new List<string>();
                    if (result.TryGetProperty("special_conditions", out var condsEl))
                        foreach (var c in condsEl.EnumerateArray())
                            if (c.GetString() is { Length: > 0 } s)
                                conditions.Add(s);

                    var isService = result.TryGetProperty("is_service", out var svcEl)
                                 && svcEl.ValueKind == JsonValueKind.True;

                    var note = result.TryGetProperty("confidence_note", out var noteEl)
                               ? noteEl.GetString() : null;

                    batch[idx] = batch[idx] with
                    {
                        SpecialConditions = [.. conditions],
                        IsService         = isService,
                        AiNote            = string.IsNullOrWhiteSpace(note) ? null : note
                    };
                }
                break;
            }
        }
        catch (Exception ex)
        {
            _log.LogWarning(ex, "[Pricing] Failed to parse Claude classification response");
        }
    }

    // ── Report assembly ──────────────────────────────────────────────────────

    private static PricingReport BuildReport(
        DateOnly                 date,
        bool                     fromCache,
        int                      totalRows,
        int                      regretExcluded,
        List<PricingReportItem>  items)
    {
        var servicesExcluded = items.Count(i => i.IsService);
        var priced = items
            .Where(i => !i.IsService && i.PricePerPound is > 0)
            .ToList();
        var unpricedExcluded = items.Count(i => !i.IsService && i.PricePerPound is null or <= 0);

        var categories = priced
            .GroupBy(i =>
            {
                var condKey = i.SpecialConditions.Length > 0
                    ? string.Join("+", i.SpecialConditions.OrderBy(c => c))
                    : "";
                return $"{i.Metal ?? "Unknown"}|{i.Shape ?? "Unknown"}|{condKey}";
            })
            .Select(g =>
            {
                var high       = g.Where(i => i.Confidence == "high").ToList();
                var allPrices  = g.Select(i => i.PricePerPound!.Value).ToList();
                var highPrices = high.Select(i => i.PricePerPound!.Value).ToList();

                var parts    = g.Key.Split('|');
                var condPart = parts.Length > 2 && parts[2].Length > 0
                    ? parts[2].Split('+') : [];

                return new PricingCategory
                {
                    Metal                = parts[0],
                    Shape                = parts[1],
                    Conditions           = condPart,
                    CategoryKey          = g.Key,
                    TotalQuotes          = g.Count(),
                    HighConfidenceQuotes = high.Count,
                    AvgPricePerPound     = highPrices.Count > 0 ? highPrices.Average() : null,
                    MinPricePerPound     = allPrices.Count > 0 ? allPrices.Min() : null,
                    MaxPricePerPound     = allPrices.Count > 0 ? allPrices.Max() : null
                };
            })
            .OrderBy(c => c.Metal)
            .ThenBy(c => c.Shape)
            .ThenBy(c => c.CategoryKey)
            .ToList();

        var lowMed = items
            .Where(i => i.Confidence is "low" or "medium")
            .ToList();

        return new PricingReport
        {
            Date                  = date.ToString("yyyy-MM-dd"),
            FromCache             = fromCache,
            GeneratedAt           = DateTime.UtcNow,
            TotalSliRows          = totalRows,
            RegretExcluded        = regretExcluded,
            ServicesExcluded      = servicesExcluded,
            UnpricedExcluded      = unpricedExcluded,
            HighConfidenceCount   = priced.Count(i => i.Confidence == "high"),
            MediumConfidenceCount = priced.Count(i => i.Confidence == "medium"),
            LowConfidenceCount    = priced.Count(i => i.Confidence == "low"),
            Categories            = categories,
            AllItems              = items,
            LowMediumConfidenceItems = lowMed
        };
    }

    // ── Logging ──────────────────────────────────────────────────────────────

    private void LogLowConfidence(DateOnly date, List<PricingReportItem> items)
    {
        if (items.Count == 0) return;
        _log.LogInformation("[Pricing] {Date}: {Count} low/medium confidence items:", date, items.Count);
        foreach (var item in items)
        {
            _log.LogInformation(
                "[Pricing]  {Conf,6} | {Rfq} | {Supplier} | {Product} | src={Source} | {Note}",
                item.Confidence, item.RfqId, item.SupplierName,
                item.CatalogProductName ?? item.ProductName,
                item.PriceSource,
                item.ConfidenceNote ?? item.AiNote ?? "");
        }
    }

    // ── Helpers ──────────────────────────────────────────────────────────────

    private static double? GetNum(Dictionary<string, object?> row, string key)
    {
        if (!row.TryGetValue(key, out var v) || v is null) return null;
        return v switch
        {
            double d                                        => d,
            int    i                                        => (double)i,
            JsonElement je when je.ValueKind == JsonValueKind.Number => je.GetDouble(),
            _                                               => double.TryParse(v.ToString(), out var d) ? d : null
        };
    }

    private static string? GetStr(Dictionary<string, object?> row, string key)
    {
        if (!row.TryGetValue(key, out var v) || v is null) return null;
        return v is JsonElement je ? je.ToString() : v.ToString();
    }

    private static bool GetBool(Dictionary<string, object?> row, string key)
    {
        if (!row.TryGetValue(key, out var v) || v is null) return false;
        return v switch
        {
            bool b                                        => b,
            JsonElement je when je.ValueKind == JsonValueKind.True => true,
            _                                             => false
        };
    }

    private static DateTime? GetDate(Dictionary<string, object?> row, string key)
    {
        if (!row.TryGetValue(key, out var v) || v is null) return null;
        var s = v is JsonElement je ? je.ToString() : v.ToString();
        return DateTime.TryParse(s, null,
            System.Globalization.DateTimeStyles.RoundtripKind, out var dt)
            ? dt.ToUniversalTime() : null;
    }

    private static double ToPounds(double w, string? unit) => unit?.ToLowerInvariant() switch
    {
        "kg" => w * 2.20462,
        "oz" => w / 16.0,
        "g"  => w / 453.592,
        _    => w
    };

    private static double ToFeet(double l, string? unit) => unit?.ToLowerInvariant() switch
    {
        "in" => l / 12.0,
        "m"  => l * 3.28084,
        "mm" => l / 304.8,
        "cm" => l / 30.48,
        _    => l
    };
}
