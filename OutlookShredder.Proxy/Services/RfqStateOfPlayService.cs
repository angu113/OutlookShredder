using System.Security.Cryptography;
using System.Text;
using System.Text.Json;

namespace OutlookShredder.Proxy.Services;

/// <summary>
/// The richer "state of play" AI comparison for an RFQ's competing quotes — supersedes the ≤3-bullet
/// summarize call. Builds an input from the STRUCTURED supplier line items (the supplier's own product
/// names + specs + prices + lead time / certs / comments / regret flags) plus each supplier's email
/// body, and asks Claude for a tight narrative on who to buy from and why.
///
/// Deliberately code-free: MSPC / ProductSearchKey / CatalogProductName are NEVER sent — our MSPCs can
/// be custom-generated on close matches, so the same item can carry different codes across suppliers.
/// The model reconciles "this is the same item" from the supplier names + specs across documents.
/// </summary>
public class RfqStateOfPlayService
{
    private readonly IHttpClientFactory             _http;
    private readonly IConfiguration                 _config;
    private readonly ILogger<RfqStateOfPlayService> _log;

    public RfqStateOfPlayService(IHttpClientFactory http, IConfiguration config, ILogger<RfqStateOfPlayService> log)
    {
        _http = http; _config = config; _log = log;
    }

    private const string SystemPrompt =
        "You are a metal-supplier purchasing analyst. Given competing supplier quotes for one RFQ — the " +
        "structured line items and each supplier's own email — write a TIGHT 'state of play' for a busy " +
        "purchasing rep: what's the current best move and why.\n" +
        "MATCH PRODUCTS BY NAME + SPEC across suppliers (e.g. \"304 SS plate 1/4 x 48 x 96\"): the same " +
        "item may be worded differently by each supplier. There are no product codes — rely on the text.\n" +
        "ACCOUNT FOR COVERAGE: a lower TOTAL that skipped/regretted items is NOT cheaper — compare only " +
        "complete, apples-to-apples quotes; a supplier 'regret' (cannot supply an item) IS material.\n" +
        "WE SPLIT BY LINE: on a multi-line RFQ we pick and choose — buy each line from whoever is cheapest on " +
        "that line. Identify the WINNER per line and the best line-by-line split. Only discuss a supplier " +
        "that WINS at least one line — do NOT spend words on suppliers who win nothing.\n" +
        "NEGOTIATION: when a supplier wins some lines but loses others, flag the play to push their LOSING " +
        "lines down (using their winning lines as leverage) so awarding them the whole order beats the split " +
        "— name the lines and the dollar gap to close. Only mention leverage when it ACTUALLY exists (2+ " +
        "suppliers competing on a line); NEVER state that there is no leverage or that an item is single-sourced.\n" +
        "WRITE TIGHTLY: a 1-line recommendation (who, and why), then a handful of short lines on the price " +
        "leader, the genuine trade-offs, and gaps/risks. Mention a dimension (lead time, freight terms, " +
        "payment terms, certs, MOQ, surcharges) ONLY when it MEANINGFULLY differs between suppliers — skip " +
        "it otherwise. MTRs (material test reports) are ALWAYS required and supplied as standard here — NEVER " +
        "note a missing or unmentioned MTR as a gap or risk; only raise certs if a supplier explicitly cannot " +
        "provide standard MTRs or offers something beyond them. IGNORE quote validity / expiry dates and " +
        "times — never flag a short, burning, or expiring quote as a risk or an action item; if a price " +
        "lapses we simply request a fresh quote. These are STANDARD — do NOT flag them: (a) removing supplier " +
        "markings from material (we ask EVERY supplier for this, so a supplier merely restating it is not " +
        "noteworthy); (b) cut mill stock (material cut down from larger mill lengths by the service center is " +
        "normal and always acceptable — never raise it as a quality/acceptability concern). The ONLY thing to " +
        "watch on cut items is the DELIVERY date (the extra processing can push it out), which the quote's " +
        "delivery/ship/due date already captures. Be concrete: supplier names + dollar figures, " +
        "no filler, no preamble or sign-off.\n" +
        "PRICE TABLE: present the price comparison as a compact table that ALWAYS includes a price-per-pound " +
        "($/lb) column next to the total, so suppliers are comparable on a unit basis. Compute $/lb from the " +
        "total and the line weight. THEORETICAL WEIGHT: a line may show 'weight ~N lb (ESTIMATED)' — we " +
        "computed that weight from the product's dimensions because the supplier did not state one. Any $/lb " +
        "you derive using an ESTIMATED weight is THEORETICAL: append a '*' to that $/lb value and add ONE " +
        "footnote — '* theoretical $/lb — based on our calculated weight; supplier gave no weight'. CRUCIAL: a " +
        "$/lb is theoretical ONLY when you DIVIDED a total (or per-piece/per-foot price) by an ESTIMATED " +
        "weight. If the supplier quoted a price PER POUND directly, that IS their own $/lb — NEVER mark it " +
        "theoretical, even when the weight shown is ESTIMATED (the estimate only affects the weight column, " +
        "not their stated $/lb).\n" +
        "TIMING: replies usually arrive within ~30 min and chasing is only warranted after ~60 min. The " +
        "input says how long ago the RFQ was sent; do NOT treat few/slow responses as a problem under 60 min.";

    /// <summary>Bumped whenever the prompt / output guidance changes, so it folds into the inputs-hash
    /// and existing cached summaries regenerate with the new prompt on next access.</summary>
    private const string PromptVersion = "sop-v8-weight-estimate";

    public record Result(string Summary, string InputsHash, string Model, string Mode);

    /// <summary>Generates the state-of-play from the SLI+SR merged rows for one RFQ (the rows are the
    /// dicts returned by <c>SharePointService.ReadSupplierItemsByRfqIdAsync</c>). Returns null on AI
    /// failure (the caller keeps the prior cached summary / the client keeps its deterministic baseline).</summary>
    public async Task<Result?> GenerateAsync(string rfqId, List<Dictionary<string, object?>> rows,
        bool includePdfs, string? sentAgo, CancellationToken ct)
    {
        var apiKey = _config["Anthropic:ApiKey"];
        if (string.IsNullOrWhiteSpace(apiKey)) return null;

        string mode  = includePdfs ? "pdf" : "text";
        string hash  = ComputeInputsHash(rows, includePdfs);
        string input = BuildInput(rfqId, rows, sentAgo);
        if (string.IsNullOrWhiteSpace(input)) return null;

        var model = _config["Claude:Model"] ?? "claude-sonnet-4-6";
        int maxTokens = _config.GetValue("RfqStateOfPlay:MaxTokens", 1200);

        // v1: text-only content. (PDF mode wires document blocks here in the next phase.)
        var body = JsonSerializer.Serialize(new
        {
            model,
            max_tokens = maxTokens,
            system     = SystemPrompt,
            messages   = new[] { new { role = "user", content = input } },
        });

        try
        {
            using var http = _http.CreateClient();
            http.Timeout = TimeSpan.FromSeconds(_config.GetValue("RfqStateOfPlay:TimeoutSeconds", 90));
            http.DefaultRequestHeaders.Add("x-api-key", apiKey);
            http.DefaultRequestHeaders.Add("anthropic-version", "2023-06-01");

            var resp = await http.PostAsync("https://api.anthropic.com/v1/messages",
                new StringContent(body, Encoding.UTF8, "application/json"), ct);
            if (!resp.IsSuccessStatusCode)
            {
                _log.LogWarning("[StateOfPlay] Claude returned {Status} for {Rfq}", resp.StatusCode, rfqId);
                return null;
            }

            var json = await resp.Content.ReadAsStringAsync(ct);
            using var doc = JsonDocument.Parse(json);
            var text = new StringBuilder();
            foreach (var block in doc.RootElement.GetProperty("content").EnumerateArray())
                if (block.TryGetProperty("type", out var t) && t.GetString() == "text" &&
                    block.TryGetProperty("text", out var txt))
                    text.Append(txt.GetString());

            var summary = text.ToString().Trim();
            return string.IsNullOrWhiteSpace(summary) ? null : new Result(summary, hash, model, mode);
        }
        catch (Exception ex)
        {
            _log.LogWarning(ex, "[StateOfPlay] generate failed for {Rfq}", rfqId);
            return null;
        }
    }

    /// <summary>A stable hash of the price-relevant inputs (+ mode). Regeneration is keyed on this, so a
    /// new/changed quote, a regret, or flipping the PDF switch produces a new hash; OOF / no-response do
    /// not change it (they carry no priced/regret signal here).</summary>
    public string ComputeInputsHash(List<Dictionary<string, object?>> rows, bool includePdfs)
    {
        var parts = rows
            .Select(r => string.Join("|", new[]
            {
                S(r, "SupplierName"), ProductLabel(r), Spec(r),
                S(r, "TotalPrice"), S(r, "PricePerPound"), S(r, "PricePerFoot"), S(r, "PricePerPiece"),
                S(r, "UnitsQuoted"), Bool(r, "IsRegret") ? "regret" : "",
            }))
            .Where(s => s.Replace("|", "").Trim().Length > 0)
            .OrderBy(s => s, StringComparer.Ordinal)
            .ToList();
        var raw = PromptVersion + (includePdfs ? "\npdf\n" : "\ntext\n") + string.Join("\n", parts);
        return Convert.ToHexString(SHA256.HashData(Encoding.UTF8.GetBytes(raw)));
    }

    // ── Input assembly (names + specs + prices + email bodies; NO codes) ─────────────────────────────
    private static string BuildInput(string rfqId, List<Dictionary<string, object?>> rows, string? sentAgo)
    {
        var sb = new StringBuilder();
        sb.Append("RFQ ").Append(rfqId);
        if (!string.IsNullOrWhiteSpace(sentAgo)) sb.Append("  (sent ").Append(sentAgo).Append(')');
        sb.Append('\n');

        foreach (var g in rows.Where(r => !string.IsNullOrWhiteSpace(S(r, "SupplierName")))
                              .GroupBy(r => S(r, "SupplierName"), StringComparer.OrdinalIgnoreCase)
                              .OrderBy(g => g.Key, StringComparer.OrdinalIgnoreCase))
        {
            sb.Append("\n=== Supplier: ").Append(g.Key).Append(" ===\n");
            foreach (var r in g)
            {
                if (Bool(r, "IsRegret")) { sb.Append("  - ").Append(ProductLabel(r)).Append("  REGRET (cannot supply)\n"); continue; }
                sb.Append("  - ").Append(ProductLabel(r));
                var spec = Spec(r);            if (spec.Length > 0)               sb.Append("  [").Append(spec).Append(']');
                var price = PriceText(r);      if (price.Length > 0)              sb.Append("  ").Append(price);
                var (wtLb, wtEst) = ResolveWeight(r);
                if (wtLb is > 0) sb.Append($"  weight ~{wtLb.Value:0.#} lb{(wtEst ? " (ESTIMATED)" : "")}");
                var lead = S(r, "LeadTimeText"); if (lead.Length > 0)             sb.Append("  lead: ").Append(lead);
                var certs = S(r, "Certifications"); if (certs.Length > 0)         sb.Append("  certs: ").Append(certs);
                var note = S(r, "SupplierProductComments"); if (note.Length > 0)  sb.Append("  note: ").Append(Trim(note, 200));
                sb.Append('\n');
            }
            var bodyKey = FirstNonEmpty(g, "EmailBody");
            if (bodyKey.Length > 0)
                sb.Append("  Email: ").Append(Trim(bodyKey.Replace("\r", " ").Replace("\n", " "), 1200)).Append('\n');
        }
        return sb.ToString();
    }

    // The supplier's own product wording — never the catalog/code-derived name.
    private static string ProductLabel(Dictionary<string, object?> r)
    {
        var n = S(r, "SupplierProductName"); if (n.Length > 0) return n;
        return S(r, "ProductName");
    }

    private static string Spec(Dictionary<string, object?> r)
    {
        var bits = new List<string>();
        var qty = S(r, "UnitsQuoted"); var req = S(r, "UnitsRequested");
        if (qty.Length > 0) bits.Add($"qty {qty}{(req.Length > 0 && req != qty ? $" of {req}" : "")}");
        var len = S(r, "LengthPerUnit"); var lu = S(r, "LengthUnit");
        if (len.Length > 0) bits.Add($"len {len}{lu}");
        var wt = S(r, "WeightPerUnit"); var wu = S(r, "WeightUnit");
        if (wt.Length > 0) bits.Add($"wt {wt}{wu}");
        return string.Join(", ", bits);
    }

    private static string PriceText(Dictionary<string, object?> r)
    {
        var bits = new List<string>();
        void Add(string key, string suffix) { var v = S(r, key); if (v.Length > 0) bits.Add($"${v}{suffix}"); }
        Add("TotalPrice", " total"); Add("PricePerPiece", "/pc"); Add("PricePerFoot", "/ft"); Add("PricePerPound", "/lb");
        return string.Join("  ", bits);
    }

    // ── price-impact gates (used by the queue + processor) ───────────────────────────────────────────
    /// <summary>A row is price-impactful if it carries a price OR is a regret — these change the
    /// competitive picture. OOF / no-response rows carry neither and are ignored.</summary>
    public static bool IsImpactful(Dictionary<string, object?> r)
        => Bool(r, "IsRegret")
        || new[] { "TotalPrice", "PricePerPound", "PricePerFoot", "PricePerPiece" }
              .Any(k => { var v = S(r, k); return v.Length > 0 && v != "0" && v != "0.0"; });

    public static bool AnyImpactful(IEnumerable<Dictionary<string, object?>> rows) => rows.Any(IsImpactful);

    /// <summary>Distinct suppliers that have a priced/regret line — the state-of-play only runs at ≥2.</summary>
    public static int CompetingSuppliers(IEnumerable<Dictionary<string, object?>> rows)
        => rows.Where(IsImpactful).Select(r => S(r, "SupplierName"))
               .Where(s => s.Length > 0).Distinct(StringComparer.OrdinalIgnoreCase).Count();

    // ── dict helpers ────────────────────────────────────────────────────────────────────────────────
    private static string S(Dictionary<string, object?> d, string k)
        => d.TryGetValue(k, out var v) && v is not null ? v.ToString()!.Trim() : "";
    private static bool Bool(Dictionary<string, object?> d, string k)
        => d.TryGetValue(k, out var v) && v is not null && (v is bool b ? b : string.Equals(v.ToString(), "true", StringComparison.OrdinalIgnoreCase));
    private static string FirstNonEmpty(IEnumerable<Dictionary<string, object?>> rows, string k)
        => rows.Select(r => S(r, k)).FirstOrDefault(s => s.Length > 0) ?? "";
    private static string Trim(string s, int max) => s.Length <= max ? s : s[..max] + "…";

    // ── weight resolution (supplier-stated, else computed from dimensions via WeightCalculator) ───────
    /// <summary>Line total weight in lb + whether it was ESTIMATED from dimensions because the supplier
    /// gave no weight (same calculator QC + the catalog use), so the model can mark such $/lb theoretical.</summary>
    private static (double? TotalLb, bool Estimated) ResolveWeight(Dictionary<string, object?> r)
    {
        double qty = ParseD(S(r, "UnitsQuoted")) ?? 1;
        if (qty <= 0) qty = 1;

        var sw = ParseD(S(r, "WeightPerUnit"));
        if (sw is > 0) return (ToLb(sw.Value, S(r, "WeightUnit")) * qty, false);

        var wc = WeightCalculator.Calculate(ProductLabel(r));
        if (wc.LbPerFoot is > 0)
        {
            var lenFt = ToFeet(ParseD(S(r, "LengthPerUnit")), S(r, "LengthUnit"));
            if (lenFt is > 0) return (wc.LbPerFoot.Value * lenFt.Value * qty, true);
        }
        return (null, false);
    }

    private static double? ParseD(string s) => double.TryParse(s, out var d) ? d : null;

    private static double? ToFeet(double? v, string unit)
    {
        if (v is not > 0) return null;
        return unit.ToLowerInvariant() switch
        {
            "in" or "inch" or "inches" or "\"" => v / 12.0,
            "mm" => v / 304.8, "cm" => v / 30.48, "m" => v * 3.28084,
            _ => v,   // ft / blank
        };
    }

    private static double ToLb(double v, string unit) => unit.ToLowerInvariant() switch
    {
        "kg" => v * 2.20462, "g" => v / 453.592, "oz" => v / 16.0,
        _ => v,   // lb / blank
    };
}
