using System.Globalization;
using System.Security.Cryptography;
using System.Text;
using System.Text.Json;
using System.Text.RegularExpressions;

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
        "that line. The per-line WINNERS, the optimal split total, and each supplier's full-order total are " +
        "GIVEN to you in the 'DETERMINISTIC WINNERS' block at the end of the input — the system computed them " +
        "from the prices and they are AUTHORITATIVE. Use them EXACTLY: bold the given winner in your table, " +
        "quote those totals verbatim, and NEVER recompute, second-guess, or override the winner or the math. " +
        "Spend PROSE only on a supplier that WINS at least one line — do NOT write paragraphs about suppliers " +
        "who win nothing; the PRICE TABLE itself, however, ALWAYS lists EVERY supplier that responded (see PRICE TABLE).\n" +
        "ALLOY DIFFERENCES ARE OK: we routinely request one alloy and accept an interchangeable one that " +
        "prices better and still works for the customer, so suppliers on the SAME line may carry different " +
        "alloys (and match a different MSPC than requested) — that is FINE, keep them in the one comparison " +
        "and pick the cheapest. When the DETERMINISTIC WINNERS block tags a line 'ALLOY VARIES', add a SHORT " +
        "note naming the alloys (e.g. '6063-T52 vs 6061-T6') so the rep can confirm the substitution — do NOT " +
        "split the line by alloy or treat it as a problem.\n" +
        "NEGOTIATION: when a supplier wins some lines but loses others, flag the play to push their LOSING " +
        "lines down (using their winning lines as leverage) so awarding them the whole order beats the split " +
        "— name the lines and the dollar gap to close. Only mention leverage when it ACTUALLY exists (2+ " +
        "suppliers competing on a line); NEVER state that there is no leverage or that an item is single-sourced.\n" +
        "OMIT THE OBVIOUS: never state the ABSENCE or non-existence of something, and never restate the setup. " +
        "Do NOT write sentences like 'no single-sourced items', 'this is one line', 'fully covered by N " +
        "respondents', 'no split needed', 'all suppliers quoted everything', or 'no leverage here'. If a topic " +
        "has nothing ACTIONABLE to say, write nothing about it — silence, not a sentence confirming the obvious.\n" +
        "SMALL GAPS → CONSOLIDATE, DON'T SPLIT: when a supplier already leads on most / the highest-value lines " +
        "and the lines they LOSE are each within the SMALL-GAP THRESHOLD given in the input, do NOT recommend a " +
        "split — recommend awarding that supplier the WHOLE order and asking for a price reduction on those " +
        "near-miss lines (name the lines + the small gap to close). Only split a line out when its gap is " +
        "MATERIAL (greater than the threshold). Always show the CONSOLIDATION PREMIUM (single-supplier total " +
        "minus the best split) so the magnitude is visible.\n" +
        "WINNER vs RECOMMENDATION ARE DIFFERENT — do NOT conflate them: the per-line 'Winner' (in the table " +
        "and any 'wins line X' wording) is ALWAYS the deterministic CHEAPEST supplier on that line, even when " +
        "you RECOMMEND consolidating the award to someone else. NEVER mark the consolidation/recommended " +
        "supplier as the 'winner' of a line they are not cheapest on, and NEVER write 'X wins every line' " +
        "unless X is genuinely cheapest on every line. Phrase consolidation as: 'recommend awarding the whole " +
        "order to X even though Y wins lines A and B (small gaps) — close $N to consolidate.'\n" +
        "SHIPPING applies to the FULL order, not per line: where a supplier states a freight DOLLAR amount, add " +
        "a SHIPPING row to the table and a delivered total (line items + shipping), and FACTOR IT INTO the " +
        "winner — even when the line items are all within the threshold, one supplier's shipping can swing the " +
        "full order materially. If only freight TERMS are given (e.g. FOB origin/destination, collect/prepaid), " +
        "state the terms; if shipping is unknown for a supplier, note it briefly so it isn't forgotten.\n" +
        "WRITE TIGHTLY: a 1-line recommendation (who, and why), then a handful of short lines on the price " +
        "leader, the genuine trade-offs, and gaps/risks. Mention a dimension (lead time, freight terms, " +
        "payment terms, certs, MOQ, surcharges) ONLY when it MEANINGFULLY differs between suppliers — skip " +
        "it otherwise. MTRs (material test reports) are ALWAYS required and supplied as standard here — NEVER " +
        "note a missing or unmentioned MTR as a gap or risk; only raise certs if a supplier explicitly cannot " +
        "provide standard MTRs or offers something beyond them. IGNORE quote validity / expiry dates and " +
        "times — never flag a short, burning, or expiring quote as a risk or an action item; if a price " +
        "lapses we simply request a fresh quote. The following are STANDARD and carry ZERO signal — NEVER " +
        "mention them ANYWHERE in the output: not as a risk, not as a trade-off, not as a minor note, not as a " +
        "'flag for the shop'. OMIT them ENTIRELY even if a supplier's quote explicitly states them: (1) a " +
        "supplier asking that THEIR markings be removed from the material — we require this of EVERY supplier, " +
        "it is routine and expected; (2) cut mill stock (material cut down from larger mill lengths by the " +
        "service center) — normal and always acceptable, never a quality/acceptability concern; the only thing " +
        "that matters on cut items is the DELIVERY date, which the quote's delivery/ship/due date already " +
        "captures. Do not write a sentence about either of these topics. Be concrete: supplier names + dollar figures, " +
        "no filler, no preamble or sign-off.\n" +
        "PRICE TABLE: present the comparison as ONE table — a row per line item, and a COLUMN FOR EVERY " +
        "supplier that responded. INCLUDE suppliers who win nothing, and suppliers who quoted only SOME lines " +
        "(show their price where they quoted, and 'regret' or '—' in the lines they did not). NEVER drop a " +
        "supplier from the table just because they lose. In EACH supplier price cell put that supplier's TOTAL " +
        "for the line on the first row and their price-per-pound ($/lb) directly beneath it in a smaller font — " +
        "write the cell LITERALLY as '$TOTAL<br><small>$X.XX/lb</small>' (emit the <br> and <small> tags " +
        "verbatim; they render in the client). BOLD the cheapest (winning) supplier's cell on each line, e.g. " +
        "'**$1,548.00**<br><small>$2.58/lb</small>' — that bolded cell IS the winner marker, so you do NOT need " +
        "a separate 'Winner' column. Add a SHIPPING row and a delivered-total row wherever a supplier states a " +
        "freight dollar amount. Each line gives a PRE-COMPUTED per-pound value as '=> USE $/lb N.NN' — put THAT " +
        "exact number in the cell's $/lb sub-line and do NOT recompute it (the raw per-piece/-foot/-pound fields " +
        "are frequently mis-slotted; ignore them for the $/lb). A trailing '*' on the value means THEORETICAL " +
        "(we derived the weight from the product's dimensions because the supplier stated none): carry the '*' " +
        "onto that cell and add ONE footnote '* theoretical $/lb — based on our calculated weight; supplier gave " +
        "no weight'. If a line shows no '=> USE $/lb' at all, omit the $/lb sub-line for that cell.\n" +
        "BRIEF + ACTIONABLE: every line must be a MOVE the rep can make (award, negotiate, chase, confirm) — " +
        "not information for its own sake. Explain ONLY where the recommendation genuinely needs the " +
        "justification to stand; otherwise cut it. If a line doesn't change what the rep does, delete it.\n" +
        "TIMING: replies usually arrive within ~30 min and chasing is only warranted after ~60 min. The " +
        "input says how long ago the RFQ was sent; do NOT treat few/slow responses as a problem under 60 min.";

    /// <summary>Bumped whenever the prompt / output guidance changes, so it folds into the inputs-hash
    /// and existing cached summaries regenerate with the new prompt on next access.</summary>
    private const string PromptVersion = "sop-v17-precomputed-perlb";

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
        int    gap   = _config.GetValue("RfqStateOfPlay:LineGapThreshold", 10);
        string input = BuildInput(rfqId, rows, sentAgo, gap);
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
    private static string BuildInput(string rfqId, List<Dictionary<string, object?>> rows, string? sentAgo, int gapThreshold)
    {
        var sb = new StringBuilder();
        sb.Append("RFQ ").Append(rfqId);
        if (!string.IsNullOrWhiteSpace(sentAgo)) sb.Append("  (sent ").Append(sentAgo).Append(')');
        sb.Append('\n');
        sb.Append("Small-gap threshold: $").Append(gapThreshold)
          .Append(" (a line lost by <= $").Append(gapThreshold).Append(" is a SMALL gap — consolidate, don't split)\n");

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
                if (wtLb is > 0)
                {
                    sb.Append($"  weight ~{wtLb.Value:0.#} lb{(wtEst ? " (ESTIMATED)" : "")}");
                    // Pre-compute the $/lb DETERMINISTICALLY (total ÷ resolved weight) so the model never has
                    // to derive it from the raw per-unit fields, which are frequently mis-slotted. The trailing
                    // '*' (when the weight was estimated) is the authoritative theoretical marker.
                    var totP = ParseD(S(r, "TotalPrice"));
                    if (totP is > 0) sb.Append($"  => USE $/lb {totP.Value / wtLb.Value:0.00}{(wtEst ? "*" : "")}");
                }
                var lead = S(r, "LeadTimeText"); if (lead.Length > 0)             sb.Append("  lead: ").Append(lead);
                var certs = S(r, "Certifications"); if (certs.Length > 0)         sb.Append("  certs: ").Append(certs);
                var note = S(r, "SupplierProductComments"); if (note.Length > 0)  sb.Append("  note: ").Append(Trim(note, 200));
                sb.Append('\n');
            }
            var freight = FirstNonEmpty(g, "FreightTerms");
            if (freight.Length > 0) sb.Append("  Freight/shipping: ").Append(Trim(freight, 200)).Append('\n');
            var bodyKey = FirstNonEmpty(g, "EmailBody");
            if (bodyKey.Length > 0)
                sb.Append("  Email: ").Append(Trim(bodyKey.Replace("\r", " ").Replace("\n", " "), 1200)).Append('\n');
        }
        sb.Append(ComputeWinnerBlock(rows));
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
        => WeightCalculator.ResolveLineWeightLb(
            ParseD(S(r, "UnitsQuoted")) ?? 1,
            ParseD(S(r, "WeightPerUnit")), S(r, "WeightUnit"),
            S(r, "CatalogProductName"), ProductLabel(r),
            ParseD(S(r, "LengthPerUnit")), S(r, "LengthUnit"));

    private static double? ParseD(string s) => double.TryParse(s, out var d) ? d : null;

    // Unit converters delegate to the single source in WeightCalculator (also used by EffectiveTotal).
    private static double? ToFeet(double? v, string unit) => WeightCalculator.ToFeet(v, unit);
    private static double  ToLb(double v, string unit)    => WeightCalculator.ToLb(v, unit);

    // ── deterministic per-line winners (group by metal+shape+dims, alloy-agnostic; NO codes) ───────────
    /// <summary>System-computed authoritative comparison so the model never miscalculates the cheapest
    /// supplier: per-line winner + runner-up gap, the optimal line-by-line split, and each supplier's
    /// full-order total (suppliers that priced every line). Lines are grouped by METAL + shape + dimensions
    /// (alloy-agnostic) so suppliers who quoted an interchangeable alloy — even one matching a different MSPC
    /// than requested — stay in ONE pool; lines where the alloy differs are FLAGGED, not split. The block
    /// shows product NAMES, never codes, so the model stays code-free.</summary>
    private static string ComputeWinnerBlock(List<Dictionary<string, object?>> rows, bool dimsAware = true)
    {
        var items = new List<(string Sup, string Label, string Mspc, double Tot, string Alloy, string Key)>();
        foreach (var r in rows)
        {
            if (Bool(r, "IsRegret")) continue;
            var sup = S(r, "SupplierName"); if (sup.Length == 0) continue;
            var tot = EffectiveTotal(r);    if (tot is not > 0) continue;
            var label = ProductLabel(r);
            items.Add((sup, label, S(r, "ProductSearchKey"), tot.Value, AlloySig(label), LineKey(label)));
        }
        if (items.Count == 0) return "";

        // Pool rows that share an MSPC (robust to wording/length/cut variance WITHIN an alloy) OR a
        // metal+shape+dimension key (alloy-agnostic — merges interchangeable-alloy variants that took a
        // different MSPC). Union-find makes the relation transitive so both effects compose.
        var parent = Enumerable.Range(0, items.Count).ToArray();
        int Find(int x) { while (parent[x] != x) { parent[x] = parent[parent[x]]; x = parent[x]; } return x; }
        void Union(int a, int b) { parent[Find(a)] = Find(b); }
        if (dimsAware)
        {
            //   • FLAT (sheet/plate): the MSPC is size-agnostic (thickness+alloy only), so pool by a robust
            //     SIZE SIGNATURE (thickness + W×L, all decimal-inch) across ALL flat rows regardless of
            //     MSPC/alloy/finish — so the same sheet quoted under different MSPCs or alloys still pools,
            //     while 48x120 / 60x144 / 72x144 stay separate.
            //   • LONG (tube/pipe/bar/angle/channel/beam/rebar): the MSPC already fixes the cross-section and
            //     the length is FUNGIBLE, so pool by MSPC alone (PA Steel 288" / Eastern 24' / Hadco 289").
            var flatByMetal = new Dictionary<string, List<int>>(StringComparer.OrdinalIgnoreCase);
            var longByMspc  = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);
            for (int i = 0; i < items.Count; i++)
            {
                if (items[i].Key.Split('|').ElementAtOrDefault(1) == "plate")
                {
                    var bk = items[i].Key.Split('|')[0];   // metal only (alloy stripped) -> alloy-agnostic
                    (flatByMetal.TryGetValue(bk, out var l) ? l : flatByMetal[bk] = new()).Add(i);
                }
                else if (items[i].Mspc.Length > 0)
                {
                    if (longByMspc.TryGetValue(items[i].Mspc, out var j)) Union(i, j); else longByMspc[items[i].Mspc] = i;
                }
            }
            foreach (var g in flatByMetal.Values)
                for (int a = 0; a < g.Count; a++)
                    for (int b = a + 1; b < g.Count; b++)
                        if (SizeSigMatch(items[g[a]].Key, items[g[b]].Key)) Union(g[a], g[b]);
        }
        else
        {
            // legacy: pool by MSPC alone — over-merges different sizes that share a size-agnostic MSPC.
            var byMspc = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);
            for (int i = 0; i < items.Count; i++)
                if (items[i].Mspc.Length > 0)
                {
                    if (byMspc.TryGetValue(items[i].Mspc, out var j)) Union(i, j); else byMspc[items[i].Mspc] = i;
                }
        }
        // (B) alloy-agnostic merge: same metal+shape+dimension LineKey (both modes).
        var byKey = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);
        for (int i = 0; i < items.Count; i++)
            if (byKey.TryGetValue(items[i].Key, out var k)) Union(i, k); else byKey[items[i].Key] = i;

        var poolSup    = new Dictionary<int, Dictionary<string, double>>();   // pool -> supplier -> cheapest
        var poolLabel  = new Dictionary<int, string>();
        var poolAlloys = new Dictionary<int, SortedSet<string>>();
        for (int i = 0; i < items.Count; i++)
        {
            var root = Find(i);
            if (!poolSup.TryGetValue(root, out var d))
            {
                poolSup[root] = d = new(StringComparer.OrdinalIgnoreCase);
                poolLabel[root]  = Trim(items[i].Label, 48);
                poolAlloys[root] = new(StringComparer.OrdinalIgnoreCase);
            }
            if (!d.TryGetValue(items[i].Sup, out var cur) || items[i].Tot < cur) d[items[i].Sup] = items[i].Tot;
            if (items[i].Alloy.Length > 0) poolAlloys[root].Add(items[i].Alloy);
        }

        var sb = new StringBuilder();
        sb.Append("\nDETERMINISTIC WINNERS (system-computed from the prices — AUTHORITATIVE; present these, do NOT recompute or override the winner or the totals):\n");
        var pools = poolSup.Keys.ToList();
        double split = 0;
        foreach (var root in pools)
        {
            var ranked = poolSup[root].OrderBy(kv => kv.Value).ToList();
            var win = ranked[0]; split += win.Value;
            sb.Append("  ").Append(poolLabel[root]).Append("  ->  winner ").Append(win.Key).Append(" $").Append(win.Value.ToString("0.00"));
            if (ranked.Count > 1)
                sb.Append("  (next ").Append(ranked[1].Key).Append(" $").Append(ranked[1].Value.ToString("0.00"))
                  .Append(", gap $").Append((ranked[1].Value - win.Value).ToString("0.00")).Append(')');
            if (poolAlloys[root].Count > 1)
                sb.Append("  [ALLOY VARIES: ").Append(string.Join(" vs ", poolAlloys[root])).Append(" — interchangeable, flag for user]");
            sb.Append('\n');
        }
        sb.Append("  Optimal line-by-line split total = $").Append(split.ToString("0.00")).Append('\n');
        var allSups = poolSup.Values.SelectMany(d => d.Keys).Distinct(StringComparer.OrdinalIgnoreCase);
        var full = allSups
            .Where(s => pools.All(p => poolSup[p].ContainsKey(s)))
            .Select(s => (Sup: s, Tot: pools.Sum(p => poolSup[p][s])))
            .OrderBy(x => x.Tot).ToList();
        if (full.Count > 0)
            sb.Append("  Single-supplier full-order totals (priced every line): ")
              .Append(string.Join(", ", full.Select(f => $"{f.Sup} ${f.Tot:0.00}"))).Append('\n');
        return sb.ToString();
    }

    /// <summary>Deterministic effective line total: the supplier's TotalPrice, else derived from per-piece /
    /// per-foot / per-pound (using the supplier or estimated weight).</summary>
    private static double? EffectiveTotal(Dictionary<string, object?> r)
    {
        var t = ParseD(S(r, "TotalPrice")); if (t is > 0) return t;
        double qty = ParseD(S(r, "UnitsQuoted")) ?? 1; if (qty <= 0) qty = 1;
        var pc = ParseD(S(r, "PricePerPiece")); if (pc is > 0) return pc.Value * qty;
        var pf = ParseD(S(r, "PricePerFoot"));
        var lenFt = ToFeet(ParseD(S(r, "LengthPerUnit")), S(r, "LengthUnit"));
        if (pf is > 0 && lenFt is > 0) return pf.Value * qty * lenFt.Value;
        var pl = ParseD(S(r, "PricePerPound"));
        var (wt, _) = ResolveWeight(r);
        if (pl is > 0 && wt is > 0) return pl.Value * wt.Value;
        return null;
    }

    /// <summary>Canonical line key for the winner pool: METAL-family + shape + sorted numeric dimensions.
    /// The metal family (aluminum / stainless / galv / steel) STAYS in the key, but the specific ALLOY and
    /// temper within a family do NOT — so alloy-only differences land in the SAME pool (and get flagged via
    /// <see cref="AlloySig"/>) rather than splitting it. Feet are normalized to inches so 20' and 240"
    /// group together. No MSPC/code is used.</summary>
    /// <summary>Deterministic winner block for diagnostics/backtesting (no AI). dimsAware=false reproduces
    /// the legacy MSPC-alone pooling so the new dimension-aware pooling can be A/B-compared per RFQ.</summary>
    public string WinnerBlockForDiag(List<Dictionary<string, object?>> rows, bool dimsAware = true)
        => ComputeWinnerBlock(rows, dimsAware);

    /// <summary>Robust SIZE SIGNATURE of a flat (sheet/plate) LineKey: (thickness, face-dim-1, face-dim-2) in
    /// decimal inches. Picks the thinnest plausible thickness (0.005-4") and the two largest plausible face
    /// dims (4-600") — so finish codes ("2B"→2, "#4"→4) and internal product codes ("4228") are ignored, and
    /// gauge/fractions (already converted to decimal in <see cref="LineKey"/>) line up. Lets the SAME sheet
    /// pool across different MSPCs / alloys / finishes.</summary>
    private static (double Th, double D1, double D2) SizeSig(string key)
    {
        var dims = (key.Contains('|') ? key[(key.LastIndexOf('|') + 1)..] : key)
            .Split(',', StringSplitOptions.RemoveEmptyEntries)
            .Select(t => double.TryParse(t, NumberStyles.Any, CultureInfo.InvariantCulture, out var d) ? d : double.NaN)
            .Where(d => !double.IsNaN(d)).ToList();
        var th = dims.Where(d => d is >= 0.005 and <= 4).DefaultIfEmpty(0).Min();
        var wl = dims.Where(d => d is > 4 and < 600).OrderByDescending(d => d).Take(2).OrderBy(d => d).ToList();
        return (th, wl.ElementAtOrDefault(0), wl.ElementAtOrDefault(1));
    }

    private static bool SizeSigMatch(string a, string b)
    {
        var (ta, a1, a2) = SizeSig(a);
        var (tb, b1, b2) = SizeSig(b);
        static bool Close(double x, double y) => (x == 0 && y == 0) || Math.Abs(x - y) <= 0.08 * Math.Max(Math.Abs(x), Math.Abs(y));
        if (!Close(a1, b1) || !Close(a2, b2)) return false;       // face W×L must match
        if (ta > 0 && tb > 0 && !Close(ta, tb)) return false;     // thickness must match when BOTH state one
        return true;
    }

    // Sheet-metal gauge -> decimal inches (material-specific). The tables live in the shared
    // DimensionNormalizer so the winner pool and the catalog tokenizer convert gauge on ONE basis.
    private static double? GaugeToInches(string metal, bool tube, int ga)
        => DimensionNormalizer.GaugeToInches(metal, tube, ga);

    internal static string LineKey(string name)
    {
        var raw = name.ToLowerInvariant();
        string metal =
            (raw.Contains("brass")  || Regex.IsMatch(raw, @"\bc(?:2\d{2}|3\d{2}|46[024])\b"))             ? "brass" :
            (raw.Contains("copper") || Regex.IsMatch(raw, @"\bc1\d{2}\b"))                                ? "copper":
            (raw.Contains("galvaniz") || Regex.IsMatch(raw, @"\bg-?90\b"))                                ? "galv"  :
            (raw.Contains("stainless") || Regex.IsMatch(raw, @"\bt?3(?:04|16|21)l?\b") || Regex.IsMatch(raw, @"\b4(?:10|30)\b")) ? "ss" :
            (raw.Contains("alum") || Regex.IsMatch(raw, @"\b(?:1100|2024|3003|5052|5086|6061|6063|7075)\b")) ? "alum" :
            "steel";
        // Normalize feet -> inches so 20' and 240" group together.
        var n = Regex.Replace(raw, @"(\d+(?:\.\d+)?)\s*(?:'|’|ft\b|feet\b|foot\b)", m =>
            double.TryParse(m.Groups[1].Value, NumberStyles.Any, CultureInfo.InvariantCulture, out var ft)
                ? (ft * 12).ToString("0.###", CultureInfo.InvariantCulture) : m.Value);
        // Normalize sheet-metal gauge -> decimal inches (material-specific) so "20 Gauge" / "7GA" compare
        // on the same basis as "3/16" / "0.188". Done before the strips so the gauge digit never leaks as a dim.
        bool isTube = raw.Contains("tube") || raw.Contains("pipe");   // tube/pipe wall uses BWG, not sheet gauge
        n = Regex.Replace(n, @"#?\s*(\d{1,2})\s*(?:gauge|ga)\b", m =>
            int.TryParse(m.Groups[1].Value, out var ga) && GaugeToInches(metal, isTube, ga) is double th
                ? th.ToString("0.###", CultureInfo.InvariantCulture) : " ");
        // Strip metallurgical grade / temper / coating tokens — they are NOT dimensions.
        n = Regex.Replace(n, @"\ba\d{2,3}(\s*/\s*a?\d{2,3})?\b", " ");          // A36, A500, A500/A513, A572/A992
        n = Regex.Replace(n, @"\b(304|316|321|410|430)l?\b",     " ");          // stainless series
        n = Regex.Replace(n, @"\b(1100|2024|3003|5052|5086|6061|6063|7075)\b", " ");  // aluminum series
        n = Regex.Replace(n, @"\bt\d{1,4}\b",                    " ");          // tempers T5/T6/T52/T6511
        n = Regex.Replace(n, @"\bh\d{2,3}\b",                    " ");          // tempers H14/H32/H34/H112/H321 (else 14 leaks as a dim)
        n = Regex.Replace(n, @"\bg-?90\b",                       " ");          // galv coating
        string shape =
            n.Contains("angle")                                            ? "angle"   :
            (n.Contains("wide flange") || n.Contains("w beam") || n.Contains("beam") || Regex.IsMatch(n, @"\bw\d")) ? "beam" :
            n.Contains("channel")                                          ? "channel" :
            n.Contains("pipe")                                             ? "pipe"    :
            n.Contains("tube")                                             ? "tube"    :
            n.Contains("rebar")                                            ? "rebar"   :
            (n.Contains("plate") || n.Contains("sheet"))                   ? "plate"   :
            (n.Contains("flat") || n.Contains("bar"))                      ? "bar"     : "other";
        var nums = new List<double>();
        foreach (Match m in Regex.Matches(n, @"(\d+)\s*/\s*(\d+)"))
            if (double.TryParse(m.Groups[1].Value, out var a) && double.TryParse(m.Groups[2].Value, out var b) && b != 0)
                nums.Add(Math.Round(a / b, 3));
        foreach (Match m in Regex.Matches(Regex.Replace(n, @"\d+\s*/\s*\d+", " "), @"\.\d+|\d+(?:\.\d+)?"))
            if (double.TryParse(m.Value, NumberStyles.Any, CultureInfo.InvariantCulture, out var d)) nums.Add(Math.Round(d, 3));
        nums.Sort();
        return metal + "|" + shape + "|" + string.Join(",", nums.Select(x => x.ToString("0.###")));
    }

    /// <summary>Coarse alloy/temper signature from a product name (e.g. "6063/t52", "a500/a513", "304") so
    /// the winner block can FLAG when otherwise-matching suppliers quoted different (interchangeable) alloys.
    /// Empty when no recognizable grade is present.</summary>
    internal static string AlloySig(string name)
    {
        var n = name.ToLowerInvariant();
        var sigs = new SortedSet<string>(StringComparer.OrdinalIgnoreCase);
        foreach (Match m in Regex.Matches(n, @"\ba\d{2,3}(?:\s*/\s*a?\d{2,3})?\b")) sigs.Add(Regex.Replace(m.Value, @"\s+", ""));
        foreach (Match m in Regex.Matches(n, @"\b(?:1100|2024|3003|5052|5086|6061|6063|7075|304|316|321|410|430)l?\b")) sigs.Add(m.Value);
        foreach (Match m in Regex.Matches(n, @"\bt\d{1,4}\b")) sigs.Add(m.Value);
        return string.Join("/", sigs);
    }

    // ── MSPC mismatch analysis (custom CUST_ vs catalog — split winner pools) ─────────────────────────
    /// <summary>Within one RFQ, groups SLI rows by the parsed dimension key and reports any group whose
    /// rows carry MORE THAN ONE distinct ProductSearchKey (MSPC) — i.e. the same physical product matched
    /// to different codes across suppliers (mixed custom CUST_ + catalog, or several distinct CUST_
    /// hashes), which splits the winner pool. Empty list = every line's MSPC is consistent.</summary>
    public static List<MspcMismatch> MspcMismatches(List<Dictionary<string, object?>> rows)
    {
        var result = new List<MspcMismatch>();
        foreach (var grp in rows.Where(r => S(r, "SupplierName").Length > 0)
                                .GroupBy(r => LineKey(ProductLabel(r))))
        {
            var members = grp.Select(r => new MspcMember(
                                 S(r, "SupplierName"), Trim(ProductLabel(r), 70),
                                 S(r, "ProductSearchKey"), S(r, "CatalogProductName")))
                             .ToList();
            var distinct = members.Select(m => m.Mspc).Where(s => s.Length > 0)
                                  .Distinct(StringComparer.OrdinalIgnoreCase).ToList();
            if (distinct.Count < 2) continue;   // consistent — not a mismatch
            int cust = distinct.Count(m => m.StartsWith("CUST_", StringComparison.OrdinalIgnoreCase));
            result.Add(new MspcMismatch(grp.Key, members[0].Product, distinct,
                                        MixedCustomAndCatalog: cust > 0 && cust < distinct.Count,
                                        MultipleCustom: cust > 1, members));
        }
        return result;
    }

    public record MspcMember(string Supplier, string Product, string Mspc, string CatalogName);
    public record MspcMismatch(string LineKey, string LineLabel, List<string> DistinctMspcs,
                               bool MixedCustomAndCatalog, bool MultipleCustom, List<MspcMember> Members);
}
