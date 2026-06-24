using System.Text;
using System.Text.Json;
using Microsoft.AspNetCore.Mvc;
using OutlookShredder.Proxy.Models;
using OutlookShredder.Proxy.Services;

namespace OutlookShredder.Proxy.Controllers;

/// <summary>
/// Dev-only eval harness (wip/sidecar/project-classification-eval-harness.md).
/// Runs the real classifier in-process over the human-labeled MailGoldenLabels corpus
/// and reports per-leaf P/R/F1, confusion matrix, and confidence calibration.
/// Zero workflow side effects — never writes to MailClassifications, the bus, or matchers.
/// </summary>
[ApiController]
[Route("api/mail-eval")]
public sealed class MailEvalController : ControllerBase
{
    private readonly MailEvalService _eval;
    private readonly MailTaxonomyService _taxonomy;
    private readonly MailWorkbenchService _workbench;
    private readonly MailRuleService _rules;

    public MailEvalController(MailEvalService eval, MailTaxonomyService taxonomy,
        MailWorkbenchService workbench, MailRuleService rules)
    {
        _eval      = eval;
        _taxonomy  = taxonomy;
        _workbench = workbench;
        _rules     = rules;
    }

    /// <summary>Serves the self-contained labeling + report UI (dev tool). Same-origin → exempt from WS2 auth.</summary>
    [HttpGet("ui")]
    public ContentResult Ui() => Content(MailEvalUiPage.Html, "text/html", Encoding.UTF8);

    /// <summary>Start an eval run in the background and return immediately.</summary>
    [HttpPost("run")]
    public IActionResult Run([FromBody] EvalRunRequest req)
        => Accepted(_eval.StartRun(req));

    /// <summary>Live snapshot: running flag, processed/total, rolling accuracy.</summary>
    [HttpGet("status")]
    public IActionResult Status() => Ok(_eval.GetSnapshot());

    /// <summary>Full metrics report from the most recent completed run (null when none yet).</summary>
    [HttpGet("report")]
    public IActionResult Report()
    {
        var r = _eval.GetReport();
        return r is null ? NoContent() : Ok(r);
    }

    /// <summary>
    /// Read-only inspection of the golden corpus: counts by labeler (bootstrap vs human) + by
    /// category + the rows. No AI, no writes — confirms how complete the labeling is before a run.
    /// </summary>
    [HttpGet("golden")]
    public async Task<IActionResult> Golden(CancellationToken ct)
        => Ok(await _eval.GetGoldenStatusAsync(ct));

    /// <summary>Apply one human correction to a golden row (sets category + LabeledBy, clears bootstrap).</summary>
    [HttpPost("golden")]
    public async Task<IActionResult> PatchGolden([FromBody] GoldenPatchRequest req, CancellationToken ct)
    {
        if (string.IsNullOrWhiteSpace(req.MailItemId) || string.IsNullOrWhiteSpace(req.GoldenCategory))
            return BadRequest(new { error = "mailItemId and goldenCategory are required." });
        await _eval.PatchGoldenAsync(req.MailItemId, req.GoldenCategory, req.Subject, req.FromAddress, req.LabeledBy, ct);
        return Ok(new { ok = true });
    }

    /// <summary>Valid category leaf paths (closed set) for the labeling dropdown.</summary>
    [HttpGet("leaves")]
    public async Task<IActionResult> Leaves(CancellationToken ct)
        => Ok((await _taxonomy.GetLeavesAsync(ct)).Select(l => l.Path).Where(p => !string.IsNullOrEmpty(p)).ToList());

    /// <summary>
    /// Add a new taxonomy leaf from the labeling UI (e.g. "Other/Voicemail"). Writes an SP taxonomy
    /// hint so the leaf reflects everywhere on the next use: the labeling dropdown (this /leaves list),
    /// the classifier prompt, and the coerce/valid-path set — no code deploy. Idempotent: an existing
    /// leaf is a no-op (no redundant hint row). Returns the updated leaf-path list so the UI can
    /// repopulate its picker and select the new leaf.
    /// </summary>
    [HttpPost("leaves")]
    public async Task<IActionResult> AddLeaf([FromBody] AddLeafRequest req, CancellationToken ct)
    {
        var (ok, path, error) = MailTaxonomyService.NormalizeLeafPath(req?.CategoryPath);
        if (!ok) return BadRequest(new { error });

        var existing = await _taxonomy.GetLeavesAsync(ct);
        bool already = existing.Any(l => string.Equals(l.Path, path, StringComparison.OrdinalIgnoreCase));
        if (!already)
            await _taxonomy.AddLeafHintAsync(path, req?.Description, "eval-ui", ct);

        var leaves = (await _taxonomy.GetLeavesAsync(ct))
            .Select(l => l.Path).Where(p => !string.IsNullOrEmpty(p)).ToList();
        return Ok(new { ok = true, path, added = !already, leaves });
    }

    public sealed class AddLeafRequest
    {
        public string  CategoryPath { get; set; } = "";
        /// <summary>Optional one-line guidance fed to the classifier prompt so the AI can target the leaf.</summary>
        public string? Description  { get; set; }
    }

    /// <summary>One captured item's detail incl. body, for the labeler to read while correcting.</summary>
    [HttpGet("item/{mailItemId}")]
    public async Task<IActionResult> Item(string mailItemId, CancellationToken ct)
    {
        var d = await _workbench.GetItemDetailAsync(mailItemId, ct);
        return d is null ? NotFound(new { error = "Item not found." }) : Ok(d);
    }

    // ── Rule-improvement loop (reuse the deterministic MailRuleEngine) ───────────────────

    /// <summary>Current deterministic rules (pass-through to MailRuleService so the UI stays under the
    /// exempt /api/mail-eval prefix rather than exposing /api/mail-rules to the browser).</summary>
    [HttpGet("rules")]
    public async Task<IActionResult> Rules(CancellationToken ct) => Ok(await _rules.GetRulesAsync(ct));

    /// <summary>Create a rule (e.g. from a human correction). Body = MailRule; ?by=labeler.</summary>
    [HttpPost("rules")]
    public async Task<IActionResult> AddRule([FromBody] MailRule rule, [FromQuery] string? by, CancellationToken ct)
    {
        if (string.IsNullOrWhiteSpace(rule.CategoryPath) || rule.Conditions.Count == 0)
            return BadRequest(new { error = "A rule needs a CategoryPath and at least one condition." });
        var id = await _rules.AddAsync(rule, string.IsNullOrWhiteSpace(by) ? "eval-ui" : by, ct);
        return Ok(new { id });
    }

    /// <summary>Start a deterministic rule-impact run (no AI): re-run the current ruleset over existing
    /// items and report what would flip. dryRun=true previews; false writes the rule result.</summary>
    [HttpPost("rule-impact")]
    public IActionResult RuleImpact([FromBody] RuleImpactRequest req) => Accepted(_eval.StartRuleImpact(req));

    [HttpGet("rule-impact/status")]
    public IActionResult RuleImpactStatus() => Ok(_eval.GetImpactSnapshot());

    [HttpGet("rule-impact/report")]
    public IActionResult RuleImpactReport()
    {
        var r = _eval.GetImpactReport();
        return r is null ? NoContent() : Ok(r);
    }

    /// <summary>Per-item results as JSONL (one JSON object per line).</summary>
    [HttpGet("results")]
    public IActionResult Results()
    {
        var run = _eval.GetResults();
        if (run is null) return NoContent();

        var sb = new StringBuilder();
        var opts = new JsonSerializerOptions { PropertyNamingPolicy = JsonNamingPolicy.CamelCase };
        foreach (var r in run.Items)
            sb.AppendLine(JsonSerializer.Serialize(r, opts));
        return Content(sb.ToString(), "application/x-ndjson", Encoding.UTF8);
    }

    /// <summary>
    /// Bootstrap MailGoldenLabels from the current in-memory AI classifications (one row per item).
    /// The seeded label is the AI's own guess — human correction is REQUIRED before running the eval.
    /// Pass ?overwrite=true to replace existing rows (default: skip existing human corrections).
    /// </summary>
    /// <summary>
    /// overwrite=true  — patch every row.
    /// overwrite=false — skip rows that already exist (default).
    /// resumeOnly=true — patch only rows that exist but are missing Subject (resumes a partial run).
    /// </summary>
    [HttpPost("seed-golden")]
    public IActionResult SeedGolden([FromQuery] bool overwrite = false, [FromQuery] bool resumeOnly = false)
    {
        _ = Task.Run(() => _eval.SeedGoldenFromCurrentsAsync(overwrite, resumeOnly, CancellationToken.None));
        return Accepted(new { message = "Seeding started in background. Watch proxy logs for [MailEval] seed N/total progress." });
    }
}
