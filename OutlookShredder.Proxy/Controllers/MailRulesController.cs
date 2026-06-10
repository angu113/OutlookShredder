using Microsoft.AspNetCore.Mvc;
using OutlookShredder.Proxy.Models;
using OutlookShredder.Proxy.Services;

namespace OutlookShredder.Proxy.Controllers;

/// <summary>
/// Back-office management of the deterministic mail-classification rules (the Tools "Mail Rules"
/// surface). Rules are evaluated BEFORE the AI classifier — a match files the email at full
/// confidence and skips the AI. Evaluation lives in MailRuleEngine; persistence in MailRuleService.
/// </summary>
[ApiController]
[Route("api/mail-rules")]
public sealed class MailRulesController : ControllerBase
{
    private readonly MailRuleService _rules;
    private readonly MailTaxonomyService _taxonomy;
    private readonly SharePointService _sp;
    public MailRulesController(MailRuleService rules, MailTaxonomyService taxonomy, SharePointService sp)
    {
        _rules = rules; _taxonomy = taxonomy; _sp = sp;
    }

    /// <summary>All rules, ordered by priority (ascending = evaluated first).</summary>
    [HttpGet]
    public async Task<IActionResult> List(CancellationToken ct)
        => Ok((await _rules.GetRulesAsync(ct)).OrderBy(r => r.Priority).ToList());

    /// <summary>Create a rule. Returns the new rule id.</summary>
    [HttpPost]
    public async Task<IActionResult> Create([FromBody] MailRule rule, [FromQuery] string? by, CancellationToken ct)
    {
        if (rule is null || string.IsNullOrWhiteSpace(rule.CategoryPath) || rule.Conditions.Count == 0)
            return BadRequest(new { error = "A rule needs a CategoryPath and at least one condition." });
        rule.Id = "";   // server-assigned
        var id = await _rules.AddAsync(rule, by ?? "", ct);
        return Ok(new { id });
    }

    /// <summary>Replace an existing rule by id.</summary>
    [HttpPut("{id}")]
    public async Task<IActionResult> Update(string id, [FromBody] MailRule rule, CancellationToken ct)
    {
        if (rule is null || string.IsNullOrWhiteSpace(rule.CategoryPath) || rule.Conditions.Count == 0)
            return BadRequest(new { error = "A rule needs a CategoryPath and at least one condition." });
        return await _rules.UpdateAsync(id, rule, ct) ? Ok(new { ok = true }) : NotFound();
    }

    /// <summary>Delete a rule by id.</summary>
    [HttpDelete("{id}")]
    public async Task<IActionResult> Delete(string id, CancellationToken ct)
        => await _rules.DeleteAsync(id, ct) ? Ok(new { ok = true }) : NotFound();

    // ── Match lists (named, growable value sets referenced by rule conditions) ──────────

    [HttpGet("lists")]
    public async Task<IActionResult> Lists(CancellationToken ct) => Ok(await _rules.GetListsAsync(ct));

    /// <summary>Upsert a match list by name (e.g. add a new payment-processor domain).</summary>
    [HttpPut("lists/{name}")]
    public async Task<IActionResult> SaveList(string name, [FromBody] MailMatchList list, CancellationToken ct)
    {
        if (list is null) return BadRequest();
        list.Name = name;
        var existed = await _rules.SaveListAsync(list, ct);
        return Ok(new { existed });
    }

    [HttpDelete("lists/{name}")]
    public async Task<IActionResult> DeleteList(string name, CancellationToken ct)
        => await _rules.DeleteListAsync(name, ct) ? Ok(new { ok = true }) : NotFound();

    /// <summary>Idempotently install the baseline protective rules + match lists + payment-processor
    /// sender hints, so genuine supplier confirmations file deterministically and the AI still resolves
    /// the real supplier behind a billing processor. Safe to call repeatedly.</summary>
    [HttpPost("seed-defaults")]
    public async Task<IActionResult> SeedDefaults(CancellationToken ct)
    {
        await _sp.EnsureMailListsAsync();   // self-provision the MailRules / MailMatchLists SP lists (idempotent)
        var created = new List<string>();
        string[] processors = ["unitedtranzactions.com", "enmarksystems.com", "cybersource.com", "intuit.com"];

        await UpsertListMerge("payment-processors", processors, created, ct);
        await UpsertListMerge("mtr-content-indicators",
            ["heat number", "heat no", "certificate of compliance", "mill test report", "tensile strength",
             "yield strength", "elongation", "chemical composition", "astm a"], created, ct);

        // Protective rule: a KNOWN supplier referencing our HSK-PO number is unambiguously an order
        // confirmation -> file at 100% so it never hits the confidence gate. Payment processors aren't
        // known suppliers, so they're excluded and fall through to the AI (with the sender hints below).
        const string ruleName = "Known supplier + our PO -> Order Confirmations";
        var rules = await _rules.GetRulesAsync(ct);
        if (!rules.Any(r => string.Equals(r.Name, ruleName, StringComparison.OrdinalIgnoreCase)))
        {
            await _rules.AddAsync(new MailRule
            {
                Name = ruleName, Enabled = true, Priority = 10, CategoryPath = "Supplier/Order Confirmations",
                Conditions =
                [
                    new() { Signal = MailRuleSignal.SenderIsKnownSupplier, Operator = MailRuleOperator.Equals, Values = ["true"] },
                    new() { Signal = MailRuleSignal.Body, Operator = MailRuleOperator.Regex, Values = [@"HSK-PO\d+"] },
                ],
            }, "seed", ct);
            created.Add($"rule:{ruleName}");
        }

        // Payment-processor sender hints so the AI reads the REAL supplier from the body (invoice vs
        // receipt stays content-driven). Idempotent against the existing sender map.
        var mapped = await _taxonomy.GetSenderDomainsAsync(ct);
        foreach (var d in processors)
            if (!mapped.Contains(d)) { await _taxonomy.AddSenderSupplierHintAsync(d, "", ct); created.Add($"sender-hint:{d}"); }

        return Ok(new { created });
    }

    private async Task UpsertListMerge(string name, string[] values, List<string> created, CancellationToken ct)
    {
        var existing = (await _rules.GetListsAsync(ct)).FirstOrDefault(l => string.Equals(l.Name, name, StringComparison.OrdinalIgnoreCase));
        var merged   = (existing?.Values ?? []).Concat(values).Distinct(StringComparer.OrdinalIgnoreCase).ToList();
        if (existing is null || merged.Count != existing.Values.Count)
        {
            await _rules.SaveListAsync(new MailMatchList { Name = name, Values = merged }, ct);
            created.Add($"list:{name}");
        }
    }
}
