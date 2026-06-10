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
    public MailRulesController(MailRuleService rules) => _rules = rules;

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
}
