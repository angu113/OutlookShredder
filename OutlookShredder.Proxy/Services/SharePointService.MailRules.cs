using System.Text.Json;
using System.Text.Json.Serialization;
using Microsoft.Graph.Models;
using OutlookShredder.Proxy.Models;

namespace OutlookShredder.Proxy.Services;

/// <summary>
/// MailRules persistence — the deterministic classification rules managed back-office in Tools and
/// applied (before the AI) by <see cref="MailRuleService"/>. Each rule's Conditions list is stored as
/// JSON in one column (string-enum serialised so reordering the enums never corrupts stored rules).
/// </summary>
public partial class SharePointService
{
    private const string MailMatchListsList = "MailMatchLists";

    private static readonly JsonSerializerOptions RuleJson = new()
    {
        Converters = { new JsonStringEnumConverter() },
    };

    public async Task<List<MailRule>> ReadMailRulesAsync(CancellationToken ct = default)
    {
        var rows = await ReadAllListItemsAsync(MailRulesList,
            ["Title", "RuleId", "Enabled", "Priority", "CategoryPath", "ConditionsJson", "HitCount"], null, ct);
        var rules = new List<MailRule>();
        foreach (var f in rows)
        {
            var id = GetStr(f, "RuleId");
            if (string.IsNullOrEmpty(id)) id = GetStr(f, "__spId");
            if (string.IsNullOrEmpty(id)) continue;
            List<MailRuleCondition> conds = [];
            var cj = GetStr(f, "ConditionsJson");
            if (!string.IsNullOrWhiteSpace(cj))
                try { conds = JsonSerializer.Deserialize<List<MailRuleCondition>>(cj, RuleJson) ?? []; }
                catch (Exception ex) { _log.LogWarning(ex, "[MailRules] bad ConditionsJson for {Id}", id); }
            rules.Add(new MailRule
            {
                Id           = id!,
                Name         = GetStr(f, "Title") ?? "",
                Enabled      = GetBool(f, "Enabled"),
                Priority     = (int)GetDouble(f, "Priority"),
                CategoryPath = GetStr(f, "CategoryPath") ?? "",
                Conditions   = conds,
                HitCount     = (int)GetDouble(f, "HitCount"),
            });
        }
        return rules;
    }

    public async Task<string> WriteMailRuleAsync(MailRule rule, string createdBy, CancellationToken ct = default)
    {
        var siteId = await GetSiteIdAsync();
        var listId = await ResolveListIdAsync(MailRulesList);
        if (string.IsNullOrEmpty(rule.Id)) rule.Id = Guid.NewGuid().ToString("N");
        await GetGraph().Sites[siteId].Lists[listId].Items.PostAsync(new ListItem
        {
            Fields = new FieldValueSet { AdditionalData = new Dictionary<string, object?>
            {
                ["Title"]          = Trunc(rule.Name, 250),
                ["RuleId"]         = rule.Id,
                ["Enabled"]        = rule.Enabled,
                ["Priority"]       = rule.Priority,
                ["CategoryPath"]   = rule.CategoryPath,
                ["ConditionsJson"] = Trunc(JsonSerializer.Serialize(rule.Conditions, RuleJson), 30000),
                ["HitCount"]       = rule.HitCount,
                ["CreatedAt"]      = DateTimeOffset.UtcNow.ToString("o"),
                ["CreatedBy"]      = createdBy,
            } }
        }, cancellationToken: ct);
        _log.LogInformation("[MailRules] rule added: {Name} -> {Cat}", rule.Name, rule.CategoryPath);
        return rule.Id;
    }

    public async Task<bool> UpdateMailRuleAsync(string ruleId, MailRule rule, CancellationToken ct = default)
    {
        var siteId = await GetSiteIdAsync();
        var listId = await ResolveListIdAsync(MailRulesList);
        var spId = await FindMailRuleSpIdAsync(siteId, listId, ruleId, ct);
        if (spId is null) return false;
        await GetGraph().Sites[siteId].Lists[listId].Items[spId].Fields
            .PatchAsync(new FieldValueSet { AdditionalData = new Dictionary<string, object?>
            {
                ["Title"]          = Trunc(rule.Name, 250),
                ["Enabled"]        = rule.Enabled,
                ["Priority"]       = rule.Priority,
                ["CategoryPath"]   = rule.CategoryPath,
                ["ConditionsJson"] = Trunc(JsonSerializer.Serialize(rule.Conditions, RuleJson), 30000),
            } }, cancellationToken: ct);
        return true;
    }

    public async Task<bool> DeleteMailRuleAsync(string ruleId, CancellationToken ct = default)
    {
        var siteId = await GetSiteIdAsync();
        var listId = await ResolveListIdAsync(MailRulesList);
        var spId = await FindMailRuleSpIdAsync(siteId, listId, ruleId, ct);
        if (spId is null) return false;
        await GetGraph().Sites[siteId].Lists[listId].Items[spId].DeleteAsync(cancellationToken: ct);
        _log.LogInformation("[MailRules] rule deleted: {Id}", ruleId);
        return true;
    }

    private async Task<string?> FindMailRuleSpIdAsync(string siteId, string listId, string ruleId, CancellationToken ct)
    {
        var res = await GetGraph().Sites[siteId].Lists[listId].Items.GetAsync(req =>
        {
            req.QueryParameters.Expand = ["fields($select=RuleId)"];
            req.QueryParameters.Filter = $"fields/RuleId eq '{Esc(ruleId)}'";
            req.QueryParameters.Top    = 1;
        }, ct);
        return res?.Value?.FirstOrDefault()?.Id;
    }

    // ── Match lists (named value sets referenced by rule conditions) ──────────────────

    public async Task<List<MailMatchList>> ReadMatchListsAsync(CancellationToken ct = default)
    {
        var rows = await ReadAllListItemsAsync(MailMatchListsList, ["Title", "ValuesJson"], null, ct);
        return rows.Select(f => new MailMatchList
        {
            Name   = GetStr(f, "Title") ?? "",
            Values = SafeList(GetStr(f, "ValuesJson")),
        }).Where(l => l.Name.Length > 0).ToList();
    }

    /// <summary>Upsert a match list by name (case-insensitive). Returns true if it already existed.</summary>
    public async Task<bool> WriteMatchListAsync(MailMatchList list, CancellationToken ct = default)
    {
        var siteId = await GetSiteIdAsync();
        var listId = await ResolveListIdAsync(MailMatchListsList);
        var spId   = await FindMatchListSpIdAsync(list.Name, ct);
        var fields = new Dictionary<string, object?>
        {
            ["Title"]      = Trunc(list.Name, 250),
            ["ValuesJson"] = Trunc(JsonSerializer.Serialize(list.Values), 30000),
        };
        if (spId is null)
        {
            fields["CreatedAt"] = DateTimeOffset.UtcNow.ToString("o");
            await GetGraph().Sites[siteId].Lists[listId].Items
                .PostAsync(new ListItem { Fields = new FieldValueSet { AdditionalData = fields } }, cancellationToken: ct);
            _log.LogInformation("[MailRules] match list created: {Name} ({N})", list.Name, list.Values.Count);
            return false;
        }
        await GetGraph().Sites[siteId].Lists[listId].Items[spId].Fields
            .PatchAsync(new FieldValueSet { AdditionalData = fields }, cancellationToken: ct);
        return true;
    }

    public async Task<bool> DeleteMatchListAsync(string name, CancellationToken ct = default)
    {
        var siteId = await GetSiteIdAsync();
        var listId = await ResolveListIdAsync(MailMatchListsList);
        var spId   = await FindMatchListSpIdAsync(name, ct);
        if (spId is null) return false;
        await GetGraph().Sites[siteId].Lists[listId].Items[spId].DeleteAsync(cancellationToken: ct);
        return true;
    }

    /// <summary>Title is the list-name column but isn't indexed (and the match-list set is tiny), so
    /// match in-memory rather than a $filter — which Graph rejects on an unindexed column.</summary>
    private async Task<string?> FindMatchListSpIdAsync(string name, CancellationToken ct)
    {
        var rows = await ReadAllListItemsAsync(MailMatchListsList, ["Title"], null, ct);
        var row  = rows.FirstOrDefault(f => string.Equals(GetStr(f, "Title"), name, StringComparison.OrdinalIgnoreCase));
        return row is null ? null : GetStr(row, "__spId");
    }
}
