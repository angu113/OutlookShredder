using OutlookShredder.Proxy.Models;

namespace OutlookShredder.Proxy.Services;

/// <summary>
/// In-memory cache of the deterministic MailRules (SP-backed), refreshed on a short TTL so a rule
/// edited in Tools takes effect on the next classification with no deploy. Evaluation delegates to the
/// pure <see cref="MailRuleEngine"/>. CRUD writes through to SP and updates the cache optimistically
/// (an immediate SP re-read can lag the write, mirroring the MailTaxonomyService pattern).
/// </summary>
public sealed class MailRuleService
{
    private readonly SharePointService _sp;
    private readonly ILogger<MailRuleService> _log;
    private readonly SemaphoreSlim _gate = new(1, 1);
    private static readonly TimeSpan Ttl = TimeSpan.FromMinutes(3);

    private volatile List<MailRule> _rules = new();
    private volatile Dictionary<string, List<string>> _lists = new(StringComparer.OrdinalIgnoreCase);
    private DateTimeOffset _loadedAt = DateTimeOffset.MinValue;

    public MailRuleService(SharePointService sp, ILogger<MailRuleService> log) { _sp = sp; _log = log; }

    private async Task EnsureFreshAsync(CancellationToken ct)
    {
        if (DateTimeOffset.UtcNow - _loadedAt < Ttl) return;
        await _gate.WaitAsync(ct);
        try
        {
            if (DateTimeOffset.UtcNow - _loadedAt < Ttl) return;
            _rules    = await _sp.ReadMailRulesAsync(ct);
            _lists    = (await _sp.ReadMatchListsAsync(ct)).ToDictionary(l => l.Name, l => l.Values, StringComparer.OrdinalIgnoreCase);
            _loadedAt = DateTimeOffset.UtcNow;
            _log.LogInformation("[MailRules] loaded {N} rule(s), {L} match list(s)", _rules.Count, _lists.Count);
        }
        catch (Exception ex) { _log.LogWarning(ex, "[MailRules] load failed — using last-good rules"); _loadedAt = DateTimeOffset.UtcNow; }
        finally { _gate.Release(); }
    }

    public async Task<List<MailRule>> GetRulesAsync(CancellationToken ct = default) { await EnsureFreshAsync(ct); return _rules; }

    /// <summary>First rule (by ascending priority) whose conditions all match, or null → fall through to
    /// the AI. ValueListRef conditions are expanded against the cached match lists.</summary>
    public async Task<MailRule?> FirstMatchAsync(MailRuleSignals signals, CancellationToken ct = default)
    {
        await EnsureFreshAsync(ct);
        return MailRuleEngine.FirstMatch(_rules, signals, _lists);
    }

    // ── Match lists (named, reusable value sets referenced by rule conditions) ──────────

    public async Task<List<MailMatchList>> GetListsAsync(CancellationToken ct = default)
    {
        await EnsureFreshAsync(ct);
        return _lists.Select(kv => new MailMatchList { Name = kv.Key, Values = kv.Value })
                     .OrderBy(l => l.Name, StringComparer.OrdinalIgnoreCase).ToList();
    }

    /// <summary>Upsert a match list by name. Returns true if it already existed.</summary>
    public async Task<bool> SaveListAsync(MailMatchList list, CancellationToken ct = default)
    {
        await EnsureFreshAsync(ct);
        var existed = await _sp.WriteMatchListAsync(list, ct);
        _lists = new Dictionary<string, List<string>>(_lists, StringComparer.OrdinalIgnoreCase) { [list.Name] = list.Values };
        return existed;
    }

    public async Task<bool> DeleteListAsync(string name, CancellationToken ct = default)
    {
        await EnsureFreshAsync(ct);
        var ok = await _sp.DeleteMatchListAsync(name, ct);
        if (ok) { var m = new Dictionary<string, List<string>>(_lists, StringComparer.OrdinalIgnoreCase); m.Remove(name); _lists = m; }
        return ok;
    }

    public async Task<string> AddAsync(MailRule rule, string by, CancellationToken ct = default)
    {
        await EnsureFreshAsync(ct);
        rule.Id = await _sp.WriteMailRuleAsync(rule, by, ct);
        _rules  = _rules.Concat([rule]).OrderBy(r => r.Priority).ToList();   // optimistic; reconciles on next TTL
        return rule.Id;
    }

    public async Task<bool> UpdateAsync(string ruleId, MailRule rule, CancellationToken ct = default)
    {
        await EnsureFreshAsync(ct);
        var ok = await _sp.UpdateMailRuleAsync(ruleId, rule, ct);
        if (ok)
        {
            rule.Id = ruleId;
            _rules  = _rules.Where(r => r.Id != ruleId).Concat([rule]).OrderBy(r => r.Priority).ToList();
        }
        return ok;
    }

    public async Task<bool> DeleteAsync(string ruleId, CancellationToken ct = default)
    {
        await EnsureFreshAsync(ct);
        var ok = await _sp.DeleteMailRuleAsync(ruleId, ct);
        if (ok) _rules = _rules.Where(r => r.Id != ruleId).ToList();
        return ok;
    }
}
