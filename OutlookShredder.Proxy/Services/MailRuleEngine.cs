using System.Text.RegularExpressions;
using OutlookShredder.Proxy.Models;

namespace OutlookShredder.Proxy.Services;

/// <summary>
/// Pure deterministic rule evaluator for mail classification. Given an email's signals, returns the
/// first enabled rule (ascending Priority) whose conditions ALL match — that rule's CategoryPath is
/// then filed at full confidence with the AI classifier skipped. No DI, no I/O: fully unit-testable.
/// </summary>
public static class MailRuleEngine
{
    // User-authored regex runs against untrusted email text; cap it so a pathological pattern can't
    // hang the poller (ReDoS). An invalid or slow pattern simply doesn't match.
    private static readonly TimeSpan RegexTimeout = TimeSpan.FromMilliseconds(100);

    /// <summary>The first matching rule, or null when none match (caller falls through to the AI).
    /// When <paramref name="lists"/> is supplied, each condition's ValueListRef is expanded against it
    /// (the named match list's values are merged in) before matching.</summary>
    public static MailRule? FirstMatch(IEnumerable<MailRule> rules, MailRuleSignals signals,
        IReadOnlyDictionary<string, List<string>>? lists = null)
    {
        foreach (var rule in rules.Where(r => r.Enabled).OrderBy(r => r.Priority))
        {
            var effective = lists is null ? rule : Resolve(rule, lists);
            if (Matches(effective, signals)) return rule;   // return the ORIGINAL rule (name/HitCount)
        }
        return null;
    }

    /// <summary>Pure expansion: a copy of the rule with each condition's ValueListRef resolved (the named
    /// list's values merged into Values), so the matcher always works on concrete values.</summary>
    public static MailRule Resolve(MailRule r, IReadOnlyDictionary<string, List<string>> lists)
    {
        if (r.Conditions.All(c => string.IsNullOrEmpty(c.ValueListRef))) return r;
        return new MailRule
        {
            Id = r.Id, Name = r.Name, Enabled = r.Enabled, Priority = r.Priority,
            CategoryPath = r.CategoryPath, HitCount = r.HitCount,
            Conditions = r.Conditions.Select(c =>
                string.IsNullOrEmpty(c.ValueListRef) || !lists.TryGetValue(c.ValueListRef!, out var extra) || extra.Count == 0
                    ? c
                    : new MailRuleCondition
                    {
                        Signal = c.Signal, Operator = c.Operator, MinMatches = c.MinMatches, ValueListRef = c.ValueListRef,
                        Values = c.Values.Concat(extra).Distinct(StringComparer.OrdinalIgnoreCase).ToList(),
                    }).ToList(),
        };
    }

    /// <summary>True when EVERY condition of the rule matches the signals (AND). An empty rule never
    /// matches — guards against an unconfigured rule silently becoming a catch-all.</summary>
    public static bool Matches(MailRule rule, MailRuleSignals signals)
        => rule.Conditions.Count > 0 && rule.Conditions.All(c => Matches(c, signals));

    private static bool Matches(MailRuleCondition c, MailRuleSignals s)
    {
        var text = SignalText(c.Signal, s);
        return c.Operator switch
        {
            MailRuleOperator.Contains =>
                c.Values.Count > 0 && text.Contains(c.Values[0], StringComparison.OrdinalIgnoreCase),
            MailRuleOperator.NotContains =>
                c.Values.Count == 0 || !text.Contains(c.Values[0], StringComparison.OrdinalIgnoreCase),
            MailRuleOperator.Equals =>
                c.Values.Count > 0 && string.Equals(text.Trim(), c.Values[0].Trim(), StringComparison.OrdinalIgnoreCase),
            MailRuleOperator.Regex =>
                c.Values.Count > 0 && SafeRegex(text, c.Values[0]),
            MailRuleOperator.AnyOf =>
                c.Values.Count(v => text.Contains(v, StringComparison.OrdinalIgnoreCase)) >= Math.Max(1, c.MinMatches),
            _ => false,
        };
    }

    private static string SignalText(MailRuleSignal signal, MailRuleSignals s) => signal switch
    {
        MailRuleSignal.SenderAddress     => s.SenderAddress,
        MailRuleSignal.SenderDomain      => s.SenderDomain,
        MailRuleSignal.Subject           => s.Subject,
        MailRuleSignal.Body              => s.Body,
        MailRuleSignal.AttachmentName    => string.Join("\n", s.AttachmentNames),
        MailRuleSignal.AttachmentContent     => s.AttachmentContent,
        MailRuleSignal.SenderIsKnownSupplier => s.SenderIsKnownSupplier ? "true" : "false",
        _ => "",
    };

    private static bool SafeRegex(string text, string pattern)
    {
        try { return Regex.IsMatch(text, pattern, RegexOptions.IgnoreCase, RegexTimeout); }
        catch { return false; }   // invalid pattern or timeout → no match (never throw on user input)
    }
}
