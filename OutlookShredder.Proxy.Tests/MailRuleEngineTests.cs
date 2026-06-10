using OutlookShredder.Proxy.Models;
using OutlookShredder.Proxy.Services;
using Xunit;

namespace OutlookShredder.Proxy.Tests;

// Contract tests for the deterministic rule engine: matching operators (incl. the AnyOf-threshold
// used for content-based MTR detection), AND-of-conditions, priority/first-match-wins, and the
// safety guards (disabled rules, empty rules, bad regex). Pure in/out assertions.
public class MailRuleEngineTests
{
    private static MailRuleCondition Cond(MailRuleSignal sig, MailRuleOperator op, int min = 1, params string[] vals)
        => new() { Signal = sig, Operator = op, Values = [.. vals], MinMatches = min };

    private static MailRule Rule(string category, int priority, bool enabled, params MailRuleCondition[] conds)
        => new() { Name = category, CategoryPath = category, Priority = priority, Enabled = enabled, Conditions = [.. conds] };

    private static MailRuleSignals Signals(
        string sender = "", string domain = "", string subject = "", string body = "",
        string attachContent = "", string[]? attachNames = null)
        => new()
        {
            SenderAddress = sender, SenderDomain = domain, Subject = subject, Body = body,
            AttachmentContent = attachContent, AttachmentNames = [.. (attachNames ?? [])],
        };

    [Fact]
    public void Contains_on_sender_domain_matches()
    {
        var rule = Rule("Supplier/Invoices and Bills", 0, true,
            Cond(MailRuleSignal.SenderDomain, MailRuleOperator.Contains, 1, "acme.com"));
        Assert.Equal("Supplier/Invoices and Bills",
            MailRuleEngine.FirstMatch([rule], Signals(domain: "acme.com"))?.CategoryPath);
    }

    [Fact]
    public void Regex_on_subject_matches_a_po_pattern()
    {
        var rule = Rule("Supplier/Order Confirmations", 0, true,
            Cond(MailRuleSignal.Subject, MailRuleOperator.Regex, 1, @"HSK-PO\d+"));
        Assert.NotNull(MailRuleEngine.FirstMatch([rule], Signals(subject: "Re: Purchase Order HSK-PO123456 confirmed")));
        Assert.Null(MailRuleEngine.FirstMatch([rule], Signals(subject: "just a normal subject")));
    }

    [Fact]
    public void AnyOf_threshold_detects_mtr_from_attachment_content()
    {
        // MTRs rarely say "MTR" — detect by N+ content indicators.
        var rule = Rule("Supplier/MTRs", 0, true,
            Cond(MailRuleSignal.AttachmentContent, MailRuleOperator.AnyOf, 2,
                "heat number", "certificate of compliance", "tensile strength", "yield strength", "ASTM A"));

        // 3 indicators present -> match
        Assert.NotNull(MailRuleEngine.FirstMatch([rule],
            Signals(attachContent: "Heat Number 9087. Tensile Strength 70ksi. Conforms to ASTM A36.")));
        // only 1 indicator -> below threshold, no match
        Assert.Null(MailRuleEngine.FirstMatch([rule],
            Signals(attachContent: "Yield strength data sheet, marketing only.")));
    }

    [Fact]
    public void All_conditions_must_match_AND()
    {
        var rule = Rule("Supplier/Statements", 0, true,
            Cond(MailRuleSignal.SenderDomain, MailRuleOperator.Contains, 1, "acme.com"),
            Cond(MailRuleSignal.Subject, MailRuleOperator.Contains, 1, "statement"));

        Assert.NotNull(MailRuleEngine.FirstMatch([rule], Signals(domain: "acme.com", subject: "August statement")));
        Assert.Null(MailRuleEngine.FirstMatch([rule], Signals(domain: "acme.com", subject: "an invoice")));  // 2nd cond fails
    }

    [Fact]
    public void Lowest_priority_rule_wins_when_several_match()
    {
        var general = Rule("Supplier/RFQ Responses", 10, true,
            Cond(MailRuleSignal.SenderDomain, MailRuleOperator.Contains, 1, "acme.com"));
        var specific = Rule("Supplier/Statements", 1, true,
            Cond(MailRuleSignal.SenderDomain, MailRuleOperator.Contains, 1, "acme.com"));
        var match = MailRuleEngine.FirstMatch([general, specific], Signals(domain: "acme.com"));
        Assert.Equal("Supplier/Statements", match?.CategoryPath);  // priority 1 beats 10
    }

    [Fact]
    public void Disabled_rules_are_skipped()
    {
        var rule = Rule("Supplier/MTRs", 0, enabled: false,
            Cond(MailRuleSignal.SenderDomain, MailRuleOperator.Contains, 1, "acme.com"));
        Assert.Null(MailRuleEngine.FirstMatch([rule], Signals(domain: "acme.com")));
    }

    [Fact]
    public void An_empty_rule_never_matches()
        => Assert.Null(MailRuleEngine.FirstMatch([Rule("Other", 0, true)], Signals(domain: "acme.com")));

    [Fact]
    public void NotContains_matches_when_value_absent()
    {
        var rule = Rule("Other/Junk", 0, true,
            Cond(MailRuleSignal.Body, MailRuleOperator.NotContains, 1, "unsubscribe"));
        Assert.NotNull(MailRuleEngine.FirstMatch([rule], Signals(body: "buy now")));
        Assert.Null(MailRuleEngine.FirstMatch([rule], Signals(body: "click unsubscribe here")));
    }

    [Fact]
    public void An_invalid_regex_does_not_throw_and_does_not_match()
    {
        var rule = Rule("Other", 0, true,
            Cond(MailRuleSignal.Subject, MailRuleOperator.Regex, 1, "([unclosed"));
        Assert.Null(MailRuleEngine.FirstMatch([rule], Signals(subject: "anything")));
    }

    [Fact]
    public void No_rules_returns_null_fall_through_to_ai()
        => Assert.Null(MailRuleEngine.FirstMatch([], Signals(domain: "acme.com")));
}
