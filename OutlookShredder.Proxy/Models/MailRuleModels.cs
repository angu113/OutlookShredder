using System.Text.Json.Serialization;

namespace OutlookShredder.Proxy.Models;

/// <summary>The signal a rule condition inspects on an email.</summary>
[JsonConverter(typeof(JsonStringEnumConverter))]
public enum MailRuleSignal
{
    SenderAddress,     // full from address, e.g. billing@acme.com
    SenderDomain,      // domain part only, e.g. acme.com
    Subject,
    Body,
    AttachmentName,    // any attachment filename
    AttachmentContent, // extracted/OCR'd text of the attachments (populated in Phase 2)
}

/// <summary>How a condition's value(s) are tested against the signal text.</summary>
[JsonConverter(typeof(JsonStringEnumConverter))]
public enum MailRuleOperator
{
    Contains,     // signal contains Values[0] (case-insensitive)
    NotContains,  // signal does NOT contain Values[0]
    Equals,       // signal equals Values[0] exactly (trimmed, case-insensitive)
    Regex,        // Values[0] is a regex matched against the signal
    AnyOf,        // signal contains at least MinMatches of Values (the MTR-indicator pattern)
}

/// <summary>One condition in a rule. A rule matches when ALL of its conditions match (AND).</summary>
public sealed class MailRuleCondition
{
    public MailRuleSignal Signal { get; set; }
    public MailRuleOperator Operator { get; set; }
    /// <summary>Value(s) to test. Contains/NotContains/Equals/Regex use the first; AnyOf uses all.</summary>
    public List<string> Values { get; set; } = [];
    /// <summary>For AnyOf: how many of Values must be present for the condition to match (min 1).</summary>
    public int MinMatches { get; set; } = 1;
}

/// <summary>
/// A deterministic classification rule. When every condition matches, the email is filed to
/// CategoryPath at full confidence and the AI classifier is skipped. Rules are evaluated in
/// ascending Priority order; the first matching rule wins. Managed back-office in Tools and
/// persisted to the MailRules SP list. Evaluation lives in the pure, unit-tested MailRuleEngine.
/// </summary>
public sealed class MailRule
{
    public string Id { get; set; } = "";
    public string Name { get; set; } = "";
    public bool Enabled { get; set; } = true;
    /// <summary>Lower = evaluated first.</summary>
    public int Priority { get; set; }
    public List<MailRuleCondition> Conditions { get; set; } = [];
    /// <summary>Target taxonomy bucket, e.g. "Supplier/MTRs".</summary>
    public string CategoryPath { get; set; } = "";
    /// <summary>Running count of how many emails this rule has filed (for cleanup/diagnostics).</summary>
    public int HitCount { get; set; }
}

/// <summary>
/// The signals extracted from one email for rule evaluation. AttachmentContent is the concatenated
/// extracted/OCR'd attachment text — empty until Phase 2 wires attachment-content extraction.
/// </summary>
public sealed class MailRuleSignals
{
    public string SenderAddress { get; set; } = "";
    public string SenderDomain { get; set; } = "";
    public string Subject { get; set; } = "";
    public string Body { get; set; } = "";
    public List<string> AttachmentNames { get; set; } = [];
    public string AttachmentContent { get; set; } = "";
}
