namespace OutlookShredder.Proxy.Models;

/// <summary>Inquiry lifecycle states. Append-only flow Open → Quoted → Closed, plus the terminal Spam.</summary>
public static class InquiryStatus
{
    public const string Open   = "Open";
    public const string Quoted = "Quoted";
    public const string Closed = "Closed";
    public const string Spam   = "Spam";
}

/// <summary>
/// A lightweight customer inquiry (one per SMS thread) that can later be promoted to one or more
/// quotations (HSK#). Persisted to the SharePoint <c>Inquiries</c> list; <see cref="Id"/> doubles as the
/// SP item Title. Channel is an attribute of the underlying messages, not the inquiry — SMS today, email
/// later — so nothing here is SMS-specific.
/// </summary>
public sealed class Inquiry
{
    public int?   SpItemId      { get; set; }
    /// <summary>"CINQ-" + 5 Crockford Base32 chars. Also the SP Title.</summary>
    public string Id            { get; set; } = "";
    /// <summary>Customer phone in E.164-ish form ("+15551234567"). Indexed.</summary>
    public string CustomerPhone { get; set; } = "";
    public string Status        { get; set; } = InquiryStatus.Open;
    /// <summary>Shredder username this inquiry is assigned to (null = unassigned). First responder or claimer
    /// auto-takes it; any user can reassign/steal afterwards.</summary>
    public string? AssignedTo   { get; set; }
    /// <summary>Resolved CRM business-partner name (denormalised from CustomerCacheService at ingest for fast
    /// list display). Null = first-time / unknown caller.</summary>
    public string? CustomerName { get; set; }
    /// <summary>Resolved CRM contact name, if matched. Null = unknown.</summary>
    public string? ContactName  { get; set; }
    public string CreatedAt     { get; set; } = "";
    public string UpdatedAt     { get; set; } = "";
    /// <summary>UTC ISO timestamp of the most recent message in the thread. Indexed (latest-first).</summary>
    public string LastMessageAt { get; set; } = "";
    /// <summary>Inbound messages not yet read by an operator.</summary>
    public int    UnreadCount   { get; set; }
    /// <summary>True when the latest activity is an unanswered inbound (we owe the customer a reply). Set on
    /// inbound, cleared on our outbound — drives the "waiting on us for N" aging indicator.</summary>
    public bool   AwaitingReply { get; set; }
}

/// <summary>CRM context for the inquiry detail panel (v1: what LookupByPhone gives us).</summary>
public sealed record CustomerCard(string? BpName, string? ContactName, string? PopupMessage, bool IsFirstTime);

/// <summary>
/// A messaging contact keyed by phone (E.164). Tracks consent + carrier opt-out so outbound is suppressed
/// to opted-out numbers. Persisted to the SharePoint <c>MessagingContacts</c> list; <see cref="Phone"/> is
/// the SP Title.
/// </summary>
public sealed class MessagingContact
{
    public int?    SpItemId          { get; set; }
    public string  Phone             { get; set; } = "";
    public string? DisplayName       { get; set; }
    public string? ConsentCapturedAt { get; set; }
    /// <summary>How consent was captured, e.g. "inbound-sms".</summary>
    public string? ConsentMethod     { get; set; }
    public bool    OptOut            { get; set; }
    public string? OptOutAt          { get; set; }
}

/// <summary>Draft lifecycle: Pending (suggested, awaiting operator) → Used (sent) | Dismissed.</summary>
public static class DraftStatus
{
    public const string Pending   = "Pending";
    public const string Used      = "Used";
    public const string Dismissed = "Dismissed";
}

/// <summary>Where a draft came from: an AI suggestion or a fixed template.</summary>
public static class DraftSource
{
    public const string Ai       = "AI";
    public const string Template = "Template";
}

/// <summary>
/// A suggested outbound reply for an inquiry. AI-generated on every inbound (Phase 2) but NEVER auto-sent —
/// an operator explicitly accepts (→ Used) or dismisses it. Persisted to the SharePoint <c>Drafts</c> list.
/// </summary>
public sealed class InquiryDraft
{
    public int?    SpItemId            { get; set; }
    public string  InquiryId           { get; set; } = "";
    /// <summary>The inbound message (SID) this draft was generated in response to, if known.</summary>
    public string? TriggeringMessageId { get; set; }
    public string  Source              { get; set; } = DraftSource.Ai;
    public string? TemplateId          { get; set; }
    public string  Body                { get; set; } = "";
    /// <summary>AI-classified intent (e.g. Quote / OrderStatus / Question / Complaint / Other).</summary>
    public string? SuggestedIntent     { get; set; }
    /// <summary>AI-classified urgency (Low / Normal / High).</summary>
    public string? SuggestedUrgency    { get; set; }
    /// <summary>AI signal that the inquiry needs a quotation raised.</summary>
    public bool    NeedsQuote          { get; set; }
    /// <summary>JSON array of small discrete clarification options (e.g. materials) the operator can tap to
    /// append. Persisted on the Drafts list; null/empty for a plain reply.</summary>
    public string? OptionsJson         { get; set; }
    public string  Status              { get; set; } = DraftStatus.Pending;
    public string  CreatedAt           { get; set; } = "";
}

/// <summary>An append-only operator note on an inquiry (author + timestamp; no edit/delete in v1).</summary>
public sealed class InquiryNote
{
    public int?   SpItemId  { get; set; }
    public string InquiryId { get; set; } = "";
    public string Author    { get; set; } = "";
    public string Body      { get; set; } = "";
    public string CreatedAt { get; set; } = "";
}

/// <summary>A quotation (HSK#) linked to an inquiry — we store the reference only, not a duplicate quote.</summary>
public sealed class InquiryQuotation
{
    public int?   SpItemId  { get; set; }
    public string InquiryId { get; set; } = "";
    public string HskNumber { get; set; } = "";
    public string LinkedAt  { get; set; } = "";
    public string LinkedBy  { get; set; } = "";
}

/// <summary>Everything an operator needs to render one inquiry thread (GET /api/inquiries/{id}).</summary>
public sealed record InquiryDetail(
    Inquiry Inquiry,
    List<MessageRecord> Messages,
    List<InquiryNote> Notes,
    List<InquiryQuotation> Quotations,
    List<InquiryDraft> Drafts,
    MessagingContact? Contact,
    CustomerCard? Customer);

/// <summary>The structured result of one AI draft call: a reply plus classification fields.</summary>
public sealed class InquiryDraftResult
{
    public string  Reply      { get; set; } = "";
    public string  Intent     { get; set; } = "Other";
    public string  Urgency    { get; set; } = "Normal";
    public bool    NeedsQuote { get; set; }
    /// <summary>Small discrete clarification choices (e.g. Steel/Stainless/Aluminum) the operator can tap to
    /// append to the reply — the "option table" when we're close to a match. Empty otherwise.</summary>
    public List<string> Options { get; set; } = [];
    public string? AiModel    { get; set; }
}
