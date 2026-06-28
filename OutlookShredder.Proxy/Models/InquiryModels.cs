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
    /// <summary>Shredder username this inquiry is assigned to (null = unassigned).</summary>
    public string? AssignedTo   { get; set; }
    public string CreatedAt     { get; set; } = "";
    public string UpdatedAt     { get; set; } = "";
    /// <summary>UTC ISO timestamp of the most recent message in the thread. Indexed (latest-first).</summary>
    public string LastMessageAt { get; set; } = "";
    /// <summary>Inbound messages not yet read by an operator.</summary>
    public int    UnreadCount   { get; set; }
}

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
