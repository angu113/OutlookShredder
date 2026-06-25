namespace OutlookShredder.Proxy.Models;

public class ConversationMessage
{
    public string? SpItemId { get; set; }
    public string RfqId { get; set; } = "";
    public string SupplierName { get; set; } = "";
    public string? SupplierResponseId { get; set; }
    public string Direction { get; set; } = "out";
    public string? MessageId { get; set; }
    public string? InReplyTo { get; set; }
    public DateTimeOffset SentAt { get; set; }
    public string? Subject { get; set; }
    public string? BodyText { get; set; }
    public bool HasAttachments { get; set; }
    /// <summary>Filename of the attachment (e.g. "quote.pdf") for inbound SR messages.</summary>
    public string? AttachmentName { get; set; }
    public bool ExtractedPricing { get; set; }
    /// <summary>
    /// For inbound messages that arrived as forwards: the original supplier's email address
    /// extracted from the forwarded body. Null for direct (non-forwarded) emails.
    /// Use this instead of EmailFrom when composing a reply.
    /// </summary>
    public string? ContactEmail { get; set; }
    /// <summary>Graph mail conversationId — groups all emails in the same reply chain.</summary>
    public string? GraphConversationId { get; set; }
    /// <summary>The SliVersion active on the SR at the time this message was sent/received.</summary>
    public int? SliVersionAtSend { get; set; }
    /// <summary>Comma-separated BCC addresses used on this outbound message.</summary>
    public string? BccAddresses { get; set; }

    /// <summary>
    /// AI-generated terse (≤2-line) summary of an INBOUND email, shown in the RFQ Mailbox in place of the
    /// raw body preview. "Regret" when the supplier declined. Null for outbound / unsummarized rows
    /// (the client then falls back to the body preview). Generated once at write time and persisted.
    /// </summary>
    public string? Summary { get; set; }

    // ── Per-message read state (team-wide; inbound messages only) ────────────
    /// <summary>True when this inbound supplier message has been marked read by someone. Outbound = always false.</summary>
    public bool IsRead { get; set; }
    /// <summary>When it was marked read (UTC ISO), or null.</summary>
    public string? ReadAt { get; set; }
    /// <summary>The Shredder username who marked it read, or null (e.g. historical/backfilled rows).</summary>
    public string? ReadBy { get; set; }
}

/// <summary>
/// Summary entry for one MSG conversation (RFQ_ID starting with "MSG"),
/// returned by GET /api/supplier-conversations/msg-list.
/// </summary>
public record MsgConversationSummary(string RfqId, string SupplierName, string? Subject, DateTime LastAt);

public class SupplierInquiryRequest
{
    public string        To                      { get; set; } = "";
    public string?       Bcc                     { get; set; }         // legacy single BCC
    public List<string>? BccAddresses            { get; set; }         // multi-BCC (preferred)
    public string        Subject                 { get; set; } = "";
    public string  Body                    { get; set; } = "";
    public string  RfqId                   { get; set; } = "";
    public string  SupplierName            { get; set; } = "";
    public string? SupplierResponseId      { get; set; }
    public string? InReplyTo               { get; set; }
    public string? AttachmentName          { get; set; }
    public string? AttachmentContentBase64 { get; set; }
    public string? AttachmentContentType   { get; set; }
}
