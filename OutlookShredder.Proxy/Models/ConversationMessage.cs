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
    public bool ExtractedPricing { get; set; }
    /// <summary>
    /// For inbound messages that arrived as forwards: the original supplier's email address
    /// extracted from the forwarded body. Null for direct (non-forwarded) emails.
    /// Use this instead of EmailFrom when composing a reply.
    /// </summary>
    public string? ContactEmail { get; set; }
}

public class SupplierInquiryRequest
{
    public string To { get; set; } = "";
    public string Subject { get; set; } = "";
    public string Body { get; set; } = "";
    public string RfqId { get; set; } = "";
    public string SupplierName { get; set; } = "";
    public string? SupplierResponseId { get; set; }
    public string? InReplyTo { get; set; }
    public string? AttachmentName { get; set; }
    public string? AttachmentContentBase64 { get; set; }
    public string? AttachmentContentType { get; set; }
}
