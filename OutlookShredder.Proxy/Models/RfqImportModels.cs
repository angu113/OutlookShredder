namespace OutlookShredder.Proxy.Models;

/// <summary>Raw email data returned by GET /api/rfq-import/scan — Shredder parses line items locally.</summary>
public class RfqScanEmailDto
{
    public string   RfqId           { get; set; } = "";
    public string   Subject         { get; set; } = "";
    public DateTime SentAt          { get; set; }
    public string   Requester       { get; set; } = "";
    public string   EmailRecipients { get; set; } = "";
    public string   MailboxSource   { get; set; } = "";
    public string   BodyText        { get; set; } = "";
    public string   ContentType     { get; set; } = "text";
}

/// <summary>
/// POST /api/rfq/submit — single call from an Excel VBA macro to create (or upsert) an
/// RFQ Reference and its line items.  The proxy handles all SP writes so the macro needs
/// no SP credentials and no Graph API setup.
/// </summary>
public class RfqSubmitRequest
{
    /// <summary>RFQ identifier, e.g. "AB1234". Required.</summary>
    public string   RfqId           { get; set; } = "";
    public string?  Requester       { get; set; }
    /// <summary>ISO 8601 date/datetime string, e.g. "2026-04-06". Defaults to today if omitted.</summary>
    public string?  DateSent        { get; set; }
    public string?  EmailRecipients { get; set; }
    public List<RfqSubmitLineItem> LineItems { get; set; } = [];
}

public class RfqSubmitLineItem
{
    public string?  Mspc           { get; set; }
    public string?  Product        { get; set; }
    public double?  Units          { get; set; }
    public string?  SizeOfUnits    { get; set; }
    public string?  SupplierEmails { get; set; }
}

/// <summary>POST /api/rfq-import/reference</summary>
public class RfqReferenceRequest
{
    public string   RfqId           { get; set; } = "";
    public string   Requester       { get; set; } = "";
    public DateTime DateSent        { get; set; }
    public string   EmailRecipients { get; set; } = "";
}

/// <summary>One entry in POST /api/rfq-import/line-items array.</summary>
public class RfqLineItemRequest
{
    public string  RfqId          { get; set; } = "";
    public string? Mspc           { get; set; }
    public string? Product        { get; set; }
    public double? Units          { get; set; }
    public string? SizeOfUnits    { get; set; }
    public string? SupplierEmails { get; set; }
}

/// <summary>One entry returned by GET /api/mail/processed-emails.</summary>
public class ProcessedEmailDto
{
    public string   MessageId  { get; set; } = "";
    public string   Subject    { get; set; } = "";
    public string   From       { get; set; } = "";
    public DateTime ReceivedAt { get; set; }
    public string   Preview    { get; set; } = "";
    public bool     IsUnknown  { get; set; }
}

/// <summary>Body for POST /api/mail/reprocess-selected.</summary>
public class ReprocessRequest
{
    public List<string> MessageIds { get; set; } = [];
}
