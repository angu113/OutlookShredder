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
