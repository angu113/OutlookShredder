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
    public string  RfqId            { get; set; } = "";
    /// <summary>Product Search Key — written to the SP "MSPC" column.</summary>
    public string? Mspc             { get; set; }
    public string? Product          { get; set; }
    public double? Units            { get; set; }
    public string? SizeOfUnits      { get; set; }
    public string? SupplierEmails   { get; set; }
    public string? ProductCategory  { get; set; }
    public string? ProductShape     { get; set; }
    public string? JobReference     { get; set; }
    public string? ProcessingSource { get; set; }
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

// ── RFQ New DTOs ─────────────────────────────────────────────────────────────

/// <summary>One row returned by GET /api/rfq-new/product-catalog.</summary>
public class ProductCatalogDto
{
    public string? Mspc             { get; set; }
    public string? ProductSearchKey { get; set; }
    public string? ProductName      { get; set; }
    public string? Category         { get; set; }
    public string? Shape            { get; set; }
}

/// <summary>One row returned by GET /api/rfq-new/supplier-relationships.</summary>
public class SupplierRelationshipDto
{
    public string  SupplierName { get; set; } = "";
    public string  Email        { get; set; } = "";
    public string  Metal        { get; set; } = "";
    public string? Shape        { get; set; }
}
