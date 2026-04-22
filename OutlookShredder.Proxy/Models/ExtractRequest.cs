namespace OutlookShredder.Proxy.Models;

/// <summary>
/// Request body sent from the Office.js add-in to POST /api/extract
/// </summary>
public class ExtractRequest
{
    /// <summary>Plain-text email body OR decoded attachment text</summary>
    public string  Content     { get; set; } = string.Empty;

    /// <summary>"body" or "attachment"</summary>
    public string  SourceType  { get; set; } = "body";

    /// <summary>Filename when SourceType == "attachment"</summary>
    public string? FileName    { get; set; }

    /// <summary>Base64-encoded attachment bytes (PDF/DOCX sent natively to the AI)</summary>
    public string? Base64Data  { get; set; }

    /// <summary>MIME type of the attachment</summary>
    public string? ContentType { get; set; }

    /// <summary>Job references already detected by the add-in regex scan</summary>
    public List<string> JobRefs { get; set; } = [];

    /// <summary>Short email body snippet — context when processing an attachment</summary>
    public string? BodyContext { get; set; }

    /// <summary>Full email body text — stored in SharePoint alongside attachment data</summary>
    public string? EmailBody { get; set; }

    // Email metadata (written straight to SharePoint alongside extracted fields)
    public string? EmailSubject    { get; set; }
    public string? EmailFrom       { get; set; }
    public string? ReceivedAt      { get; set; }
    public bool    HasAttachment   { get; set; }

    // EWS fallback — populated by attachmentReader.js when getAttachmentContentAsync is unavailable
    // The proxy uses these to fetch the raw attachment via EWS SOAP if Base64Data is null
    public string? EwsToken  { get; set; }
    public string? EwsUrl    { get; set; }
    public string? ItemId    { get; set; }
    public string? AttachId  { get; set; }

    /// <summary>
    /// Supplier name pre-resolved by ShrConvInRouter (domain map or historical SR lookup).
    /// When set, WriteProductRowAsync uses this as the authoritative supplier name instead of
    /// re-resolving from the AI output or email domain, preventing WHOIS misroutes on
    /// follow-up emails where the AI might mis-identify the supplier from quoted reply text.
    /// </summary>
    public string? ResolvedSupplierName { get; set; }

    /// <summary>
    /// RFQ line items (what was originally requested) for the matched RFQ.
    /// Injected by MailPollerService before calling the AI so it can
    /// anchor each extracted supplier product to the nearest requested item.
    /// Empty when the RFQ ID is unknown or no RLI rows exist.
    /// </summary>
    public List<RliContextItem> RliItems { get; set; } = [];
}

/// <summary>One requested item from the RFQ Line Items list, used to anchor the AI's product matching.</summary>
public class RliContextItem
{
    /// <summary>Internal catalog code (MSPC / ProductSearchKey). Null when the RFQ was created without a catalog selection.</summary>
    public string? Mspc        { get; set; }
    /// <summary>Human-readable product name as written on the RFQ.</summary>
    public string? ProductName { get; set; }
}
