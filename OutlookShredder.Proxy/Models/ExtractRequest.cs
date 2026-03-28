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

    /// <summary>Base64-encoded attachment bytes (PDF/DOCX sent natively to Claude)</summary>
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
}
