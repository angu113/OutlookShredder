using System.Collections.Generic;

namespace OutlookShredder.MailAddin.Models;

public class MailMessagePayload
{
    public string  EntryId             { get; set; } = string.Empty;
    public string? StoreId             { get; set; }
    public string? InternetMessageId   { get; set; }
    public string? Subject             { get; set; }
    public string? FromAddress         { get; set; }
    public string? FromName            { get; set; }
    public string? ToAddress           { get; set; }
    public string? ReceivedAt          { get; set; } // ISO 8601 UTC string
    public string? BodyText            { get; set; }
    public string? BodyHtml            { get; set; }
    public string? MailboxDisplayName  { get; set; }
    public List<AttachmentPayload> Attachments { get; set; } = new List<AttachmentPayload>();
}

public class AttachmentPayload
{
    public string  FileName        { get; set; } = string.Empty;
    public string? ContentType     { get; set; }
    public int     SizeBytes       { get; set; }
    public string  ContentBase64   { get; set; } = string.Empty;
}
