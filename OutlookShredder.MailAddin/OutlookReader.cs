using System;
using System.Collections.Generic;
using System.IO;
using OutlookShredder.MailAddin.Models;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace OutlookShredder.MailAddin;

internal static class OutlookReader
{
    private const string PropInternetMessageId = "http://schemas.microsoft.com/mapi/proptag/0x1035001E";
    private const string PropAttachDataBin     = "http://schemas.microsoft.com/mapi/proptag/0x37010102";

    public static MailMessagePayload BuildPayload(Outlook.MailItem mail, string? storeId = null, string? mailboxDisplayName = null)
    {
        var payload = new MailMessagePayload
        {
            EntryId            = mail.EntryID,
            StoreId            = storeId,
            Subject            = mail.Subject,
            FromAddress        = mail.SenderEmailAddress,
            FromName           = mail.SenderName,
            ToAddress          = mail.To,
            ReceivedAt         = mail.ReceivedTime.ToUniversalTime().ToString("o"),
            BodyText           = mail.Body,
            BodyHtml           = mail.HTMLBody,
            MailboxDisplayName = mailboxDisplayName
        };

        try
        {
            payload.InternetMessageId = mail.PropertyAccessor.GetProperty(PropInternetMessageId) as string;
        }
        catch { }

        foreach (Outlook.Attachment att in mail.Attachments)
        {
            if (att.Type != Outlook.OlAttachmentType.olByValue) continue;
            try
            {
                payload.Attachments.Add(ReadAttachment(att));
            }
            catch { }
        }

        return payload;
    }

    private static AttachmentPayload ReadAttachment(Outlook.Attachment att)
    {
        byte[] bytes;
        try
        {
            bytes = (byte[])att.PropertyAccessor.GetProperty(PropAttachDataBin);
        }
        catch
        {
            var tmp = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + Path.GetExtension(att.FileName));
            att.SaveAsFile(tmp);
            bytes = File.ReadAllBytes(tmp);
            File.Delete(tmp);
        }

        return new AttachmentPayload
        {
            FileName      = att.FileName ?? "attachment",
            ContentType   = GetMimeType(att.FileName),
            SizeBytes     = att.Size,
            ContentBase64 = Convert.ToBase64String(bytes)
        };
    }

    private static string GetMimeType(string? fileName) =>
        Path.GetExtension(fileName)?.ToLowerInvariant() switch
        {
            ".pdf"  => "application/pdf",
            ".docx" => "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            ".doc"  => "application/msword",
            ".xlsx" => "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            ".xls"  => "application/vnd.ms-excel",
            ".txt"  => "text/plain",
            ".csv"  => "text/csv",
            ".html" => "text/html",
            ".htm"  => "text/html",
            ".rtf"  => "application/rtf",
            ".png"  => "image/png",
            ".jpg"  => "image/jpeg",
            ".jpeg" => "image/jpeg",
            _       => "application/octet-stream"
        };
}
