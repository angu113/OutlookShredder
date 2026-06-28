using System.Text.Json;

namespace OutlookShredder.Proxy.Models;

public sealed class MessageRecord
{
    private static readonly JsonSerializerOptions _mediaJson = new() { PropertyNameCaseInsensitive = true };

    public int?    SpItemId       { get; set; }
    public string  From           { get; set; } = "";
    public string  To             { get; set; } = "";
    public string  Channel        { get; set; } = "internal";
    public string  Direction      { get; set; } = "out";
    public string? Subject        { get; set; }
    public string  Body           { get; set; } = "";
    public string  ConversationId { get; set; } = "";
    public string  TimestampUtc   { get; set; } = "";
    public bool    IsRead         { get; set; }
    public string? ExternalId     { get; set; }
    /// <summary>CINQ inquiry this message threads into (SMS customer-inquiry pipeline). Null for plain
    /// internal/email/legacy SMS rows. Indexed on the Messages list for per-thread reads.</summary>
    public string? InquiryId      { get; set; }
    /// <summary>Delivery status for outbound SMS, updated by the SignalWire status callback (queued →
    /// sent → delivered / failed / undelivered). Null for inbound + non-SMS rows.</summary>
    public string? Status         { get; set; }
    /// <summary>Persisted JSON array of <see cref="MessageMedia"/> — the attachments (MMS / future email)
    /// stored durably under <c>InquiryMedia/{InquiryId}/</c>. Null/empty for text-only messages.</summary>
    public string? MediaJson      { get; set; }

    /// <summary>Typed view of <see cref="MediaJson"/> for the client (image/pdf/cad/file each carry a
    /// stored file name the client fetches via <c>GET /api/inquiries/{id}/media?name=</c>).</summary>
    public IReadOnlyList<MessageMedia> Media =>
        string.IsNullOrWhiteSpace(MediaJson)
            ? []
            : (JsonSerializer.Deserialize<List<MessageMedia>>(MediaJson, _mediaJson) ?? []);
}

/// <summary>One stored attachment on a message. <see cref="Kind"/> drives the client UI: image/pdf preview
/// inline; cad → "Save to CAD" (OneDrive); file → "Download" (Downloads folder, for inspection).</summary>
public sealed class MessageMedia
{
    public string Name        { get; set; } = "";   // stored file name under InquiryMedia/{inquiryId}/
    public string ContentType { get; set; } = "";
    public string Kind        { get; set; } = "file"; // image | pdf | cad | file
}

public sealed class ConversationSummary
{
    public string  ConversationId  { get; set; } = "";
    public string  Contact         { get; set; } = "";
    public string  Channel         { get; set; } = "";
    public string? Subject         { get; set; }
    public string  LastMessageBody { get; set; } = "";
    public string  LastTimestamp   { get; set; } = "";
    public int     UnreadCount     { get; set; }
}

public sealed class SendMessageRequest
{
    public string  To      { get; set; } = "";
    public string  Body    { get; set; } = "";
    public string  Channel { get; set; } = "internal";
    public string  From    { get; set; } = "";
    public string? Subject { get; set; }
}
