namespace OutlookShredder.Proxy.Models;

public sealed class MessageRecord
{
    public int?    SpItemId       { get; set; }
    public string  From           { get; set; } = "";
    public string  To             { get; set; } = "";
    public string  Channel        { get; set; } = "internal";
    public string  Direction      { get; set; } = "out";
    public string  Body           { get; set; } = "";
    public string  ConversationId { get; set; } = "";
    public string  TimestampUtc   { get; set; } = "";
    public bool    IsRead         { get; set; }
    public string? ExternalId     { get; set; }
}

public sealed class ConversationSummary
{
    public string ConversationId  { get; set; } = "";
    public string Contact         { get; set; } = "";
    public string Channel         { get; set; } = "";
    public string LastMessageBody { get; set; } = "";
    public string LastTimestamp   { get; set; } = "";
    public int    UnreadCount     { get; set; }
}

public sealed class SendMessageRequest
{
    public string To      { get; set; } = "";
    public string Body    { get; set; } = "";
    public string Channel { get; set; } = "internal";
    public string From    { get; set; } = "";
}
