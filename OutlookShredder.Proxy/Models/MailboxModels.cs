namespace OutlookShredder.Proxy.Models;

/// <summary>
/// Mailbox-bridge models for the relay-envelope mail view (see wip/mailbox-bridge.md).
///
/// The bridge watches a mailbox we cannot authenticate into directly
/// (e.g. hackensack@metalsupermarkets.com). Inbound mail reaches us via a
/// server-side "Forward as attachment" Exchange rule that mirrors every message
/// into a folder on a mailbox we DO own (store@mithril → Inbox/Hackensack-Mirror).
/// Each mirrored message therefore carries the original supplier email as an
/// embedded message/rfc822 part; the bridge parses that part so the UI shows the
/// real sender/subject/body rather than "from hackensack, 1 attachment".
///
/// The HTTP/UI identity ("upn") is always the WATCHED mailbox. The bridge maps it
/// to the destination mailbox + mirror folder internally; the UI never sees that.
/// </summary>

/// <summary>One configured watched mailbox (from MailboxBridge:Mailboxes in secrets).</summary>
public sealed class MailboxConfig
{
    /// <summary>Watched mailbox identity shown to the user (e.g. hackensack@metalsupermarkets.com).</summary>
    public string WatchedUpn { get; set; } = "";

    /// <summary>Friendly name for the folder tree (e.g. "Hackensack").</summary>
    public string DisplayName { get; set; } = "";

    /// <summary>Mailbox we actually poll via Graph (e.g. store@mithrilmetals.com).</summary>
    public string DestinationUpn { get; set; } = "";

    /// <summary>Folder path inside the destination mailbox the mirror rule routes to (e.g. "Inbox/Hackensack-Mirror").</summary>
    public string DestinationFolderPath { get; set; } = "";

    /// <summary>Mailbox the outbound relay envelope is sent FROM (Phase 1.1; e.g. store@mithrilmetals.com).</summary>
    public string RelaySenderUpn { get; set; } = "";

    /// <summary>Subject prefix on the outbound relay envelope (Phase 1.1).</summary>
    public string EnvelopeSubjectPrefix { get; set; } = "[SHR-SEND]";
}

/// <summary>Live per-mailbox poll state (read by HealthController + GET /api/mailbox).</summary>
public sealed class MailboxStatus
{
    public string WatchedUpn   { get; set; } = "";
    public string DisplayName  { get; set; } = "";
    public bool   PollSucceeded { get; set; }
    public string? LastError   { get; set; }
    public DateTimeOffset? LastPollAt { get; set; }
    public int    MessageCount { get; set; }
    public int    UnreadCount  { get; set; }
}

/// <summary>A node in the (v1: single-level) folder tree presented to the UI.</summary>
public sealed class FolderNode
{
    public string Id          { get; set; } = "";
    public string DisplayName { get; set; } = "";
    public int    UnreadCount { get; set; }
    public int    TotalCount  { get; set; }
    public List<FolderNode> Children { get; set; } = [];
}

/// <summary>List-row view of a mirrored message — original sender/subject surfaced from the embedded mail.</summary>
public sealed class MailboxMessageHeader
{
    /// <summary>Wrapper message Graph id in the destination mailbox (stable key for detail/attachment fetch).</summary>
    public string Id             { get; set; } = "";
    public string Subject        { get; set; } = "";
    public string FromAddress    { get; set; } = "";
    public string FromName       { get; set; } = "";
    public string ReceivedAt     { get; set; } = "";  // ISO-8601 (when the mirror copy arrived)
    public bool   IsRead         { get; set; }
    public bool   HasAttachments { get; set; }
    public string Preview        { get; set; } = "";
}

/// <summary>Full message detail — original headers + plain-text body + attachment metadata.</summary>
public sealed class MailboxMessageBody
{
    public string Id          { get; set; } = "";
    /// <summary>The embedded original email's own Internet Message-ID — the global dedup key.</summary>
    public string InternetMessageId { get; set; } = "";
    public string Subject     { get; set; } = "";
    public string FromAddress { get; set; } = "";
    public string FromName    { get; set; } = "";
    public string ToLine      { get; set; } = "";
    public string CcLine      { get; set; } = "";
    /// <summary>The franchise mailbox this forward came through (the wrapper's From) — distinguishes
    /// hackensack@ vs awathen@ now that multiple mailboxes forward into the same mirror folder.</summary>
    public string SourceMailbox { get; set; } = "";
    /// <summary>Stable per-thread key from the original headers (Thread-Index conv prefix → References
    /// root → In-Reply-To → Message-ID) so replies in one conversation group together.</summary>
    public string ConversationId { get; set; } = "";
    public string ReceivedAt  { get; set; } = "";
    public bool   IsRead      { get; set; }
    public string BodyText    { get; set; } = "";
    public List<MailboxAttachmentMeta> Attachments { get; set; } = [];
}

/// <summary>Metadata for one attachment on the original (embedded) message.</summary>
public sealed class MailboxAttachmentMeta
{
    public string Name        { get; set; } = "";
    public string ContentType { get; set; } = "application/octet-stream";
    public long   Size        { get; set; }
}

/// <summary>Outbound compose request (Phase 1.1 — relay-envelope send). Read-side ignores this.</summary>
public sealed class MailboxSendRequest
{
    public List<string> To  { get; set; } = [];
    public List<string> Cc  { get; set; } = [];
    public List<string> Bcc { get; set; } = [];
    public string Subject   { get; set; } = "";
    public string Body      { get; set; } = "";
    public string? ReplyToMessageId { get; set; }
}
