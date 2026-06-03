namespace OutlookShredder.Proxy.Models;

/// <summary>Response body returned by POST /api/extract.</summary>
public class ExtractResponse
{
    /// <summary>True if at least one SharePoint row was written successfully.</summary>
    public bool              Success    { get; set; }

    /// <summary>The structured data the AI extracted from the email / attachment.</summary>
    public RfqExtraction?    Extracted  { get; set; }

    /// <summary>One entry per product line — parallel to <see cref="RfqExtraction.Products"/>.</summary>
    public List<SpWriteResult> Rows     { get; set; } = [];

    /// <summary>Error message when Success is false and no rows were written.</summary>
    public string?           Error      { get; set; }
}

/// <summary>Payload broadcast via SSE and Azure Service Bus when RFQ data changes.</summary>
public class RfqProcessedNotification
{
    /// <summary>"SR" = new/updated Supplier Response; "RFQ" = new/updated RFQ Reference; "PO" = purchase order received; "ERP" = ERP document processed.</summary>
    public string  EventType    { get; set; } = "SR";
    public string? SupplierName { get; set; }
    public string? RfqId        { get; set; }
    /// <summary>
    /// Graph message ID of the source email.  Used by Shredder as the dedup key so that
    /// SSE and Service Bus delivering the same event never double-toast, while two distinct
    /// emails from the same supplier still each produce their own toast.
    /// </summary>
    public string? MessageId    { get; set; }
    public List<RfqNotificationProduct> Products { get; set; } = [];
    /// <summary>Populated only when EventType = "Synonym".</summary>
    public SynonymGroup? SynonymGroup { get; set; }
    /// <summary>Populated only when EventType = "ERP".</summary>
    public ErpBusRecord? ErpDocument { get; set; }
    /// <summary>Populated only when EventType = "IncomingCall".</summary>
    public string? CallerName  { get; set; }
    /// <summary>Populated only when EventType = "IncomingCall". Normalized, e.g. "(973) 752-2193".</summary>
    public string? CallerPhone   { get; set; }
    /// <summary>Populated only when EventType = "IncomingCall". Business partner name from CRM lookup.</summary>
    public string? BpName        { get; set; }
    /// <summary>Populated only when EventType = "IncomingCall". BP popup message from CRM lookup.</summary>
    public string? PopupMessage  { get; set; }
    /// <summary>Populated only when EventType = "IncomingCall". Matched contact name from CRM lookup.</summary>
    public string? ContactName       { get; set; }
    /// <summary>Populated only when EventType = "IncomingCall". SharePoint item ID of the call log entry.</summary>
    public string? CallLogSpItemId   { get; set; }
    /// <summary>
    /// Identity of the proxy that published this event: "{MachineName}:{startupGuid}".
    /// Receiving proxies skip cache updates for their own events; Shredder logs it for diagnostics.
    /// </summary>
    public string? ProxyId { get; set; }
    /// <summary>
    /// Complete fresh SLI rows for the affected RFQ, read from SP immediately after write.
    /// When present, bus recipients merge these rows directly into their local copy without
    /// a proxy round-trip, eliminating the SP-index-lag window and cross-proxy stale-cache gap.
    /// Null when RfqId is unknown (orphan / WHOIS emails) or when the SP read failed.
    /// </summary>
    public List<Dictionary<string, object?>>? SliRows { get; set; }
    /// <summary>Populated when EventType = "WorkflowCard".</summary>
    public WorkflowCard? WorkflowCard     { get; set; }
    public string?       WorkflowAction   { get; set; }
    public int?          WorkflowDeletedId { get; set; }
    /// <summary>Populated when EventType = "Todo".</summary>
    public ShredderTodo? TodoItem         { get; set; }
    public string?       TodoAction       { get; set; }  // Created | Updated | Deleted
    public string?       TodoDeletedId    { get; set; }
    // Populated only when EventType = "Message"
    public string? MsgFrom           { get; set; }
    public string? MsgTo             { get; set; }
    public string? MsgBody           { get; set; }
    public string? MsgConversationId { get; set; }
    public string? MsgChannel        { get; set; }
    public string? MsgDirection      { get; set; }
    public string? MsgTimestamp      { get; set; }
    /// <summary>
    /// All CRM matches when the caller's phone number appears at more than one company.
    /// Null when there is 0 or 1 match (legacy BpName/ContactName/PopupMessage fields cover those cases).
    /// </summary>
    public List<CrmBusMatchDto>? CrmMatches { get; set; }
    /// <summary>Populated only when EventType = "Mail". Captured | Classified | Amended | Completed | Deleted.</summary>
    public string? MailAction { get; set; }
    /// <summary>Stable MailItemId of the affected mail-workbench item (EventType="Mail").</summary>
    public string? MailItemId { get; set; }
    /// <summary>
    /// Full item snapshot for cross-proxy cache sync (EventType="Mail"); null on "Deleted".
    /// Bus recipients apply this directly to their MailCache without an SP round-trip.
    /// </summary>
    public MailBusItem? MailItem { get; set; }
}

/// <summary>
/// Compact mail-workbench item carried on EventType="Mail" bus messages: the immutable item fields
/// merged with its current classification, sufficient to update a peer's MailCache and the Inbox grid
/// without a proxy/SP round-trip.
/// </summary>
public sealed class MailBusItem
{
    public string  MailItemId      { get; set; } = "";
    public string  SpId            { get; set; } = "";
    public string  WrapperGraphId  { get; set; } = "";
    public string  ConversationId  { get; set; } = "";
    public string  RefsJson        { get; set; } = "";
    public string  SourceType      { get; set; } = "email";
    public string  SourceMailbox   { get; set; } = "";
    public string  FromAddress     { get; set; } = "";
    public string  FromName        { get; set; } = "";
    public string  Subject         { get; set; } = "";
    public string  ReceivedAt      { get; set; } = "";
    public bool    HasAttachments  { get; set; }
    public string  AttachmentsJson { get; set; } = "";
    public bool    Completed       { get; set; }
    public string? CompletedAt     { get; set; }
    public string? CompletedBy     { get; set; }
    public bool    IsRead          { get; set; }
    public string? ReadAt          { get; set; }
    public string? ReadBy          { get; set; }
    public string? ClaimedBy       { get; set; }
    public string? ClaimedAt       { get; set; }
    // current classification
    public string  CategoryPath    { get; set; } = "Other";
    public string? OtherLabel      { get; set; }
    public double  Confidence      { get; set; }
    public string  KeywordTags     { get; set; } = "";
    public string? PoNumber        { get; set; }
    public string? SoNumber        { get; set; }
    public string? Amount          { get; set; }
}

/// <summary>One CRM company match carried on an IncomingCall bus event.</summary>
public sealed record CrmBusMatchDto(string BpName, string? ContactName, string? PopupMessage);

/// <summary>
/// ERP document data carried on EventType="ERP" bus messages.
/// Mirrors ErpDocumentRecord but excludes the PDF attachment bytes.
/// IsNew=true means first write; IsNew=false means the record was updated (e.g. archived).
/// </summary>
public class ErpBusRecord
{
    public string? SpItemId          { get; set; }
    public string? DocumentNumber    { get; set; }
    public string? DocumentType      { get; set; }
    public string? DocumentDate      { get; set; }
    public string? CustomerName      { get; set; }
    public string? CustomerReference { get; set; }
    public string? TotalAmount       { get; set; }
    public string? Currency          { get; set; }
    public string? FileName          { get; set; }
    public string? PdfUrl            { get; set; }
    public string? ReceivedAt        { get; set; }
    public bool    IsArchived        { get; set; }
    public bool    IsNew             { get; set; }
    public string? SourceMachine     { get; set; }
    public string? SourceUser        { get; set; }
    /// <summary>JSON array of ErpAnnotation — carried on bus messages so the Focus view can display stamps without a round-trip.</summary>
    public string? UserAnnotations   { get; set; }
    public string? DeliveryMethod    { get; set; }
    public string? DeliveryAddress   { get; set; }
    /// <summary>JSON array of ErpLineItem — carried so Order Products can read line items on fresh bus arrivals without an SP re-fetch.</summary>
    public string? LineItemsJson     { get; set; }
}

public class RfqNotificationProduct
{
    public string? Name       { get; set; }
    public double? TotalPrice { get; set; }
    /// <summary>MSPC code — populated for PO events only.</summary>
    public string? Mspc       { get; set; }
    /// <summary>Size/dimensions — populated for PO events only.</summary>
    public string? Size       { get; set; }
}

/// <summary>Response for GET /api/rfq/changes — new supplier activity since a given timestamp.</summary>
public class ChangesResult
{
    public List<SupplierActivity> Activities { get; set; } = [];
    /// <summary>New RFQ References (outbound RFQ emails) created since the timestamp.</summary>
    public List<NewRfqActivity>   NewRfqs    { get; set; } = [];
    /// <summary>Server UTC time at the moment the query ran — use as the next 'since' value.</summary>
    public DateTime ServerTime { get; set; }
}

/// <summary>A new RFQ reference detected by the change poll (outbound RFQ sent by a requester).</summary>
public class NewRfqActivity
{
    public string             RfqId     { get; set; } = "";
    public string             Requester { get; set; } = "";
    public DateTime?          DateSent  { get; set; }
    public List<RfqLineItemSummary> LineItems { get; set; } = [];
}

public class RfqLineItemSummary
{
    public string? Mspc        { get; set; }
    public string? Product     { get; set; }
    public string? Units       { get; set; }
    public string? SizeOfUnits { get; set; }
}

public class SupplierActivity
{
    public string SupplierName { get; set; } = "";
    public string RfqId        { get; set; } = "";
    public List<ActivityProduct> Products { get; set; } = [];
}

public class ActivityProduct
{
    public string  Name       { get; set; } = "";
    public decimal TotalPrice { get; set; }
}

/// <summary>A purchase order record as stored in the PurchaseOrders SharePoint list.</summary>
public class PurchaseOrderRecord
{
    public string  SpItemId     { get; set; } = "";
    public string  RfqId        { get; set; } = "";
    public string  SupplierName { get; set; } = "";
    public string? PoNumber     { get; set; }
    public string? ReceivedAt   { get; set; }
    public string? MessageId    { get; set; }
    /// <summary>JSON array of { mspc, product, quantity, size } objects.</summary>
    public string  LineItems    { get; set; } = "[]";
    /// <summary>SharePoint web URL of the uploaded PO PDF attachment, if available.</summary>
    public string? PdfUrl       { get; set; }

    // ── Supplier-confirmation tracking (Fulfillment loop) ────────────────────
    /// <summary>"Pending" until a supplier confirmation is matched (manual or AI), then "Confirmed".</summary>
    public string? ConfirmStatus { get; set; }
    public string? ConfirmedAt   { get; set; }
    /// <summary>How the confirmation arrived: email | phone | payment | manual.</summary>
    public string? ConfirmedVia  { get; set; }
    /// <summary>ETA from the confirmation, if known.</summary>
    public string? ExpectedDate  { get; set; }
    public string? ConfirmNote   { get; set; }

    // ── Pay-to-release tracking (Fulfillment loop) ───────────────────────────
    /// <summary>None | Required | Paid. "Required" = a payment-to-release ball WE must clear.</summary>
    public string? PaymentStatus     { get; set; }
    public string? PaymentRequiredAt { get; set; }
    public string? PaidAt            { get; set; }
    public string? PaymentNote       { get; set; }

    // ── Matched bill (the payment-processor bill that set PaymentStatus=Required) ─────────────
    /// <summary>MailItemId of the bill email matched to this PO — lets the surface open the bill +
    /// its pay link when the user clicks a payment-due PO.</summary>
    public string? BillMailItemId { get; set; }
    /// <summary>Bill total + supplier reference captured at match time (audit / display).</summary>
    public string? BillAmount      { get; set; }
    public string? BillSupplierRef { get; set; }
    public string? BillMatchedAt   { get; set; }

    // ── Derived at-risk (computed on read by PoConfirmationMonitor; NOT stored in SP) ────────
    /// <summary>Supplier-acknowledgment risk: green | amber | red. Only "Pending" POs are scored.</summary>
    public string? AckLevel           { get; set; }
    /// <summary>Minutes since the PO was booked (ReceivedAt) - for "waiting N min" display.</summary>
    public int?    MinutesSincePlaced { get; set; }
    /// <summary>True if the supplier-ack EST cutoff has passed (won't be processed today).</summary>
    public bool    AckCutoffPassed    { get; set; }
    /// <summary>Pay-to-release risk: green | amber | red. Only PaymentStatus="Required" is scored.</summary>
    public string? PayLevel                    { get; set; }
    public int?    MinutesSincePaymentRequired { get; set; }
}

// ── RLI anchoring dry-run models ─────────────────────────────────────────────

/// <summary>Compact SLI row for RLI anchoring comparison.</summary>
public class SliCompactRow
{
    public string? SupplierName        { get; set; }
    public string? ProductName         { get; set; }
    public string? SupplierProductName { get; set; }
    public string? ProductSearchKey    { get; set; }
    public string? CatalogProductName  { get; set; }
}

/// <summary>SR email row for RLI anchoring AI dry-run.</summary>
public class SrEmailRow
{
    public string? SupplierName  { get; set; }
    public string? EmailBody     { get; set; }
    public string? EmailFrom     { get; set; }
    public string? EmailSubject  { get; set; }
    public string? MessageId     { get; set; }
}

/// <summary>One entry in the phone call log stored in SharePoint.</summary>
public class PhoneCallLogRecord
{
    public string  SpItemId     { get; set; } = "";
    public string  CallerName   { get; set; } = "";
    public string? CallerPhone  { get; set; }
    public string? BpName       { get; set; }
    public string? ContactName  { get; set; }
    public string? PopupMessage { get; set; }
    public string? ReceivedAt   { get; set; }
    public string? Notes        { get; set; }
}

/// <summary>SharePoint write outcome for a single extracted product line.</summary>
public class SpWriteResult
{
    /// <summary>Product name as extracted by the AI (for display/logging).</summary>
    public string? ProductName { get; set; }

    /// <summary>True if the SharePoint list item was created successfully.</summary>
    public bool    Success     { get; set; }

    /// <summary>SharePoint item ID of the newly created list item (SupplierResponse row).</summary>
    public string? SpItemId    { get; set; }

    /// <summary>SharePoint item ID of the SupplierLineItem row written for this product.</summary>
    public string? SliSpItemId { get; set; }

    /// <summary>Web URL to the SharePoint list item (used by the task pane link).</summary>
    public string? SpWebUrl    { get; set; }

    /// <summary>True if an existing row was updated; false if a new row was inserted.</summary>
    public bool    Updated         { get; set; }

    /// <summary>Resolved canonical supplier name (populated after SP write).</summary>
    public string? SupplierName     { get; set; }

    /// <summary>Resolved RFQ ID from the matched SupplierResponse row (may differ from req.JobRefs when the email subject has no bracket reference).</summary>
    public string? RfqId            { get; set; }

    /// <summary>True when the supplier name could not be matched to the reference list — no row was written.</summary>
    public bool    SupplierUnknown { get; set; }

    /// <summary>Error message if Success is false.</summary>
    public string? Error           { get; set; }
}
