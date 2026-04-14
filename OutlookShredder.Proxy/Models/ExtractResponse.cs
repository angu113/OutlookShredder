namespace OutlookShredder.Proxy.Models;

/// <summary>Response body returned by POST /api/extract.</summary>
public class ExtractResponse
{
    /// <summary>True if at least one SharePoint row was written successfully.</summary>
    public bool              Success    { get; set; }

    /// <summary>The structured data Claude extracted from the email / attachment.</summary>
    public RfqExtraction?    Extracted  { get; set; }

    /// <summary>One entry per product line — parallel to <see cref="RfqExtraction.Products"/>.</summary>
    public List<SpWriteResult> Rows     { get; set; } = [];

    /// <summary>Error message when Success is false and no rows were written.</summary>
    public string?           Error      { get; set; }
}

/// <summary>Payload broadcast via SSE and Azure Service Bus when RFQ data changes.</summary>
public class RfqProcessedNotification
{
    /// <summary>"SR" = new/updated Supplier Response; "RFQ" = new/updated RFQ Reference; "PO" = purchase order received.</summary>
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

/// <summary>SR email row for RLI anchoring Claude dry-run.</summary>
public class SrEmailRow
{
    public string? SupplierName  { get; set; }
    public string? EmailBody     { get; set; }
    public string? EmailFrom     { get; set; }
    public string? EmailSubject  { get; set; }
    public string? MessageId     { get; set; }
}

/// <summary>SharePoint write outcome for a single extracted product line.</summary>
public class SpWriteResult
{
    /// <summary>Product name as extracted by Claude (for display/logging).</summary>
    public string? ProductName { get; set; }

    /// <summary>True if the SharePoint list item was created successfully.</summary>
    public bool    Success     { get; set; }

    /// <summary>SharePoint item ID of the newly created list item.</summary>
    public string? SpItemId    { get; set; }

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
