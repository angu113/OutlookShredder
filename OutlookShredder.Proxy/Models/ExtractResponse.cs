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

/// <summary>Payload broadcast via SSE when a supplier quote is processed.</summary>
public class RfqProcessedNotification
{
    public string? SupplierName { get; set; }
    public string? RfqId       { get; set; }
    public List<RfqNotificationProduct> Products { get; set; } = [];
}

public class RfqNotificationProduct
{
    public string? Name       { get; set; }
    public double? TotalPrice { get; set; }
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

    /// <summary>True when the supplier name could not be matched to the reference list — no row was written.</summary>
    public bool    SupplierUnknown { get; set; }

    /// <summary>Error message if Success is false.</summary>
    public string? Error           { get; set; }
}
