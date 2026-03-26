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

    /// <summary>Error message if Success is false.</summary>
    public string? Error       { get; set; }
}
