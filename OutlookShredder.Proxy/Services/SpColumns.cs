namespace OutlookShredder.Proxy.Services;

/// <summary>
/// SharePoint column-name hygiene. SharePoint reserves a set of internal field names on every list (Author,
/// Editor, Created, Modified, …). A custom column with one of those names silently SHADOWS the built-in
/// read-only system field, so writes to it 500 — this is exactly the InquiryNotes "Author" bug. Provisioning
/// checks names against this set so a collision is caught LOUD at construction, not at the first write.
/// </summary>
public static class SpColumns
{
    // Internal names SharePoint reserves on every list. A custom column must never use one of these.
    private static readonly HashSet<string> Reserved = new(StringComparer.OrdinalIgnoreCase)
    {
        "Author", "Editor", "Created", "Modified", "ID", "GUID", "ContentType", "ContentTypeId",
        "Attachments", "FileLeafRef", "FileRef", "FileDirRef", "Order", "Version", "_UIVersionString",
        "owshiddenversion", "WorkflowVersion", "ParentLeafName", "ParentVersionString",
        // "Title" is the built-in default column — usable, but never (re)create it as a custom column.
        "Title",
    };

    public static bool IsReserved(string name) => Reserved.Contains(name);
}
