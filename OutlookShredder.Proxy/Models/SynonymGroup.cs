namespace OutlookShredder.Proxy.Models;

/// <summary>
/// One synonym group: a canonical product-term and all supplier-specific variants
/// that map to it. Stored in the SharePoint ProductSynonyms list and injected
/// into AI system prompts so extraction handles supplier-specific terminology.
/// </summary>
public sealed class SynonymGroup
{
    /// <summary>SharePoint item ID — empty for groups not yet persisted.</summary>
    public string   SpItemId  { get; set; } = string.Empty;
    /// <summary>Canonical name used in our catalog (e.g. "#4 Brushed Finish").</summary>
    public string   Canonical { get; set; } = string.Empty;
    /// <summary>Loose grouping for the UI (e.g. "finish", "alloy", "condition").</summary>
    public string?  Category  { get; set; }
    /// <summary>All known supplier variants that mean the same thing as Canonical.</summary>
    public string[] Variants  { get; set; } = [];
}
