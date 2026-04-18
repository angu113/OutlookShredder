namespace OutlookShredder.Proxy.Models;

/// <summary>Structured data extracted by the AI from a Purchase Order PDF.</summary>
public class PoExtraction
{
    public string?         JobReference { get; set; }
    public string?         SupplierName { get; set; }
    public string?         PoNumber     { get; set; }
    public List<PoLineItem> LineItems   { get; set; } = [];
}

public class PoLineItem
{
    public string? Mspc     { get; set; }
    public string? Product  { get; set; }
    public double? Quantity { get; set; }
    public string? Size     { get; set; }
}
