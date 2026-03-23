namespace OutlookShredder.Proxy.Models;

public class ProductLine
{
    public string? ProductName           { get; set; }
    public int?    UnitsRequested        { get; set; }
    public int?    UnitsQuoted           { get; set; }
    public double? LengthPerUnit         { get; set; }
    public string? LengthUnit            { get; set; }
    public double? WeightPerUnit         { get; set; }
    public string? WeightUnit            { get; set; }
    public double? PricePerPound         { get; set; }
    public double? PricePerFoot          { get; set; }
    public string? SupplierProductComments { get; set; }
}
