namespace OutlookShredder.Proxy.Models;

public class ProductLine
{
    public string? ProductName             { get; set; }
    public double? UnitsRequested          { get; set; }
    public double? UnitsQuoted             { get; set; }
    public double? LengthPerUnit           { get; set; }
    public string? LengthUnit              { get; set; }
    public double? WeightPerUnit           { get; set; }
    public string? WeightUnit              { get; set; }
    public double? PricePerPound           { get; set; }
    public double? PricePerFoot            { get; set; }
    public double? PricePerPiece           { get; set; }
    public double? TotalPrice              { get; set; }
    public string? LeadTimeText            { get; set; }
    public string? Certifications          { get; set; }
    public string? SupplierProductComments { get; set; }
}
