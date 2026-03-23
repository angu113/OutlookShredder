namespace OutlookShredder.Proxy.Models;

public class RfqExtraction
{
    public string?             JobReference           { get; set; }
    public string?             SupplierName           { get; set; }
    public string?             DateOfQuote            { get; set; }
    public string?             EstimatedDeliveryDate  { get; set; }
    public List<ProductLine>   Products               { get; set; } = [];
}
