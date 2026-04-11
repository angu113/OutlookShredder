namespace OutlookShredder.Proxy.Models;

public class RfqExtraction
{
    public string?             JobReference           { get; set; }
    public string?             QuoteReference         { get; set; }
    public string?             SupplierName           { get; set; }
    public string?             FreightTerms           { get; set; }
    public List<ProductLine>   Products               { get; set; } = [];
}
