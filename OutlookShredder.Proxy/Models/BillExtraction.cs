using System.Text.Json.Serialization;

namespace OutlookShredder.Proxy.Models;

/// <summary>
/// Second-pass extraction over a SUPPLIER bill/invoice/receipt PDF (BillExtractionService). The fields
/// the bill -> PO matcher needs that live in the PDF rather than the email text: the bill total, the
/// supplier's OWN reference, and our PO# if printed. Distinct from <see cref="ErpExtraction"/>, which
/// is oriented to OUR outbound ERP documents (roles reversed: on a supplier bill, we are the buyer).
/// </summary>
public sealed class BillExtraction
{
    [JsonPropertyName("is_bill")]            public bool    IsBill            { get; set; }
    [JsonPropertyName("supplier_name")]      public string? SupplierName      { get; set; }
    [JsonPropertyName("amount")]             public string? Amount            { get; set; }
    [JsonPropertyName("supplier_reference")] public string? SupplierReference { get; set; }
    [JsonPropertyName("our_po_number")]      public string? OurPoNumber       { get; set; }
    [JsonPropertyName("currency")]           public string? Currency          { get; set; }
}
