using System.Text.Json.Serialization;

namespace OutlookShredder.Proxy.Models;

/// <summary>
/// Second-pass extraction over a SUPPLIER Sales Order Confirmation / acknowledgement PDF
/// (ConfirmationExtractionService). The field the PO-confirmation matcher needs that lives in the PDF
/// rather than the email text: the supplier's promised ship / delivery (ETA) date, plus our PO# and
/// the supplier name for corroboration. Mirrors <see cref="BillExtraction"/> (same Claude→Gemini
/// document pattern) but oriented to inbound order confirmations rather than bills.
/// </summary>
public sealed class ConfirmationExtraction
{
    [JsonPropertyName("is_confirmation")] public bool    IsConfirmation { get; set; }
    /// <summary>Promised ship/delivery/due date as ISO yyyy-MM-dd, if explicitly stated. Null when only
    /// a lead time is given (we do not compute a date from "ships in N weeks").</summary>
    [JsonPropertyName("expected_date")]   public string? ExpectedDate   { get; set; }
    [JsonPropertyName("our_po_number")]   public string? OurPoNumber    { get; set; }
    [JsonPropertyName("supplier_name")]   public string? SupplierName   { get; set; }
}
