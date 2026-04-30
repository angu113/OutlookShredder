using System.Text.RegularExpressions;

namespace OutlookShredder.Proxy.Services;

/// <summary>
/// Parses ERP document identity from a PDF filename.
///
/// Format: {Type}_{Organization}_{RecordType}{RecordNumber}[_ or . optional suffix].pdf
///
/// Type (case-insensitive):
///   Invoice | Order | SalesConfirmation | Sales | PaymentIn |
///   PickingSlip | ShippingSlip | GoodsShipment | Purchase_Order | PurchaseOrder
///
/// Organization:  020803  |  HSK
///
/// RecordType:    2–5 uppercase letters, e.g. SI, SO, ARR, GS, PO
/// RecordNumber:  exactly 7 digits
///
/// DocumentNumber (the unique ERP key stored in SharePoint):
///   "{Organization}-{RecordType}{RecordNumber}"
///   e.g. "020803-SI1234567", "HSK-SO9876543"
///
/// Special case — PurchaseOrder with no record info:
///   Filename may just be "PurchaseOrder.pdf" or "PurchaseOrder_HSK.pdf".
///   IsErp=true, HasDocNumber=false.  Caller must use AI to extract the document number.
/// </summary>
public static class ErpFilenameParser
{
    // Full match: Type_Org_RecordTypeRecordNumber[_suffix or .suffix]
    // The ERP system appends date/time stamps with either '_' or '.' before the file extension,
    // e.g. "Invoice_HSK_SI1023754.pdf" or "SalesConfirmation_HSK_SO1034957.20260428-205320.pdf".
    // Alternatives are ordered longest-first to avoid prefix ambiguity (SalesConfirmation > Sales,
    // Purchase_Order > PurchaseOrder, GoodsShipment > ShippingSlip > PickingSlip).
    private static readonly Regex _full = new(
        @"^(?<type>SalesConfirmation|Purchase_Order|GoodsShipment|ShippingSlip|PurchaseOrder|PickingSlip|PickList|PaymentIn|Invoice|Order|Sales)" +
        @"_(?<org>020803|HSK)" +
        @"_(?<rt>[A-Za-z]{1,5})(?<rn>\d{7})" +
        @"(?:[_.].*)?$",
        RegexOptions.IgnoreCase | RegexOptions.Compiled);

    // Fallback: PurchaseOrder (or Purchase_Order) recognised by type alone — record info absent.
    // Still an ERP file, but the document number must come from AI reading the PDF.
    private static readonly Regex _poFallback = new(
        @"^(?:Purchase_Order|PurchaseOrder)(?:[_.].*)?$",
        RegexOptions.IgnoreCase | RegexOptions.Compiled);

    /// <summary>
    /// Returns an <see cref="ErpFilenameResult"/> if the filename matches an ERP naming pattern,
    /// or <c>null</c> if the file should be skipped.
    /// </summary>
    // Matches Windows duplicate-file counter: " (1)", " (2)", " (10)", etc.
    private static readonly Regex _dupCounter =
        new(@"\s+\(\d+\)$", RegexOptions.Compiled);

    public static ErpFilenameResult? Parse(string fileName)
    {
        // Strip Windows duplicate counter before matching so "PickingSlip_HSK_PS1234567 (1).pdf"
        // resolves to the same document number as "PickingSlip_HSK_PS1234567.pdf".
        var stem = _dupCounter.Replace(Path.GetFileNameWithoutExtension(fileName), "");

        var m = _full.Match(stem);
        if (m.Success)
        {
            var org = m.Groups["org"].Value;
            var rt  = m.Groups["rt"].Value.ToUpperInvariant();
            var rn  = m.Groups["rn"].Value;

            return new ErpFilenameResult
            {
                IsErp          = true,
                HasDocNumber   = true,
                DocumentType   = MapDocumentType(m.Groups["type"].Value),
                DocumentNumber = $"{org}-{rt}{rn}",
            };
        }

        if (_poFallback.IsMatch(stem))
        {
            return new ErpFilenameResult
            {
                IsErp        = true,
                HasDocNumber = false,
                DocumentType = "PurchaseOrder",
            };
        }

        return null;
    }

    private static string MapDocumentType(string raw) =>
        raw.ToLowerInvariant() switch
        {
            "invoice"           => "Invoice",
            "order"             => "SalesOrder",
            "salesconfirmation" => "SalesOrder",
            "sales"             => "Quotation",
            "paymentin"         => "Payment",
            "pickingslip"       => "PickingSlip",
            "picklist"          => "PickingSlip",
            "shippingslip"      => "ShippingNote",
            "goodsshipment"     => "ShippingNote",
            "purchase_order"    => "PurchaseOrder",
            "purchaseorder"     => "PurchaseOrder",
            _                   => raw,
        };
}

public sealed class ErpFilenameResult
{
    /// <summary>True when the filename matches a known ERP document pattern.</summary>
    public bool    IsErp          { get; init; }

    /// <summary>
    /// True when the filename contains a full Organization + RecordType + RecordNumber,
    /// so <see cref="DocumentNumber"/> is reliable.
    /// False for PurchaseOrder files where the record number is absent — caller must derive
    /// it from the PDF content via AI.
    /// </summary>
    public bool    HasDocNumber   { get; init; }

    /// <summary>
    /// Mapped document type string: Invoice | SalesOrder | Quotation | Payment |
    /// PickingSlip | ShippingNote | PurchaseOrder.
    /// </summary>
    public string  DocumentType   { get; init; } = "";

    /// <summary>
    /// Unique ERP identifier: "{Organization}-{RecordType}{RecordNumber}",
    /// e.g. "020803-SI1234567".  Null when <see cref="HasDocNumber"/> is false.
    /// </summary>
    public string? DocumentNumber { get; init; }
}
