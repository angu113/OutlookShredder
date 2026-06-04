using System.Text;
using System.Text.Json.Serialization;

namespace OutlookShredder.Proxy.Models;

/// <summary>
/// Fixed classification taxonomy for the mail workbench (see wip/mail-classification.md).
/// Closed set of leaves + an "Other" escape hatch where the AI proposes a free sub-label.
/// The leaf descriptions are the verbatim guidance fed to the classifier prompt.
/// </summary>
public static class MailTaxonomy
{
    public sealed record Leaf(string Top, string Sub, string Description)
    {
        /// <summary>Canonical path, e.g. "Supplier/MTRs". "Other" has no sub.</summary>
        public string Path => string.IsNullOrEmpty(Sub) ? Top : $"{Top}/{Sub}";
    }

    public static readonly IReadOnlyList<Leaf> Leaves =
    [
        new("Supplier", "RFQ Responses",
            "Pricing/quote responses NOT sent to store@mithrilmetals.com (those go through the existing RFQ pipeline); supplier pricing and quote-related mail."),
        new("Supplier", "Order Confirmations",
            "Supplier Sales Order confirmations acknowledging a PO we placed; usually contain our PO number HSK-POxxxxxxx."),
        new("Supplier", "MTRs",
            "Mill Certificates, Material Test Reports, Certificates of Compliance, and other standards/cert emails suppliers send us."),
        new("Supplier", "Invoices and Bills",
            "An UNPAID request to pay money we still OWE a supplier or service provider: an invoice or bill with an amount DUE (after we receive goods/metal, or for services like power, water, security), typically carrying 'amount due', 'please remit', or a 'pay now' link. NOT a confirmation that a payment already happened — a receipt/authorization is Supplier/Receipts."),
        new("Supplier", "Statements",
            "Statements of account from suppliers and providers."),
        new("Supplier", "Receipts",
            "Confirmation that a payment we made to a supplier or service provider has ALREADY been made or charged: payment receipts, credit-card authorizations/approvals, 'Purchase Receipt', 'transaction approved/charged', 'payment received', 'thank you for your payment' — even when it shows an amount, and even when sent via a billing processor (QuickBooks/Intuit, Enmark, SlimCD, etc.; identify the real supplier from the body). The test is do we still OWE it (Invoices and Bills) or has it already been PAID (Receipts). Distinct from Corporate/Receipts (our own corporate/admin purchases and subscriptions)."),

        new("Customer", "Inquiries",
            "General requests for metal or fabrication."),
        new("Customer", "Website Inquiries",
            "Requests for material or fabrication sent to us from the franchise website."),
        new("Customer", "Orders",
            "Follow-ups on orders in progress; usually carry the sales order reference HSK-SOxxxxxxx in the subject; includes customer POs sent to us."),
        new("Customer", "Web Orders",
            "Orders booked from the franchise website, notified by email with an attached invoice/order to process."),
        new("Customer", "Statements",
            "A customer (or someone acting for one) requesting their open-invoice / account statement from us. Distinct from Supplier/Statements, which is a statement a supplier sends to us."),

        new("Corporate", "Tax",
            "Items specifically related to tax — sales tax, corporation tax."),
        new("Corporate", "Accounting",
            "Messages from our accounting partner."),
        new("Corporate", "Receipts",
            "Email receipts (and placeholders for receipt PDFs/scans)."),
        new("Corporate", "Payroll",
            "Payroll-related messages and reminders."),
        new("Corporate", "Training",
            "Training reports, course reminders, and follow-ups — e.g. Metal Supermarkets University (MSU) / smarteru.com supervisor and completion reports."),

        new("Franchise", "Newsletters",
            "Newsletters, bulletins, and corporate/franchisor marketing — e.g. \"The Pipeline\", \"Fundamentals That Matter to Every Fabricator\", and mail from emarketing@ / communications@Metalsupermarkets.com."),
        new("Franchise", "Inquiries",
            "Prospective franchisees contacting us for information about owning a Metal Supermarkets franchise."),

        new("Other", "Login Codes",
            "MFA / 2FA / one-time verification or login codes from any service."),
        new("Other", "Junk",
            "Unsolicited marketing, cold sales solicitations, and other low-value mail that can sit here until reviewed."),
        new("Other", "",
            "Everything else that does not fit the categories above. Propose a concise free-text sub-label naming the emergent category."),
    ];

    public static readonly IReadOnlySet<string> ValidPaths =
        Leaves.Select(l => l.Path).ToHashSet(StringComparer.OrdinalIgnoreCase);

    /// <summary>Snap an AI-returned category to a known leaf path; unknown values fall back to "Other".</summary>
    public static string Coerce(string? category)
    {
        if (string.IsNullOrWhiteSpace(category)) return "Other";
        var c = category.Trim();
        var match = Leaves.FirstOrDefault(l => string.Equals(l.Path, c, StringComparison.OrdinalIgnoreCase));
        return match?.Path ?? "Other";
    }

    /// <summary>Renders the taxonomy block for the classifier system prompt.</summary>
    public static string RenderForPrompt()
    {
        var sb = new StringBuilder();
        foreach (var top in Leaves.Select(l => l.Top).Distinct())
        {
            sb.Append("- ").Append(top).Append('\n');
            foreach (var leaf in Leaves.Where(l => l.Top == top))
            {
                if (string.IsNullOrEmpty(leaf.Sub))
                    sb.Append("    (").Append(top).Append("): ").Append(leaf.Description).Append('\n');
                else
                    sb.Append("    \"").Append(leaf.Path).Append("\": ").Append(leaf.Description).Append('\n');
            }
        }
        return sb.ToString();
    }
}

/// <summary>Text inputs to the classifier (email or file-derived).</summary>
public sealed class MailClassifyInput
{
    public string Subject         { get; set; } = "";
    public string FromAddress     { get; set; } = "";
    public string FromName        { get; set; } = "";
    public string ToLine          { get; set; } = "";
    public string BodyText        { get; set; } = "";
    public List<string> AttachmentNames { get; set; } = [];
    /// <summary>Category a prior message in the same conversation got — biases this one for thread consistency.</summary>
    public string? ThreadCategoryHint { get; set; }
}

/// <summary>
/// Result of one classification run. Mirrors what gets persisted to MailClassifications
/// (Phase 1b); for now returned by the preview endpoint for quality validation.
/// </summary>
public sealed class MailClassificationResult
{
    /// <summary>Canonical taxonomy path, coerced to a known leaf ("Other" if the model strayed).</summary>
    public string Category        { get; set; } = "Other";
    /// <summary>AI's proposed free-text sub-label when Category == "Other".</summary>
    public string? OtherLabel     { get; set; }
    public double Confidence      { get; set; }
    public List<string> Keywords  { get; set; } = [];
    /// <summary>The actual supplier/vendor the email is from or about — for payment-processor senders the
    /// real supplier read from the body, not the sender domain. Drives the supplier grouping.</summary>
    public string? SupplierName   { get; set; }
    public string? PoNumber       { get; set; }
    public string? SoNumber       { get; set; }
    public string? Amount         { get; set; }
    /// <summary>The SUPPLIER's own reference printed on a bill/invoice/receipt - their invoice #,
    /// sales-order, or quote reference (consistent with the supplier quote ref we captured). The key
    /// the bill -> PO matcher compares to our `QuoteReference`. Distinct from our PoNumber/SoNumber.</summary>
    public string? SupplierReference { get; set; }
    /// <summary>A "pay now" URL extracted from the body (payment-processor bills). Lets the surface
    /// deep-link to payment. Deterministic regex, not AI.</summary>
    public string? PayLink        { get; set; }
    /// <summary>Supplier's promised ship/delivery (ETA) date from a sales order confirmation, as ISO
    /// yyyy-MM-dd. Drives the PO ExpectedDate so the waiting card self-schedules out of Prioritize.
    /// Set for Supplier/Order Confirmations (text pass + the confirmation-PDF second pass).</summary>
    public string? ExpectedDate   { get; set; }
    public string? Reasoning      { get; set; }
    public string  AiProvider     { get; set; } = "";
    public string  AiModel        { get; set; } = "";
    /// <summary>Full raw provider JSON, kept for audit/analysis (RawAiResponse in SP).</summary>
    [JsonIgnore]
    public string  RawResponse    { get; set; } = "";
}
