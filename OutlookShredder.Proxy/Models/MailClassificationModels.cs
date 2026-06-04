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
            "A SUPPLIER's pricing / quote / availability response to a request for quote that did NOT come through our automated RFQ pipeline (the pipeline handles supplier replies sent to store@mithrilmetals.com carrying a [SHR:...] tracking token - those do NOT belong here). Capture two off-pipeline sources: (1) supplier quote replies FORWARDED in from the mailboxes we cannot monitor directly (a forwarding rule routes them to us); (2) ad-hoc RFQs a staff member ran by plain email instead of the formal RFQ system, and the supplier replies to them. Detect them by a supplier quoting prices, availability, or lead times on metal/material. We keep these so off-pipeline quotes are not lost and can later be pulled into the RFQ grid. This is a SUPPLIER quoting US - NOT a customer asking us to quote (that is Customer/Inquiries)."),
        new("Supplier", "Order Confirmations",
            "Supplier Sales Order confirmations acknowledging a PO we placed; usually contain our PO number HSK-POxxxxxxx."),
        new("Supplier", "MTRs",
            "Mill Certificates, Material Test Reports, Certificates of Compliance, and other standards/cert emails suppliers send us."),
        new("Supplier", "Invoices and Bills",
            "An UNPAID request to pay money we still OWE a supplier or service provider: an invoice or bill with an amount DUE (after we receive goods/metal, or for services like power, water, security), typically carrying 'amount due', 'please remit', or a 'pay now' link. NOT a confirmation that a payment already happened — a receipt/authorization is Supplier/Receipts."),
        new("Supplier", "Statements",
            "Statements of account from suppliers and providers."),
        new("Supplier", "Receipts",
            "Confirmation that a payment we made to a supplier or service provider has ALREADY been made or charged: payment receipts, credit-card authorizations/approvals, 'Purchase Receipt', 'transaction approved/charged', 'payment received', 'thank you for your payment' — even when it shows an amount, and even when sent via a billing processor (QuickBooks/Intuit, Enmark, SlimCD, etc.; identify the real supplier from the body). The test is do we still OWE it (Invoices and Bills) or has it already been PAID (Receipts). Supplier/Receipts cover payments to our SUPPLIERS for INVENTORY (the metal we resell) and SERVICES (e.g. powder coating, processing) - a separate accounting bucket from Corporate/Receipts, which is our own operating expenses (shop/office supplies, fuel, postage, subscriptions)."),

        new("Customer", "Inquiries",
            "A customer's request for material or fabrication/services that they emailed us DIRECTLY - the customer is the sender. Same intent as Customer/Website Inquiries, but it came straight to us rather than via a website form. Not a booked order (that is Customer/Web Orders or Customer/Orders)."),
        new("Customer", "Website Inquiries",
            "A request for material or fabrication submitted via a web form on the Franchisor website, notified to us by email. Same intent as a direct Customer/Inquiry, but because it arrives through the website the real customer CONTACT EMAIL is inside the message BODY, not the sender address (important when replying). Still an inquiry/quote request - NOT a booked order (that is Customer/Web Orders)."),
        new("Customer", "Orders",
            "Follow-ups on orders in progress, usually carrying OUR sales-order reference HSK-SOxxxxxxx; also a purchase order a CUSTOMER sends US to buy from us - the CUSTOMER's OWN PO number, NOT our HSK-PO (this includes another Metal Supermarkets franchise store ordering from us). NOT a supplier acknowledging our HSK-PO - that is Supplier/Order Confirmations."),
        new("Customer", "Web Orders",
            "An ACTUAL order placed through the Franchisor website (not just an inquiry), notified to us by email - typically with an attached invoice/order to process and often an EC###### reference. These are monitored and carry SLAs, so they are time-sensitive. Distinct from Customer/Website Inquiries, which is a request/quote with no order yet."),
        new("Customer", "Statements",
            "A customer (or someone acting for one) requesting their open-invoice / account statement from us. Distinct from Supplier/Statements, which is a statement a supplier sends to us."),

        new("Corporate", "Tax",
            "Items specifically related to tax — sales tax, corporation tax."),
        new("Corporate", "Accounting",
            "Correspondence to and from our outside accountant / bookkeeper (our accounting partner): financial statements they prepare for us, bookkeeping and reconciliation questions, and similar. Specifically our accountant's own mail - NOT franchise reports (Franchise/Reports), NOT payment receipts (Corporate/Receipts or Supplier/Receipts), and NOT tax-specific items (Corporate/Tax)."),
        new("Corporate", "Receipts",
            "Receipts/confirmations for our own OPERATING EXPENSES and overhead incurred running the business: shop and office supplies, fuel, postage, utilities, software subscriptions, and similar. OpEx we pay for ourselves - NOT a payment to a supplier for inventory or services (that is Supplier/Receipts). A separate accounting bucket."),
        new("Corporate", "Payroll",
            "Payroll-related messages and reminders."),
        new("Corporate", "Training",
            "Training reports, course reminders, and follow-ups — e.g. Metal Supermarkets University (MSU) / smarteru.com supervisor and completion reports."),
        new("Corporate", "IT & Security",
            "IT and security NOTIFICATIONS (not one-time codes): security alerts, suspicious or new sign-in notices, email-gateway / spam-quarantine notices, account-verification and account-management messages, and IT help-desk / support-ticket notifications. A bare one-time login / MFA / verification CODE is Other/Login Codes, not here."),

        new("Franchise", "Newsletters",
            "Newsletters, bulletins, and corporate/franchisor marketing — e.g. \"The Pipeline\", \"Fundamentals That Matter to Every Fabricator\", and mail from emarketing@ / communications@Metalsupermarkets.com. Automated operational reports about our store go to Franchise/Reports, not here."),
        new("Franchise", "Reports",
            "Automated reports from the franchisor / franchise systems about our store's operations - e.g. \"New Customer Report for Hackensack Store\", and recurring sales / activity / performance / customer reports generated by the franchise. Distinct from Franchise/Newsletters (bulletins and marketing) and from Corporate/Accounting (our own bookkeeper's mail)."),
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
