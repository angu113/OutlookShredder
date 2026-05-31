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
            "Invoices sent after we receive goods (often metal) or for services we pay for (power, water, security, etc.)."),
        new("Supplier", "Statements",
            "Statements of account from suppliers and providers."),

        new("Customer", "Inquiries",
            "General requests for metal or fabrication."),
        new("Customer", "Website Inquiries",
            "Requests for material or fabrication sent to us from the franchise website."),
        new("Customer", "Orders",
            "Follow-ups on orders in progress; usually carry the sales order reference HSK-SOxxxxxxx in the subject; includes customer POs sent to us."),
        new("Customer", "Web Orders",
            "Orders booked from the franchise website, notified by email with an attached invoice/order to process."),

        new("Corporate", "Tax",
            "Items specifically related to tax — sales tax, corporation tax."),
        new("Corporate", "Accounting",
            "Messages from our accounting partner."),
        new("Corporate", "Receipts",
            "Email receipts (and placeholders for receipt PDFs/scans)."),
        new("Corporate", "Payroll",
            "Payroll-related messages and reminders."),

        new("Franchise", "Newsletters",
            "Newsletters, bulletins, and corporate/franchisor marketing — e.g. \"The Pipeline\", \"Fundamentals That Matter to Every Fabricator\", and mail from emarketing@ / communications@Metalsupermarkets.com."),

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
    public string? PoNumber       { get; set; }
    public string? SoNumber       { get; set; }
    public string? Amount         { get; set; }
    public string? Reasoning      { get; set; }
    public string  AiProvider     { get; set; } = "";
    public string  AiModel        { get; set; } = "";
    /// <summary>Full raw provider JSON, kept for audit/analysis (RawAiResponse in SP).</summary>
    [JsonIgnore]
    public string  RawResponse    { get; set; } = "";
}
