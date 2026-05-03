using System.Text.Json.Serialization;

namespace OutlookShredder.Proxy.Models;

public class ErpExtraction
{
    [JsonPropertyName("is_erp_document")]
    public bool IsErpDocument { get; set; }

    [JsonPropertyName("document_type")]
    public string? DocumentType { get; set; }

    [JsonPropertyName("document_number")]
    public string? DocumentNumber { get; set; }

    [JsonPropertyName("customer_name")]
    public string? CustomerName { get; set; }

    [JsonPropertyName("customer_reference")]
    public string? CustomerReference { get; set; }

    [JsonPropertyName("document_date")]
    public string? DocumentDate { get; set; }

    [JsonPropertyName("total_amount")]
    public string? TotalAmount { get; set; }

    public string? Currency { get; set; }

    [JsonPropertyName("line_items")]
    public List<ErpLineItem> LineItems { get; set; } = [];

    public string? Notes { get; set; }
}

public class ErpLineItem
{
    public string? Description { get; set; }
    public string? Code { get; set; }
    public string? Quantity { get; set; }
    public string? Unit { get; set; }

    [JsonPropertyName("unit_price")]
    public string? UnitPrice { get; set; }

    [JsonPropertyName("total_price")]
    public string? TotalPrice { get; set; }
}

public class ErpDocumentRecord
{
    public string? SpItemId { get; set; }
    public string? DocumentNumber { get; set; }
    public string? DocumentType { get; set; }
    public string? DocumentDate { get; set; }
    public string? CustomerName { get; set; }
    public string? CustomerReference { get; set; }
    public string? TotalAmount { get; set; }
    public string? Currency    { get; set; }
    public string? FileName { get; set; }
    public string? PdfUrl { get; set; }
    public string? ReceivedAt { get; set; }
    public bool IsArchived { get; set; }
    public string? SourceMachine { get; set; }
    public string? SourceUser { get; set; }
    /// <summary>JSON array of ErpAnnotation objects added by the user in the Focus view.</summary>
    public string? UserAnnotations { get; set; }
}

/// <summary>User-applied stamp on an ERP document in the Focus view.</summary>
public class ErpAnnotation
{
    public string Label   { get; set; } = "";
    public string Color   { get; set; } = "#2E7D32";
    public string AddedAt { get; set; } = "";
    public string AddedBy { get; set; } = "";
}

public class ErpScanResult
{
    public string Folder { get; set; } = "";
    public int FilesFound { get; set; }
    public int AlreadyProcessed { get; set; }
    public int NonErpFiles { get; set; }
    public int ErpDocuments { get; set; }
    public int Errors { get; set; }
    public List<string> ProcessedFiles { get; set; } = [];
    public List<string> ErrorFiles { get; set; } = [];
}
