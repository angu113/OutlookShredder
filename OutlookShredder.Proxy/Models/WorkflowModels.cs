namespace OutlookShredder.Proxy.Models;

public class WorkflowCard
{
    public int     SpItemId       { get; set; }
    public string  DocumentNumber { get; set; } = "";
    public string? CustomerName   { get; set; }
    public string? DocumentType   { get; set; }
    public string  Tab            { get; set; } = "Processing"; // "Processing" | "Delivery"
    public string  AssignedDate   { get; set; } = "";           // ISO date "2026-05-05"
    public int     SortOrder      { get; set; }
    public string? Notes          { get; set; }
    public string? ErpSpItemId    { get; set; }
}

public class CreateWorkflowCardRequest
{
    public string  DocumentNumber { get; set; } = "";
    public string? CustomerName   { get; set; }
    public string? DocumentType   { get; set; }
    public string  Tab            { get; set; } = "Processing";
    public string  AssignedDate   { get; set; } = "";
    public string? Notes          { get; set; }
    public string? ErpSpItemId    { get; set; }
}

public class UpdateWorkflowCardRequest
{
    public string? Tab          { get; set; }
    public string? AssignedDate { get; set; }
    public int?    SortOrder    { get; set; }
    public string? Notes        { get; set; }
}
