namespace OutlookShredder.Proxy.Models;

public class WorkflowCard
{
    public int     SpItemId       { get; set; }
    public string  DocumentNumber { get; set; } = "";
    public string? CustomerName   { get; set; }
    public string? DocumentType   { get; set; }
    public string  Tab            { get; set; } = "Worklist"; // "Worklist" | "Delivery" | "Transfer"
    public string  AssignedDate   { get; set; } = "";           // ISO date "2026-05-05"
    public int     SortOrder      { get; set; }
    public string? Notes          { get; set; }
    public string? ErpSpItemId       { get; set; }
    public bool    IsCompleted       { get; set; }
    public string? DeliveryAddress   { get; set; }
    public string? DeliveryMethod    { get; set; }   // "Pickup", "Delivery", "UPS Ground", "Our Truck", etc.
    /// <summary>"Red" | "Amber" | "Green" | null</summary>
    public string? RagStatus         { get; set; }
    public string? DeliveryService   { get; set; }
    public bool    WasAutoCreated    { get; set; }
    /// <summary>The user this card belongs to (full name). Focus-chip cards: the creating Shredder
    /// user; auto-created cards: the doc's sales rep ("Customer Rep:") or the importer. Null = none.</summary>
    public string? OwnerUser         { get; set; }
}

public class CreateWorkflowCardRequest
{
    public string  DocumentNumber  { get; set; } = "";
    public string? CustomerName    { get; set; }
    public string? DocumentType    { get; set; }
    public string  Tab             { get; set; } = "Worklist";
    public string  AssignedDate    { get; set; } = "";
    public string? Notes           { get; set; }
    public string? ErpSpItemId     { get; set; }
    public string? DeliveryAddress { get; set; }
    public string? DeliveryMethod  { get; set; }
    public string? RagStatus       { get; set; }
    public string? DeliveryService { get; set; }
    public bool    WasAutoCreated  { get; set; }
    public string? OwnerUser       { get; set; }
}

public class UpdateWorkflowCardRequest
{
    public string? Tab          { get; set; }
    public string? AssignedDate { get; set; }
    public int?    SortOrder    { get; set; }
    public string? Notes           { get; set; }
    public bool?   IsCompleted     { get; set; }
    /// <summary>Pass "" to clear. Null means no change.</summary>
    public string? RagStatus       { get; set; }
    /// <summary>Pass "" to clear. Null means no change.</summary>
    public string? DeliveryService { get; set; }
    public bool?   WasAutoCreated  { get; set; }
}
