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
    /// <summary>Worklist progress: null/"" (no status, e.g. Prioritize) | "InProgress" | "Ready".</summary>
    public string? Status            { get; set; }
    /// <summary>Customer told their order is Ready (only meaningful once Status = "Ready").</summary>
    public bool    CustomerNotified  { get; set; }
    /// <summary>Shop ops parsed from the slip (comma-joined, e.g. "Laser Cutting, Bending"), rendered as
    /// fabrication-type chips on the Worklist card. Null/empty = a plain cut job ("Cutting").</summary>
    public string? ProcessOps        { get; set; }
    /// <summary>Customer contact person from the slip's "Attention:" line. Null if not on the slip (the card
    /// then falls back to the CRM primary contact for display).</summary>
    public string? ContactName        { get; set; }
    /// <summary>Customer contact phone from the slip's "Contact Phone:" line. Null if not on the slip.</summary>
    public string? ContactPhone        { get; set; }
    /// <summary>Pickup Location # (1–20) assigned when the order is marked Ready. Null/0 = unset.</summary>
    public int?    LocationNumber      { get; set; }
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
    public string? Status          { get; set; }
    public bool    CustomerNotified { get; set; }
    public string? ProcessOps      { get; set; }
    public string? ContactName     { get; set; }
    public string? ContactPhone    { get; set; }
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
    /// <summary>Worklist progress: "" (none) | "InProgress" | "Ready". Pass "" to clear; null = no change.</summary>
    public string? Status          { get; set; }
    public bool?   CustomerNotified { get; set; }
    /// <summary>Pickup Location # (1–20). Null = no change.</summary>
    public int?    LocationNumber  { get; set; }
}
