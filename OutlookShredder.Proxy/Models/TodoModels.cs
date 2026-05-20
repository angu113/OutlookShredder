namespace OutlookShredder.Proxy.Models;

public class ShredderTodo
{
    public string? SpItemId    { get; set; }
    public string  Title       { get; set; } = "";
    public string  Status      { get; set; } = "Open"; // Open | Claimed | Done
    public string? ClaimedBy   { get; set; }
    public string? CreatedBy   { get; set; }
    public string? Notes       { get; set; }
    public DateTimeOffset? DueDate      { get; set; }
    public string? RelatedRfqId { get; set; }
    public DateTimeOffset? CreatedAt   { get; set; }
    public DateTimeOffset? CompletedAt { get; set; }
    public string?         TodoId      { get; set; }
}

public class CreateTodoRequest
{
    public string  Title        { get; set; } = "";
    public string? CreatedBy    { get; set; }
    public string? Notes        { get; set; }
    public DateTimeOffset? DueDate      { get; set; }
    public string? RelatedRfqId { get; set; }
}

public class UpdateTodoRequest
{
    public string? Status       { get; set; }
    public string? ClaimedBy    { get; set; }
    public string? Notes        { get; set; }
    public DateTimeOffset? DueDate { get; set; }
    public bool ClearCompletedAt { get; set; }
}
