namespace OutlookShredder.Proxy.Models;

public record ForgeTaskRecord(
    string   SpItemId,
    string   TaskName,
    string?  TaskType,
    string?  ScheduleTime,
    bool     Enabled,
    DateTime? LastRunAt,
    string?  LastRunStatus,
    string?  LastRunMessage,
    string?  LastRunBy,
    string?  ResultData,
    string?  ResultCustomers
);

public record ForgeTaskQueueMessage(string TaskName);

/// <summary>
/// Health snapshot for a Forge scheduled task — the durable SharePoint record merged with the
/// serving proxy's in-memory state.  <see cref="Health"/> is derived: <c>ok</c> (success today),
/// <c>stale</c> (success but from a prior day), <c>fail</c>, <c>running</c>, or <c>unknown</c>.
/// </summary>
public record ForgeTaskStatus(
    string    TaskName,
    bool      Enabled,
    string?   ScheduleTime,
    string?   TaskType,
    string?   LastRunStatus,
    DateTime? LastRunAt,
    string?   LastRunMessage,
    string?   LastRunBy,
    bool      Running,
    bool      CacheLoaded,
    string    Health
);

public class CustomerStatementDto
{
    public string              CustomerName { get; init; } = "";
    public string              Terms        { get; init; } = "";
    public List<InvoiceLineDto> Invoices    { get; init; } = [];
}

public class InvoiceLineDto
{
    public string  InvoiceDate   { get; init; } = "";   // yyyy-MM-dd
    public string  InvoiceNumber { get; init; } = "";
    public string? DueDate       { get; init; }          // yyyy-MM-dd or null
    public decimal Amount        { get; init; }
}
