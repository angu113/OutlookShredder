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
