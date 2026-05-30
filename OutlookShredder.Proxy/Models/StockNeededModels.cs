namespace OutlookShredder.Proxy.Models;

public class StockNeededItem
{
    public int     SpItemId         { get; set; }
    public string  ProductName      { get; set; } = "";
    public string? ProductSearchKey { get; set; }
    public string? Category         { get; set; }
    public string? Shape            { get; set; }
    public string? QuantityNeeded   { get; set; }
    public string? SizeRequested    { get; set; }
    public string? Notes            { get; set; }
    public string? RfqId            { get; set; }
    public DateTime CreatedAt       { get; set; }
    public string? CreatedBy        { get; set; }
}

public class CreateStockNeededItemRequest
{
    public string  ProductName      { get; set; } = "";
    public string? ProductSearchKey { get; set; }
    public string? Category         { get; set; }
    public string? Shape            { get; set; }
    public string? QuantityNeeded   { get; set; }
    public string? SizeRequested    { get; set; }
    public string? Notes            { get; set; }
    public string? CreatedBy        { get; set; }
}

public class PatchStockNeededItemRequest
{
    public string? ProductName      { get; set; }
    public string? ProductSearchKey { get; set; }
    public string? Category         { get; set; }
    public string? Shape            { get; set; }
    public string? QuantityNeeded   { get; set; }
    public string? SizeRequested    { get; set; }
    public string? Notes            { get; set; }
    public string? RfqId            { get; set; }
}
