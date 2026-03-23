namespace OutlookShredder.Proxy.Models;

public class ExtractResponse
{
    public bool              Success    { get; set; }
    public RfqExtraction?    Extracted  { get; set; }
    public List<SpWriteResult> Rows     { get; set; } = [];
    public string?           Error      { get; set; }
}

public class SpWriteResult
{
    public string? ProductName { get; set; }
    public bool    Success     { get; set; }
    public string? SpItemId    { get; set; }
    public string? SpWebUrl    { get; set; }
    public string? Error       { get; set; }
}
