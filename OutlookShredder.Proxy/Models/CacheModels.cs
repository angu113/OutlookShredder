using System.Text.Json.Serialization;

namespace OutlookShredder.Proxy.Models;

public record CacheStatusDto
{
    [JsonPropertyName("name")]          public string Name { get; init; } = "";
    [JsonPropertyName("displayName")]   public string DisplayName { get; init; } = "";
    [JsonPropertyName("status")]        public string Status { get; init; } = "cold";
    [JsonPropertyName("cacheBuiltUtc")] public DateTime? CacheBuiltUtc { get; init; }
    [JsonPropertyName("lastDeltaUtc")]  public DateTime? LastDeltaUtc { get; init; }
    [JsonPropertyName("itemCount")]     public int ItemCount { get; init; }
    [JsonPropertyName("schemaVersion")] public int SchemaVersion { get; init; }
}

public record CacheConfigDto
{
    [JsonPropertyName("fullRefreshIntervalDays")] public int FullRefreshIntervalDays { get; set; } = 7;
    [JsonPropertyName("deltaBufferMinutes")]       public int DeltaBufferMinutes { get; set; } = 10;
    [JsonPropertyName("enabled")]                 public bool Enabled { get; set; } = true;
}

public record CacheStatusResponse
{
    [JsonPropertyName("caches")] public List<CacheStatusDto> Caches { get; init; } = [];
    [JsonPropertyName("config")] public CacheConfigDto Config { get; init; } = new();
}

public record RebuildResult
{
    [JsonPropertyName("type")]       public string Type { get; init; } = "";
    [JsonPropertyName("itemCount")]  public int ItemCount { get; init; }
    [JsonPropertyName("durationMs")] public long DurationMs { get; init; }
}

public record ArchiveRfqRef
{
    [JsonPropertyName("rfqId")]          public string RfqId { get; init; } = "";
    [JsonPropertyName("customerName")]   public string? CustomerName { get; init; }
    [JsonPropertyName("requester")]      public string? Requester { get; init; }
    [JsonPropertyName("notes")]          public string? Notes { get; init; }
    [JsonPropertyName("hskNumber")]      public string? HskNumber { get; init; }
    [JsonPropertyName("dateSent")]       public DateTime? DateSent { get; init; }
    [JsonPropertyName("emailRecipients")] public string? EmailRecipients { get; init; }
}

public record ArchiveCacheFile
{
    [JsonPropertyName("schemaVersion")]  public int SchemaVersion { get; init; }
    [JsonPropertyName("cacheBuiltUtc")]  public DateTime? CacheBuiltUtc { get; init; }
    [JsonPropertyName("lastDeltaUtc")]   public DateTime? LastDeltaUtc { get; init; }
    [JsonPropertyName("refs")]           public List<ArchiveRfqRef> Refs { get; init; } = [];
    [JsonPropertyName("rliByRfqId")]     public Dictionary<string, List<RliContextItem>> RliByRfqId { get; init; } = [];
}

public record ArchiveSearchRequest
{
    [JsonPropertyName("rfqId")]          public string? RfqId { get; init; }
    [JsonPropertyName("customerName")]   public string? CustomerName { get; init; }
    [JsonPropertyName("requester")]      public string? Requester { get; init; }
    [JsonPropertyName("supplierName")]   public string? SupplierName { get; init; }
    [JsonPropertyName("product")]        public string? Product { get; init; }
    [JsonPropertyName("hskNumber")]      public string? HskNumber { get; init; }
    [JsonPropertyName("notesContains")]  public string? NotesContains { get; init; }
    [JsonPropertyName("dateFrom")]       public DateTime? DateFrom { get; init; }
    [JsonPropertyName("dateTo")]         public DateTime? DateTo { get; init; }
    [JsonPropertyName("cursor")]         public DateTime? Cursor { get; init; }
    [JsonPropertyName("pageSize")]       public int PageSize { get; init; } = 20;
}

public record ArchiveSearchResponse
{
    [JsonPropertyName("sli")]        public List<Dictionary<string, object?>> Sli { get; init; } = [];
    [JsonPropertyName("rli")]        public List<object> Rli { get; init; } = [];
    [JsonPropertyName("totalCount")] public int TotalCount { get; init; }
    [JsonPropertyName("nextCursor")] public DateTime? NextCursor { get; init; }
}
