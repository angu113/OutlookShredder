using System.Text.Json.Serialization;

namespace OutlookShredder.Proxy.Models;

/// <summary>Catalog entry stored in analysis-cache/catalog.json</summary>
public record CatalogEntry(string Name, string? SearchKey);

/// <summary>SLI row stored in analysis-cache/sli-sample.json (only rows with a known SearchKey)</summary>
public record SliEntry(
    string ProductName,
    string? ProductSearchKey,
    string? CatalogProductName,
    string? SupplierName,
    string? RfqId);

/// <summary>
/// AI-extracted token record for one product name.
/// Stored in analysis-cache/catalog-tokens.json and sli-tokens.json.
/// JSON property names match what Claude returns so the same class can be
/// deserialised directly from the AI response batch.
/// </summary>
public class ProductTokens
{
    // Identity — populated from catalog/SLI context, not AI
    public string  Name       { get; set; } = "";
    public string? SearchKey  { get; set; }

    // AI-extracted structured fields
    [JsonPropertyName("metal")]      public string?   TkMetal      { get; set; }
    [JsonPropertyName("alloy")]      public string?   TkAlloy      { get; set; }
    [JsonPropertyName("temper")]     public string?   TkTemper     { get; set; }
    [JsonPropertyName("shape")]      public string?   TkShape      { get; set; }
    [JsonPropertyName("dims")]       public string?   TkDims       { get; set; }
    [JsonPropertyName("conditions")] public string[]  TkConditions { get; set; } = [];

    // Metadata
    public bool     TokenizationFailed { get; set; }
    public DateTime TokenizedAt        { get; set; }
}

/// <summary>One row in a match-test run</summary>
public class MatchCase
{
    public string  ProductName        { get; set; } = "";
    public string? SupplierName       { get; set; }
    public string? ExpectedSearchKey  { get; set; }
    public string? ExpectedCatalogName{ get; set; }
    public string? ActualSearchKey    { get; set; }
    public string? ActualCatalogName  { get; set; }
    public double  Score              { get; set; }
    public bool    IsHit              { get; set; }
    public bool    IsNoMatch          { get; set; }
    public string? FailReason         { get; set; }
}

/// <summary>Results from one match-test run, saved to analysis-cache/match-test-results.json</summary>
public class MatchTestRun
{
    public DateTime       RunAt      { get; set; }
    public int            Total      { get; set; }
    public int            Hits       { get; set; }
    public int            Misses     { get; set; }
    public int            NoMatches  { get; set; }
    public double         HitRate    => Total == 0 ? 0 : Math.Round((double)Hits / Total * 100, 1);
    public List<MatchCase>Cases      { get; set; } = [];
}
