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
    string? RfqId,
    string? SpItemId = null);

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
    public string? RfqId              { get; set; }
    public string? SpItemId           { get; set; }
    public string? ExpectedSearchKey  { get; set; }
    public string? ExpectedCatalogName{ get; set; }
    public string? ActualSearchKey    { get; set; }
    public string? ActualCatalogName  { get; set; }
    public double  Score              { get; set; }
    public bool    IsHit              { get; set; }
    public bool    IsNoMatch          { get; set; }
    public string? FailReason         { get; set; }
}

/// <summary>One entry in the IndustryDictionary SharePoint list</summary>
public class IndustryDictionaryEntry
{
    public string  Term         { get; set; } = "";
    /// <summary>abbreviation | standard | alloy_series | condition | temper | shape_alias</summary>
    public string  TermType     { get; set; } = "";
    /// <summary>Comma-separated shapes/metals where this term appears in the catalog+SLI data</summary>
    public string? AppliesTo    { get; set; }
    /// <summary>The token field value this term resolves to (e.g. "welded", "sch40", "a500")</summary>
    public string? MapsToToken  { get; set; }
    public string? Definition   { get; set; }
    public string? Examples     { get; set; }
    public int     CatalogCount { get; set; }
    public int     SliCount     { get; set; }
    [JsonIgnore] public string? SpItemId { get; set; }
}

/// <summary>One planned or applied GT audit action</summary>
public class GtAuditAction
{
    public string  ProductName        { get; set; } = "";
    public string? SupplierName       { get; set; }
    public string? RfqId              { get; set; }
    public string? SpItemId           { get; set; }
    public string  Action             { get; set; } = ""; // "clear", "update", "review", "skip"
    public string? OldSearchKey       { get; set; }
    public string? NewSearchKey       { get; set; }
    public string? NewCatalogName     { get; set; }
    public string? Note               { get; set; }
    public bool    Applied            { get; set; }
    public string? Error              { get; set; }
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

/// <summary>In-memory record loaded from the SupplierProductMappings SP list by SupplierProductMappingsCacheService.</summary>
public class SupplierProductMappingEntry
{
    public string  SupplierName       { get; set; } = "";
    public string  SupplierTerm       { get; set; } = "";
    public string? ProductSearchKey   { get; set; }
    public string? CatalogProductName { get; set; }
}

/// <summary>
/// Result from CatalogAnalysisService.MatchProductAsync — the best catalog match for one supplier product.
/// Source values: "user_mapping" (SupplierProductMappings list), "token_scorer" (live tokenisation + scoring),
/// or null when no match was found.
/// </summary>
public class TokenMatchResult
{
    public string? SearchKey    { get; set; }
    public string? CatalogName  { get; set; }
    public double  Score        { get; set; }
    public string? Source       { get; set; } // "user_mapping" | "token_scorer" | null
    public List<TokenMatchCandidate> TopCandidates { get; set; } = [];
}

/// <summary>One scored candidate returned alongside the best match for diagnostic/review UI.</summary>
public class TokenMatchCandidate
{
    public string? SearchKey   { get; set; }
    public string? CatalogName { get; set; }
    public double  Score       { get; set; }
}

/// <summary>
/// Row in the TokenMatchDiagnostics SharePoint list.
/// Records the RLI anchor result vs the token scorer result for every SLI write so agreement
/// can be reviewed and used to improve matching confidence.
/// ReviewStatus values: "pending" | "confirmed" | "rejected" | "overridden"
/// </summary>
public class TokenMatchDiagnosticEntry
{
    [JsonIgnore] public string? SpItemId       { get; set; }
    public string? RfqId              { get; set; }
    public string? SliSpItemId        { get; set; }
    public string? SupplierName       { get; set; }
    public string  ProductName        { get; set; } = "";
    /// <summary>Resolution source written to the SLI: "rli_anchor" | "token_scorer" | "user_mapping" | "fuzzy" | "none"</summary>
    public string? ResolvedSource     { get; set; }
    public string? RliMspc            { get; set; }
    public string? TokenMspc          { get; set; }
    public double  TokenScore         { get; set; }
    public bool    Agreed             { get; set; }
    public string  ReviewStatus       { get; set; } = "pending";
    public string? OverriddenMspc     { get; set; }
    public DateTime CreatedAt         { get; set; }
}

/// <summary>Aggregate stats over the TokenMatchDiagnostics list, returned by GET /api/token-match/diagnostics/stats</summary>
public class TokenMatchStats
{
    public int    Total         { get; set; }
    public int    Agreed        { get; set; }
    public int    Disagreed     { get; set; }
    public int    Pending       { get; set; }
    public int    Confirmed     { get; set; }
    public int    Rejected      { get; set; }
    public int    Overridden    { get; set; }
    public double AgreementRate => Total == 0 ? 0 : Math.Round((double)Agreed / Total * 100, 1);
}
