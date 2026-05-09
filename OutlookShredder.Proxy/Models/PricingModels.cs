namespace OutlookShredder.Proxy.Models;

public record PricingReportItem
{
    public string   SpItemId           { get; init; } = "";
    public string   RfqId              { get; init; } = "";
    public string   SupplierName       { get; init; } = "";
    public string   ProductName        { get; init; } = "";
    public string?  CatalogProductName { get; init; }
    public string?  Metal              { get; init; }
    public string?  Shape              { get; init; }
    public string[] SpecialConditions  { get; init; } = [];
    public bool     IsService          { get; init; }
    public double?  PricePerPound      { get; init; }
    public string   PriceSource        { get; init; } = "";
    public string   Confidence         { get; init; } = "";
    public string?  ConfidenceNote     { get; init; }
    public string?  AiNote             { get; init; }
    public DateTime ReceivedAt         { get; init; }
}

public record PricingCategory
{
    public string   Metal                { get; init; } = "";
    public string   Shape                { get; init; } = "";
    public string[] Conditions           { get; init; } = [];
    public string   CategoryKey          { get; init; } = "";
    public int      TotalQuotes          { get; init; }
    public int      HighConfidenceQuotes { get; init; }
    public double?  AvgPricePerPound     { get; init; }
    public double?  MinPricePerPound     { get; init; }
    public double?  MaxPricePerPound     { get; init; }
}

public record PricingReport
{
    public string   Date                  { get; init; } = "";
    public bool     FromCache             { get; init; }
    public DateTime GeneratedAt           { get; init; }
    public int      TotalSliRows          { get; init; }
    public int      RegretExcluded        { get; init; }
    public int      ServicesExcluded      { get; init; }
    public int      UnpricedExcluded      { get; init; }
    public int      HighConfidenceCount   { get; init; }
    public int      MediumConfidenceCount { get; init; }
    public int      LowConfidenceCount    { get; init; }

    public List<PricingCategory>   Categories               { get; init; } = [];
    public List<PricingReportItem> AllItems                 { get; init; } = [];
    public List<PricingReportItem> LowMediumConfidenceItems { get; init; } = [];
}
