using Microsoft.Extensions.Logging.Abstractions;
using OutlookShredder.Proxy.Services;
using Xunit;

namespace OutlookShredder.Proxy.Tests;

// Behaviour tests for the "Customer Info" master CSV parser + schema that feed the enrichment load
// (SharePointService.EnrichCustomersAsync). Pure in/out — no SharePoint. Covers field mapping by header,
// name dedup, blank-name drop, unparseable-cell reporting, and the canonical change-detection forms.
public class CustomerInfoParserTests
{
    private static CustomerImportService NewSvc() => new(NullLogger<CustomerImportService>.Instance);

    private const string Header =
        "Business Partner,Active,Payment Terms,On Hold,Credit Line Limit,Margin Type";

    [Fact]
    public void Maps_cells_to_their_sharepoint_column_names()
    {
        var csv = Header + "\n" +
                  "Acme Corp,true,Net 30 Days,false,150000,*\n";

        var parsed = NewSvc().ParseCustomerInfo(csv);

        var row = Assert.Single(parsed.Rows);
        Assert.Equal("Acme Corp", row.Name);
        Assert.Equal("Net 30 Days", row.Fields["PaymentTerms"]);
        Assert.Equal("true",        row.Fields["Active"]);
        Assert.Equal("false",       row.Fields["OnHold"]);
        Assert.Equal("150000",      row.Fields["CreditLineLimit"]);
        Assert.Equal("*",           row.Fields["MarginType"]);
    }

    [Fact]
    public void Blank_cells_are_omitted_not_stored_as_empty()
    {
        var csv = Header + "\n" +
                  "Acme Corp,true,,false,,*\n";

        var row = Assert.Single(NewSvc().ParseCustomerInfo(csv).Rows);
        Assert.False(row.Fields.ContainsKey("PaymentTerms"));    // blank -> absent (won't overwrite SP)
        Assert.False(row.Fields.ContainsKey("CreditLineLimit"));
        Assert.Equal("true", row.Fields["Active"]);
    }

    [Fact]
    public void Duplicate_business_partner_keeps_first_and_reports_the_rest()
    {
        var csv = Header + "\n" +
                  "Acme Corp,true,Net 30 Days,false,150000,*\n" +
                  "Acme Corp,true,Immediate,true,0,*\n";

        var parsed = NewSvc().ParseCustomerInfo(csv);

        var row = Assert.Single(parsed.Rows);
        Assert.Equal("Net 30 Days", row.Fields["PaymentTerms"]);            // first kept
        Assert.Contains(parsed.Skipped, s => s.Name == "Acme Corp" && s.Reason.Contains("duplicate"));
    }

    [Fact]
    public void Blank_business_partner_rows_are_dropped_silently()
    {
        var csv = Header + "\n" +
                  ",true,Net 30 Days,false,150000,*\n" +     // export footer/total noise
                  "Acme Corp,true,Net 30 Days,false,150000,*\n";

        var parsed = NewSvc().ParseCustomerInfo(csv);

        Assert.Single(parsed.Rows);
        Assert.DoesNotContain(parsed.Skipped, s => string.IsNullOrEmpty(s.Name));
    }

    [Fact]
    public void Unparseable_numeric_cell_is_reported_and_the_row_still_loads()
    {
        var csv = Header + "\n" +
                  "Acme Corp,true,Net 30 Days,false,N/A,*\n";

        var parsed = NewSvc().ParseCustomerInfo(csv);

        var row = Assert.Single(parsed.Rows);
        Assert.False(row.Fields.ContainsKey("CreditLineLimit"));            // bad number not stored
        Assert.Equal("Net 30 Days", row.Fields["PaymentTerms"]);           // rest of the row still loads
        Assert.Contains(parsed.Skipped, s => s.Name == "Acme Corp" && s.Reason.Contains("unparseable"));
    }

    [Fact]
    public void Missing_business_partner_column_yields_a_warning_and_no_rows()
    {
        var csv = "Active,Payment Terms\ntrue,Net 30 Days\n";

        var parsed = NewSvc().ParseCustomerInfo(csv);

        Assert.Empty(parsed.Rows);
        Assert.Contains(parsed.Warnings, w => w.Contains("Business Partner"));
    }

    [Theory]
    [InlineData("150000", "150000.0", CustomerInfoSchema.Kind.Number)]   // trailing-zero noise ignored
    [InlineData("true", "True",       CustomerInfoSchema.Kind.Boolean)]  // case-insensitive bool
    [InlineData("Net 30", "Net 30",   CustomerInfoSchema.Kind.Text)]
    public void Canon_treats_equivalent_values_as_unchanged(string a, string b, CustomerInfoSchema.Kind kind)
    {
        Assert.Equal(CustomerInfoSchema.Canon(a, kind), CustomerInfoSchema.Canon(b, kind));
    }

    [Fact]
    public void Canon_treats_genuinely_different_values_as_changed()
    {
        Assert.NotEqual(
            CustomerInfoSchema.Canon("Net 30 Days", CustomerInfoSchema.Kind.Text),
            CustomerInfoSchema.Canon("Immediate",   CustomerInfoSchema.Kind.Text));
    }

    [Theory]
    [InlineData("false", true)]    // explicit NO -> inactive (ERP logical delete)
    [InlineData("False", true)]
    [InlineData("true", false)]    // active
    [InlineData("", false)]        // blank -> treated as active (don't drop unknown)
    public void IsInactive_flags_only_explicit_false(string activeRaw, bool expectedInactive)
    {
        var fields = new Dictionary<string, string?> { [CustomerInfoSchema.ActiveColumn] = activeRaw };
        Assert.Equal(expectedInactive, CustomerInfoSchema.IsInactive(fields));
    }

    [Fact]
    public void IsInactive_is_false_when_active_column_absent()
    {
        var fields = new Dictionary<string, string?> { ["PaymentTerms"] = "Net 30 Days" };
        Assert.False(CustomerInfoSchema.IsInactive(fields));
    }

    [Fact]
    public void ToTyped_produces_the_right_clr_types()
    {
        Assert.IsType<bool>  (CustomerInfoSchema.ToTyped("true", CustomerInfoSchema.Kind.Boolean));
        Assert.IsType<double>(CustomerInfoSchema.ToTyped("150000", CustomerInfoSchema.Kind.Number));
        Assert.Equal("Net 30 Days", CustomerInfoSchema.ToTyped("Net 30 Days", CustomerInfoSchema.Kind.Text));
        Assert.Null(CustomerInfoSchema.ToTyped("",    CustomerInfoSchema.Kind.Text));
        Assert.Null(CustomerInfoSchema.ToTyped("N/A", CustomerInfoSchema.Kind.Number));
    }
}
