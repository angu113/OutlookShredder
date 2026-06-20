using System.Text;
using Microsoft.Extensions.Logging.Abstractions;
using OutlookShredder.Proxy.Services;
using Xunit;

namespace OutlookShredder.Proxy.Tests;

// Behaviour tests for the Sales-Orders contact parser (CustomerImportService.ParseContactsFromOrders),
// which mines the orders export's "Contact" column ("[phone] - [name] - ") into (Customer, ContactName,
// Phone) triples. Pure in/out — no SharePoint (the existing-customer-only filter is applied downstream in
// SharePointService.UpsertContactsExistingOnlyAsync and is not exercised here).
public class OrderContactsParserTests
{
    private static CustomerImportService NewSvc() => new(NullLogger<CustomerImportService>.Instance);

    // Builds a minimal orders CSV (the parser only needs Customer + Contact; Order Date + Doc # are
    // included so the rows look like the real export) from (customer, contact-cell) pairs.
    private static CustomerImportService.ParseResult<CustomerImportService.ContactRow> ParseRows(
        params (string Customer, string Contact)[] rows)
    {
        var sb = new StringBuilder();
        sb.AppendLine("\"Order Date\",\"Doc #\",\"Customer\",\"Contact\"");
        int n = 0;
        foreach (var (customer, contact) in rows)
            sb.AppendLine($"\"06-19-2026\",\"HSK-SO{1000 + n++}\",\"{customer}\",\"{contact}\"");
        return NewSvc().ParseContactsFromOrders(sb.ToString());
    }

    private static CustomerImportService.ParseResult<CustomerImportService.ContactRow> ParseOne(
        string customer, string contact) => ParseRows((customer, contact));

    [Theory]
    [InlineData("(555) 010-1234 - Jane Doe - ",            "5550101234", "Jane Doe")]      // standard formatted
    [InlineData("555 010 5678 Ex 355 - John Smith - ",     "5550105678", "John Smith")]    // extension dropped
    [InlineData("555 010 1111/555 010 2222 - Pat Lee - ",  "5550101111", "Pat Lee")]       // first of two numbers
    [InlineData("1 555 010 9999 - Sam Roe - ",             "5550109999", "Sam Roe")]       // leading country code dropped
    [InlineData("201 449 3553  - Richard Caicedo - ",      "2014493553", "Richard Caicedo")] // double-space delimiter
    [InlineData("9730990995 - Jose Santos - ",             "9730990995", "Jose Santos")]   // unformatted 10 digits
    public void Parses_phone_and_name_from_contact_cell(string contact, string expectedPhone, string expectedName)
    {
        var parsed = ParseOne("Acme Corp", contact);

        var row = Assert.Single(parsed.Rows);
        Assert.Equal("Acme Corp",    row.CustomerName);
        Assert.Equal(expectedPhone,  row.Phone);
        Assert.Equal(expectedName,   row.ContactName);
    }

    [Fact]
    public void Name_only_cell_with_no_phone_is_skipped_and_reported()
    {
        var parsed = ParseOne("Acme Corp", "Accounting - ");

        Assert.Empty(parsed.Rows);
        Assert.Contains(parsed.Skipped, s => s.Reason.Contains("phone"));
    }

    [Fact]
    public void Phone_present_but_no_name_is_skipped()
    {
        var parsed = ParseOne("Acme Corp", "555 010 1234 -  - ");

        Assert.Empty(parsed.Rows);
        Assert.Contains(parsed.Skipped, s => s.Reason.Contains("name"));
    }

    [Fact]
    public void Duplicate_tuples_across_orders_collapse_to_one()
    {
        // The same customer/contact recurs on every order — must dedup to a single (cust, name, phone) row.
        var parsed = ParseRows(
            ("Acme Corp", "555 010 1234 - Jane Doe - "),
            ("Acme Corp", "555 010 1234 - Jane Doe - "),
            ("Acme Corp", "555 010 1234 - Jane Doe - "));

        Assert.Single(parsed.Rows);
    }

    [Fact]
    public void Two_distinct_contacts_at_one_customer_both_kept()
    {
        var parsed = ParseRows(
            ("Acme Corp", "555 010 1111 - Jane Doe - "),
            ("Acme Corp", "555 010 2222 - John Roe - "));

        Assert.Equal(2, parsed.Rows.Count);
    }

    [Fact]
    public void Repeated_unparseable_contact_is_reported_only_once()
    {
        var parsed = ParseRows(
            ("Acme Corp", "Accounting - "),
            ("Acme Corp", "Accounting - "),
            ("Acme Corp", "Accounting - "));

        Assert.Single(parsed.Skipped);
    }

    [Fact]
    public void Missing_required_columns_yields_a_warning_and_no_rows()
    {
        var csv = "\"Order Date\",\"Doc #\",\"Status\"\n\"06-19-2026\",\"HSK-SO1\",\"Booked\"\n";

        var parsed = NewSvc().ParseContactsFromOrders(csv);

        Assert.Empty(parsed.Rows);
        Assert.Contains(parsed.Warnings, w => w.Contains("Customer") || w.Contains("Contact"));
    }

    [Fact]
    public void Blank_customer_or_blank_contact_cell_is_dropped()
    {
        var parsed = ParseRows(
            ("",          "555 010 1234 - Jane Doe - "),   // no customer
            ("Acme Corp", ""),                              // no contact
            ("Acme Corp", "555 010 9999 - Real One - "));   // the only keepable row

        var row = Assert.Single(parsed.Rows);
        Assert.Equal("Real One", row.ContactName);
    }

    [Theory]
    [InlineData("(555) 010-1234 - Jane Doe - ",        "5550101234", "Jane Doe")]
    [InlineData("555 010 5678 Ex 355 - John Smith - ", "5550105678", "John Smith")]
    [InlineData("Accounting - ",                       null,         "")]   // nothing after the delimiter -> empty name, no phone
    public void ParseOrderContact_splits_phone_and_name(string cell, string? expectedPhone, string expectedName)
    {
        var (phone, name) = CustomerImportService.ParseOrderContact(cell);

        Assert.Equal(expectedPhone, phone);
        Assert.Equal(expectedName,  name);
    }
}
