using OutlookShredder.Proxy.Services;
using Xunit;

namespace OutlookShredder.Proxy.Tests;

// Behaviour tests for the OpenBravo Sales-Invoice CSV -> CustomerStatementDto parser feeding the
// ShadowCat statement generator. Pure in/out. Covers the Customer PO Number column (new), the
// grouping + due-date derivation, and the credit-memo / fully-paid / zero-balance exclusions that
// mirror the Shredder-side StatementBuilder.
public class StatementCsvParserTests
{
    private const string Header =
        "Invoice Date,Document No.,Business Partner,Customer PO Number,Total Gross Amount,Total Paid,Payment Terms";

    [Fact]
    public void Carries_customer_po_through_to_the_invoice_line()
    {
        var csv = Header + "\n" +
                  "06/01/2026,INV-1,ACME,PO-12345,100,0,Net 30 Days\n";

        var stmts = StatementCsvParser.Parse(csv);

        var acme = Assert.Single(stmts);
        var line = Assert.Single(acme.Invoices);
        Assert.Equal("PO-12345", line.CustomerPO);
        Assert.Equal("INV-1", line.InvoiceNumber);
        Assert.Equal("2026-06-01", line.InvoiceDate);
        Assert.Equal("2026-07-01", line.DueDate);   // Net 30 -> +30 days
        Assert.Equal(100m, line.Amount);
    }

    [Fact]
    public void Missing_po_value_yields_empty_string_not_null()
    {
        // PO column present in the header but blank on the row.
        var csv = Header + "\n" +
                  "06/01/2026,INV-1,ACME,,100,0,Net 30 Days\n";

        var line = Assert.Single(Assert.Single(StatementCsvParser.Parse(csv)).Invoices);
        Assert.Equal("", line.CustomerPO);
    }

    [Fact]
    public void Absent_po_column_is_tolerated_and_defaults_to_empty()
    {
        // Older export with no Customer PO Number column at all — must still parse (idxPo < 0 branch).
        var csv =
            "Invoice Date,Document No.,Business Partner,Total Gross Amount,Total Paid,Payment Terms\n" +
            "06/01/2026,INV-1,ACME,100,0,Net 30 Days\n";

        var line = Assert.Single(Assert.Single(StatementCsvParser.Parse(csv)).Invoices);
        Assert.Equal("", line.CustomerPO);
        Assert.Equal(100m, line.Amount);
    }

    [Fact]
    public void Excludes_credit_memos_paid_lines_and_zero_balance_customers()
    {
        var csv = Header + "\n" +
                  "06/01/2026,INV-1,ACME,PO-1,100,0,Net 30 Days\n" +    // outstanding -> kept
                  "06/02/2026,INV-2,ACME,PO-2,50,50,Net 30 Days\n" +    // fully paid -> dropped line
                  "06/03/2026,CM-1,ACME,PO-3,-25,0,Net 30 Days\n" +     // credit memo -> dropped
                  "06/01/2026,INV-9,PAID,PO-9,80,80,Net 30 Days\n";     // customer nets zero -> dropped

        var stmts = StatementCsvParser.Parse(csv);

        var acme = Assert.Single(stmts);                  // PAID customer excluded entirely
        Assert.Equal("ACME", acme.CustomerName);
        var line = Assert.Single(acme.Invoices);          // only the outstanding INV-1 survives
        Assert.Equal("INV-1", line.InvoiceNumber);
        Assert.Equal("PO-1", line.CustomerPO);
    }
}
