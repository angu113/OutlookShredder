using OutlookShredder.Proxy.Models;
using OutlookShredder.Proxy.Services;
using Xunit;

namespace OutlookShredder.Proxy.Tests;

// Behaviour tests for the PO -> picking-slip dependency rules: a slip is pinned to the LATEST receipt
// date of the active POs whose SalesOrders include the slip's HSK-SO. Pure (records in, pins out).
public class PoSlipDependencyTests
{
    private static PurchaseOrderRecord Po(string po, string salesOrders,
        string? boardDate = null, string? eta = null, string? received = null) =>
        new() { SpItemId = po, PoNumber = po, RfqId = "R", SupplierName = "S",
                SalesOrders = salesOrders, BoardDate = boardDate, ExpectedDate = eta, MaterialReceivedAt = received };

    private static WorkflowCard Slip(int id, string so, string assigned = "", bool done = false) =>
        new() { SpItemId = id, DocumentNumber = so, AssignedDate = assigned, Tab = "Worklist", IsCompleted = done };

    [Fact]
    public void ReceiptDate_prefers_board_date_then_eta_and_is_normalized()
    {
        Assert.Equal("2026-06-20", PoSlipDependency.ReceiptDate(Po("P", "X", boardDate: "2026-06-20", eta: "2026-06-25")));
        Assert.Equal("2026-06-25", PoSlipDependency.ReceiptDate(Po("P", "X", eta: "2026-06-25T00:00:00")));
        Assert.Null(PoSlipDependency.ReceiptDate(Po("P", "X")));                                  // unscheduled
        Assert.Null(PoSlipDependency.ReceiptDate(Po("P", "X", boardDate: "2026-06-20", received: "2026-06-21"))); // received → released
    }

    [Fact]
    public void Pins_a_scheduled_slip_to_its_pos_receipt_date()
    {
        var pos   = new[] { Po("HSK-PO1", "HSK-SO9", boardDate: "2026-06-20") };
        var slips = new[] { Slip(1, "HSK-SO9", assigned: "2026-06-15") };   // booked too early
        var pins  = PoSlipDependency.ComputePins(pos, slips);
        var (slip, date) = Assert.Single(pins);
        Assert.Equal(1, slip.SpItemId);
        Assert.Equal("2026-06-20", date);
    }

    [Fact]
    public void Pins_to_the_latest_date_when_material_is_split_across_pos()
    {
        var pos = new[]
        {
            Po("HSK-PO1", "HSK-SO9", boardDate: "2026-06-20"),
            Po("HSK-PO2", "HSK-SO9, HSK-SO8", eta: "2026-06-27"),   // later — governs
        };
        var slips = new[] { Slip(1, "HSK-SO9", assigned: "2026-06-20") };
        var (_, date) = Assert.Single(PoSlipDependency.ComputePins(pos, slips));
        Assert.Equal("2026-06-27", date);
    }

    [Fact]
    public void Already_pinned_slip_produces_no_change()
    {
        var pos   = new[] { Po("HSK-PO1", "HSK-SO9", boardDate: "2026-06-20") };
        var slips = new[] { Slip(1, "HSK-SO9", assigned: "2026-06-20") };
        Assert.Empty(PoSlipDependency.ComputePins(pos, slips));
    }

    [Fact]
    public void Unscheduled_completed_and_ungoverned_slips_are_left_alone()
    {
        var pos = new[] { Po("HSK-PO1", "HSK-SO9", boardDate: "2026-06-20") };
        var slips = new[]
        {
            Slip(1, "HSK-SO9", assigned: ""),                    // Prioritize — leave
            Slip(2, "HSK-SO9", assigned: "2026-06-15", done: true), // completed — ignore
            Slip(3, "HSK-SO7", assigned: "2026-06-15"),          // no matching PO — free
        };
        Assert.Empty(PoSlipDependency.ComputePins(pos, slips));
    }

    [Fact]
    public void Received_po_releases_the_slip()
    {
        var pos   = new[] { Po("HSK-PO1", "HSK-SO9", boardDate: "2026-06-20", received: "2026-06-19") };
        var slips = new[] { Slip(1, "HSK-SO9", assigned: "2026-06-15") };
        Assert.Empty(PoSlipDependency.ComputePins(pos, slips));   // no longer governed
    }

    [Fact]
    public void PinnedDateFor_returns_latest_or_null()
    {
        var pos = new[]
        {
            Po("HSK-PO1", "HSK-SO9", boardDate: "2026-06-20"),
            Po("HSK-PO2", "HSK-SO9", eta: "2026-06-27"),
        };
        Assert.Equal("2026-06-27", PoSlipDependency.PinnedDateFor("HSK-SO9", pos));
        Assert.Null(PoSlipDependency.PinnedDateFor("HSK-SO1", pos));   // ungoverned
        Assert.Null(PoSlipDependency.PinnedDateFor("", pos));
    }
}
