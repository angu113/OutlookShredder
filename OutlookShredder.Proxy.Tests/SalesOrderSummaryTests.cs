using System.Collections.Generic;
using OutlookShredder.Proxy.Models;
using OutlookShredder.Proxy.Services;
using Xunit;

namespace OutlookShredder.Proxy.Tests;

// The Raptor call-card sales summary: the headline $ total (LifetimeNet) counts ONLY realized revenue —
// Booked orders with secondary status Completed. Quotes, open/in-progress, and cancelled orders are
// excluded from the dollar figure, but the order COUNT + the listed orders stay the full history.
// Changing this behaviour must update this test in the same commit.
public class SalesOrderSummaryTests
{
    private static SalesOrderRecord Order(string id, string? status, string? secondary, double net) =>
        new() { OrderId = id, CustomerName = "Acme", Status = status, SecondaryStatus = secondary, NetAmount = net };

    [Fact]
    public void Lifetime_total_counts_only_booked_completed()
    {
        var orders = new List<SalesOrderRecord>
        {
            Order("HSK-SO1", "Booked", "Completed", 100),   // counts
            Order("HSK-SO2", "Booked", "Completed", 250),   // counts
            Order("HSK-SO3", "Booked", "In Progress", 999), // excluded (secondary not Completed)
            Order("HSK-SO4", "Quote",  "Completed", 500),   // excluded (not Booked)
            Order("HSK-SO5", "Booked", null, 700),          // excluded (no secondary)
            Order("HSK-SO6", null,     "Completed", 800),   // excluded (no status)
        };

        var resp = SalesOrderHistoryService.Project(orders, top: 10);

        Assert.Equal(350, resp.Summary.LifetimeNet, 3);   // 100 + 250 only
        Assert.Equal(6, resp.Summary.OrderCount);         // count is the full history, unfiltered
        Assert.Equal(6, resp.Orders.Count);               // all orders still listed
    }

    [Fact]
    public void Booked_completed_match_is_case_and_whitespace_insensitive()
    {
        var orders = new List<SalesOrderRecord>
        {
            Order("HSK-SO1", " booked ", "completed", 40),
            Order("HSK-SO2", "BOOKED",   "COMPLETED ", 60),
        };

        var resp = SalesOrderHistoryService.Project(orders, top: 10);
        Assert.Equal(100, resp.Summary.LifetimeNet, 3);
    }

    [Fact]
    public void No_booked_completed_orders_means_zero_total_but_real_count()
    {
        var orders = new List<SalesOrderRecord>
        {
            Order("HSK-SO1", "Quote",  "Completed",   500),
            Order("HSK-SO2", "Booked", "In Progress", 300),
        };

        var resp = SalesOrderHistoryService.Project(orders, top: 10);
        Assert.Equal(0, resp.Summary.LifetimeNet, 3);
        Assert.Equal(2, resp.Summary.OrderCount);
    }
}
