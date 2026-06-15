using OutlookShredder.Proxy.Services;
using Xunit;

namespace OutlookShredder.Proxy.Tests;

// Pins the Trigger Delivery-lane auto-create rule: any non-pickup "Delivery Method:" routes to Delivery.
// The old rule keyed only on the literal word "Delivery" and silently dropped real deliveries like
// "Our Truck" (the value on HSK-SO1036170) — these cases guard against that regression returning.
public class WorkflowDeliveryMethodTests
{
    [Theory]
    [InlineData("Our Truck")]      // company truck — the most common real delivery value
    [InlineData("our truck")]      // case-insensitive
    [InlineData("Delivery")]       // plain "Delivery" still counts
    [InlineData("UPS Ground")]     // common carrier
    [InlineData("FedEx")]
    [InlineData("Common Carrier")]
    [InlineData("Freight Collect")] // freight-collect is still a carrier delivery
    [InlineData("  Our Truck  ")]  // surrounding whitespace trimmed
    public void Counts_non_pickup_methods_as_delivery(string method)
        => Assert.True(WorkflowCardService.IsDeliveryMethod(method));

    [Theory]
    [InlineData("Pickup")]
    [InlineData("Customer Pickup")]
    [InlineData("PICK UP")]
    [InlineData("Will Call")]
    [InlineData("Will-Call")]
    [InlineData("Counter Pickup")]
    [InlineData(null)]             // unknown → not a delivery (no card)
    [InlineData("")]
    [InlineData("   ")]
    public void Excludes_pickup_and_empty(string? method)
        => Assert.False(WorkflowCardService.IsDeliveryMethod(method));
}
