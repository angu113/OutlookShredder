using OutlookShredder.Proxy.Controllers;
using OutlookShredder.Proxy.Models;
using Xunit;

namespace OutlookShredder.Proxy.Tests;

// Behaviour tests for PhoneController.ClassifyBackfill — the pure decision the call-log CRM backfill makes
// per row (add a blank BpName / overwrite an existing one / leave unchanged). The phone-match itself uses
// the live CustomerCacheService.LookupAllByPhone and is verified by the dry-run on real data, not here.
public class CallLogBackfillTests
{
    private static PhoneCallLogRecord Rec(string? bp, string? contact = null, string? popup = null) =>
        new() { CallerPhone = "(973) 555-0100", BpName = bp, ContactName = contact, PopupMessage = popup };

    [Fact]
    public void Blank_bp_with_a_match_is_an_add()
    {
        var a = PhoneController.ClassifyBackfill(Rec(bp: null), "Acme Corp", "Jane Doe", "On hold");
        Assert.Equal(PhoneController.CrmBackfillAction.AddBp, a);
    }

    [Fact]
    public void Whitespace_only_bp_is_treated_as_blank_add()
    {
        var a = PhoneController.ClassifyBackfill(Rec(bp: "   "), "Acme Corp", null, null);
        Assert.Equal(PhoneController.CrmBackfillAction.AddBp, a);
    }

    [Fact]
    public void All_three_fields_already_equal_is_unchanged()
    {
        var a = PhoneController.ClassifyBackfill(
            Rec(bp: "Acme Corp", contact: "Jane Doe", popup: "On hold"), "Acme Corp", "Jane Doe", "On hold");
        Assert.Equal(PhoneController.CrmBackfillAction.Unchanged, a);
    }

    [Fact]
    public void Null_vs_empty_match_fields_count_as_equal()
    {
        // record has bp only (contact/popup null); match has bp + empty contact/popup -> still unchanged.
        var a = PhoneController.ClassifyBackfill(Rec(bp: "Acme Corp"), "Acme Corp", "", "");
        Assert.Equal(PhoneController.CrmBackfillAction.Unchanged, a);
    }

    [Fact]
    public void Trim_differences_count_as_equal()
    {
        var a = PhoneController.ClassifyBackfill(Rec(bp: " Acme Corp "), "Acme Corp", null, null);
        Assert.Equal(PhoneController.CrmBackfillAction.Unchanged, a);
    }

    [Fact]
    public void Different_bp_is_a_change_overwrite()
    {
        var a = PhoneController.ClassifyBackfill(Rec(bp: "Old Name"), "Acme Corp", null, null);
        Assert.Equal(PhoneController.CrmBackfillAction.ChangeBp, a);
    }

    [Fact]
    public void Same_bp_but_new_popup_is_a_change()
    {
        // bp matches but the customer gained a popup since the row was logged -> refresh (not blank -> Change).
        var a = PhoneController.ClassifyBackfill(Rec(bp: "Acme Corp"), "Acme Corp", null, "NEW: credit hold");
        Assert.Equal(PhoneController.CrmBackfillAction.ChangeBp, a);
    }
}
