using OutlookShredder.Proxy.Controllers;
using OutlookShredder.Proxy.Models;
using OutlookShredder.Proxy.Services;
using Xunit;

namespace OutlookShredder.Proxy.Tests;

// Behaviour tests for CallLogCrmBackfillService.Classify — the pure decision the call-log CRM backfill makes
// per row (add a blank BpName / overwrite an existing one / leave unchanged). The phone-match itself uses
// the live CustomerCacheService.LookupAllByPhone and is verified by the dry-run on real data, not here.
public class CallLogBackfillTests
{
    private static PhoneCallLogRecord Rec(string? bp, string? contact = null, string? popup = null) =>
        new() { CallerPhone = "(973) 555-0100", BpName = bp, ContactName = contact, PopupMessage = popup };

    [Fact]
    public void Blank_bp_with_a_match_is_an_add()
    {
        var a = CallLogCrmBackfillService.Classify(Rec(bp: null), "Acme Corp", "Jane Doe", "On hold");
        Assert.Equal(CallLogCrmBackfillService.CrmBackfillAction.AddBp, a);
    }

    [Fact]
    public void Whitespace_only_bp_is_treated_as_blank_add()
    {
        var a = CallLogCrmBackfillService.Classify(Rec(bp: "   "), "Acme Corp", null, null);
        Assert.Equal(CallLogCrmBackfillService.CrmBackfillAction.AddBp, a);
    }

    [Fact]
    public void All_three_fields_already_equal_is_unchanged()
    {
        var a = CallLogCrmBackfillService.Classify(
            Rec(bp: "Acme Corp", contact: "Jane Doe", popup: "On hold"), "Acme Corp", "Jane Doe", "On hold");
        Assert.Equal(CallLogCrmBackfillService.CrmBackfillAction.Unchanged, a);
    }

    [Fact]
    public void Null_vs_empty_match_fields_count_as_equal()
    {
        // record has bp only (contact/popup null); match has bp + empty contact/popup -> still unchanged.
        var a = CallLogCrmBackfillService.Classify(Rec(bp: "Acme Corp"), "Acme Corp", "", "");
        Assert.Equal(CallLogCrmBackfillService.CrmBackfillAction.Unchanged, a);
    }

    [Fact]
    public void Trim_differences_count_as_equal()
    {
        var a = CallLogCrmBackfillService.Classify(Rec(bp: " Acme Corp "), "Acme Corp", null, null);
        Assert.Equal(CallLogCrmBackfillService.CrmBackfillAction.Unchanged, a);
    }

    [Fact]
    public void Different_bp_is_a_change_overwrite()
    {
        var a = CallLogCrmBackfillService.Classify(Rec(bp: "Old Name"), "Acme Corp", null, null);
        Assert.Equal(CallLogCrmBackfillService.CrmBackfillAction.ChangeBp, a);
    }

    [Fact]
    public void Same_bp_but_new_popup_is_a_change()
    {
        // bp matches but the customer gained a popup since the row was logged -> refresh (not blank -> Change).
        var a = CallLogCrmBackfillService.Classify(Rec(bp: "Acme Corp"), "Acme Corp", null, "NEW: credit hold");
        Assert.Equal(CallLogCrmBackfillService.CrmBackfillAction.ChangeBp, a);
    }

    // The targeted UpdateCallLogCrmByPhoneAsync matches a stored call-log row against the link target by
    // NORMALIZED phone, not a raw string compare / SP $filter. Stored CallerPhone is formatted
    // ("(201) 868-0293") while the manual-link/lookup target is digits ("2018680293"); both must collapse to
    // the same key or the backfill silently matches nothing (the broken-Kamino regression, 2026-06-20).
    [Theory]
    [InlineData("(201) 868-0293", "2018680293")]
    [InlineData("201-868-0293", "2018680293")]
    [InlineData("201.868.0293", "2018680293")]
    [InlineData("+1 (201) 868-0293", "2018680293")]   // 11-digit country code dropped
    [InlineData(" 2018680293 ", "2018680293")]
    public void Stored_phone_formatting_matches_digit_target(string storedCallerPhone, string linkTarget)
    {
        Assert.Equal(
            CustomerImportService.NormalizePhone(linkTarget),
            CustomerImportService.NormalizePhone(storedCallerPhone));
    }

    [Fact]
    public void Different_numbers_do_not_match()
    {
        Assert.NotEqual(
            CustomerImportService.NormalizePhone("2018680293"),
            CustomerImportService.NormalizePhone("2018680292"));
    }
}
