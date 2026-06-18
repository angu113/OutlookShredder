using System;
using System.Collections.Generic;
using System.Linq;
using OutlookShredder.Proxy.Services;
using Xunit;

// Per-user read-state rules (SupplierReadModel). Pure logic, no Graph. A user is "caught up" to their
// AdoptedAt clean-slate watermark, PLUS two explicit override sets: a read-set (overrides a
// watermark-unread) and an unread-set (overrides a watermark-read — so an OLD message can be marked
// unread). Changing any of this behaviour must update this test in the same commit.
public class SupplierReadModelTests
{
    private static readonly DateTimeOffset T0 = new(2026, 6, 1, 12, 0, 0, TimeSpan.Zero);
    private static HashSet<string> Set(params string[] ids) => new(ids, SupplierReadModel.IdComparer);
    private static readonly HashSet<string> None = Set();

    // ── IsUnread (watermark + overrides) ───────────────────────────────────────
    [Fact]
    public void Message_at_or_before_watermark_is_read()
    {
        Assert.False(SupplierReadModel.IsUnread(T0, "m1", T0, None, None));                 // exactly at watermark
        Assert.False(SupplierReadModel.IsUnread(T0.AddMinutes(-1), "m1", T0, None, None));  // before
    }

    [Fact]
    public void Message_after_watermark_is_unread_until_marked_read()
    {
        Assert.True(SupplierReadModel.IsUnread(T0.AddMinutes(1), "m1", T0, None, None));            // after, default
        Assert.False(SupplierReadModel.IsUnread(T0.AddMinutes(1), "m1", T0, Set("m1"), None));      // explicit read
    }

    [Fact]
    public void Old_message_can_be_marked_unread_via_the_unread_override()   // the bug Angus found
    {
        Assert.False(SupplierReadModel.IsUnread(T0.AddMinutes(-5), "old", T0, None, None));        // default read
        Assert.True(SupplierReadModel.IsUnread(T0.AddMinutes(-5), "old", T0, None, Set("old")));   // overridden unread
    }

    [Fact]
    public void Unread_override_wins_over_read_override()
        => Assert.True(SupplierReadModel.IsUnread(T0.AddMinutes(1), "m1", T0, Set("m1"), Set("m1")));

    [Fact]
    public void Null_time_with_no_override_is_read()
        => Assert.False(SupplierReadModel.IsUnread(null, "m1", T0, None, None));

    // ── Parse / serialize round-trip ───────────────────────────────────────────
    [Fact]
    public void Parse_handles_empty_whitespace_and_dedup()
    {
        Assert.Empty(SupplierReadModel.ParseReadIds(null));
        Assert.Empty(SupplierReadModel.ParseReadIds("  \n \n"));
        Assert.Equal(new[] { "a", "b" }, SupplierReadModel.ParseReadIds(" a \nb\nb\n").OrderBy(x => x));
    }

    [Fact]
    public void Serialize_then_parse_round_trips_case_sensitively()
    {
        var back = SupplierReadModel.ParseReadIds(SupplierReadModel.SerializeReadIds(new[] { "P", "p", "a", "P" }));
        Assert.Equal(3, back.Count);   // P != p (Ordinal); duplicate "P" collapses
        Assert.Contains("P", back); Assert.Contains("p", back); Assert.Contains("a", back);
    }

    // ── ApplyMark (moves an id between the read/unread override sets) ───────────
    [Fact]
    public void Mark_read_adds_to_read_set_and_clears_unread()
    {
        var s = SupplierReadModel.ApplyMark(None, Set("m1"), "m1", read: true);
        Assert.Contains("m1", s.Read);
        Assert.DoesNotContain("m1", s.Unread);
    }

    [Fact]
    public void Mark_unread_adds_to_unread_set_and_clears_read()
    {
        var s = SupplierReadModel.ApplyMark(Set("m1"), None, "m1", read: false);
        Assert.Contains("m1", s.Unread);
        Assert.DoesNotContain("m1", s.Read);
    }

    // ── Tally (per-user counts + cross-list dedup + grouping) ──────────────────
    private static SupplierReadModel.ReadRow Row(string rfq, string sup, string? msg, DateTimeOffset? t)
        => new(rfq, sup, msg, t);

    [Fact]
    public void Clean_slate_user_has_zero_unread()
    {
        var rows = new[] { Row("AW0001", "Eastern", "m1", T0.AddMinutes(-5)), Row("AW0001", "Eastern", "m2", T0.AddMinutes(-1)) };
        Assert.Equal(0, SupplierReadModel.Tally(rows, T0, None, None).Total);
    }

    [Fact]
    public void New_message_after_watermark_counts_until_marked()
    {
        var rows = new[] { Row("AW0001", "Eastern", "m3", T0.AddMinutes(10)) };
        Assert.Equal(1, SupplierReadModel.Tally(rows, T0, None, None).Total);
        Assert.Equal(0, SupplierReadModel.Tally(rows, T0, Set("m3"), None).Total);
    }

    [Fact]
    public void Old_message_marked_unread_counts()   // the fix for Angus's "snaps back" report
    {
        var rows = new[] { Row("AW0001", "Eastern", "old", T0.AddMinutes(-30)) };
        Assert.Equal(0, SupplierReadModel.Tally(rows, T0, None, None).Total);        // default read
        Assert.Equal(1, SupplierReadModel.Tally(rows, T0, None, Set("old")).Total);  // override unread -> counts
    }

    [Fact]
    public void Two_users_diverge()
    {
        var rows = new[] { Row("AW0001", "Eastern", "m3", T0.AddMinutes(10)) };
        Assert.Equal(0, SupplierReadModel.Tally(rows, T0, Set("m3"), None).Total);   // user A read it
        Assert.Equal(1, SupplierReadModel.Tally(rows, T0, None, None).Total);        // user B hasn't
    }

    [Fact]
    public void Same_id_in_both_lists_dedups_case_differs_counts_twice()
    {
        var after = T0.AddMinutes(10);
        var rows = new[]
        {
            Row("AW0001", "Eastern", "M", after),   // appears in both lists ...
            Row("AW0001", "Eastern", "M", after),   // ... dedups to one
            Row("AW0001", "Eastern", "m", after),   // different message (Ordinal) -> +1
        };
        var t = SupplierReadModel.Tally(rows, T0, None, None);
        Assert.Equal(2, t.Total);
        Assert.Equal(2, t.ByRfq["AW0001"]);
        Assert.Equal(2, t.BySupplier["AW0001|Eastern"]);
    }

    [Fact]
    public void Grouping_splits_by_rfq_and_supplier()
    {
        var after = T0.AddMinutes(10);
        var rows = new[]
        {
            Row("AW0001", "Eastern", "a", after),
            Row("AW0001", "Penn",    "b", after),
            Row("BX0002", "Eastern", "c", after),
        };
        var t = SupplierReadModel.Tally(rows, T0, None, None);
        Assert.Equal(3, t.Total);
        Assert.Equal(2, t.ByRfq["AW0001"]);
        Assert.Equal(1, t.ByRfq["BX0002"]);
        Assert.Equal(1, t.BySupplier["AW0001|Eastern"]);
        Assert.Equal(1, t.BySupplier["AW0001|Penn"]);
    }
}
