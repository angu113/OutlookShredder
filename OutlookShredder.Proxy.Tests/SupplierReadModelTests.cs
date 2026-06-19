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

    // ── ParseSpInstant (UTC normalization of SP timestamps) ────────────────────
    [Fact]
    public void ParseSpInstant_treats_a_tzless_value_as_utc_not_local()   // the first-load leak root cause
    {
        // SharePoint echoes these columns with the UTC clock value but NO 'Z'/offset. A bare TryParse would
        // attach the machine's local offset (timezone-dependent); ParseSpInstant must pin it to UTC so it
        // compares correctly against the DateTimeOffset.UtcNow clean-slate watermark.
        var dt = SupplierReadModel.ParseSpInstant("2026-06-19T11:37:36");
        Assert.NotNull(dt);
        Assert.Equal(new DateTimeOffset(2026, 6, 19, 11, 37, 36, TimeSpan.Zero), dt!.Value);
        Assert.Equal(TimeSpan.Zero, dt!.Value.Offset);
    }

    [Fact]
    public void ParseSpInstant_normalizes_offset_bearing_values_to_utc()
    {
        Assert.Equal(new DateTimeOffset(2026, 6, 19, 11, 37, 36, TimeSpan.Zero),
            SupplierReadModel.ParseSpInstant("2026-06-19T11:37:36Z"));
        Assert.Equal(new DateTimeOffset(2026, 6, 19, 12, 0, 0, TimeSpan.Zero),
            SupplierReadModel.ParseSpInstant("2026-06-19T08:00:00-04:00"));   // -04:00 -> 12:00Z (same instant)
    }

    [Fact]
    public void ParseSpInstant_returns_null_on_empty_or_garbage()
    {
        Assert.Null(SupplierReadModel.ParseSpInstant(null));
        Assert.Null(SupplierReadModel.ParseSpInstant(""));
        Assert.Null(SupplierReadModel.ParseSpInstant("   "));
        Assert.Null(SupplierReadModel.ParseSpInstant("not-a-date"));
    }

    [Fact]
    public void Fresh_user_first_tally_ignores_a_tzless_received_time_before_the_watermark()   // first-load leak
    {
        // Repro of the shipped bug: a brand-new user's watermark is DateTimeOffset.UtcNow (true UTC); the
        // message's ReceivedAt comes back from SP tz-less. Parsed as UTC, an older message stays read (0),
        // so the user does NOT inherit another user's unread count on first load.
        var watermark = new DateTimeOffset(2026, 6, 19, 12, 29, 57, TimeSpan.Zero);
        var msgTime   = SupplierReadModel.ParseSpInstant("2026-06-19T11:37:36");   // ~52 min before the watermark
        Assert.False(SupplierReadModel.IsUnread(msgTime, "m1", watermark, None, None));
        Assert.Equal(0, SupplierReadModel.Tally(new[] { Row("JM0079", "J F Fazzio", "m1", msgTime) }, watermark, None, None).Total);
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
    public void Junk_or_misclassified_rfq_id_is_not_counted()   // WHOIS badge-desync bug Angus found
    {
        var after = T0.AddMinutes(10);
        var rows = new[]
        {
            Row("WHOIS",    "Unknown", "j1", after),   // 5 chars -> not a real RFQ; grid shows no group to badge
            Row("123456",   "Unknown", "j2", after),   // no leading letters
            Row("[000000]", "Unknown", "j3", after),   // orphan sentinel (8 chars)
            Row("AW0001",   "Eastern", "ok", after),   // a real RFQ -> the only one that counts
        };
        var t = SupplierReadModel.Tally(rows, T0, None, None);
        Assert.Equal(1, t.Total);
        Assert.Equal(1, t.ByRfq["AW0001"]);
        Assert.False(t.ByRfq.ContainsKey("WHOIS"));
        Assert.False(t.BySupplier.ContainsKey("WHOIS|Unknown"));
    }

    [Fact]
    public void IsCountableRfq_mirrors_the_grid_validity_rule()
    {
        Assert.True(SupplierReadModel.IsCountableRfq("AW0001"));
        Assert.True(SupplierReadModel.IsCountableRfq("JM12AB"));
        Assert.False(SupplierReadModel.IsCountableRfq(null));
        Assert.False(SupplierReadModel.IsCountableRfq("WHOIS"));     // 5 chars
        Assert.False(SupplierReadModel.IsCountableRfq("AW00012"));   // 7 chars
        Assert.False(SupplierReadModel.IsCountableRfq("1W0001"));    // first char not a letter
        Assert.False(SupplierReadModel.IsCountableRfq("A_0001"));    // non-alphanumeric
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
