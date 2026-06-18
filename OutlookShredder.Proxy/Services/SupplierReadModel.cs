namespace OutlookShredder.Proxy.Services;

/// <summary>
/// Pure, Graph-free per-user read-state logic for RFQ supplier messages. Unread is now PER USER, not
/// team-wide: a user is "caught up" to their <c>AdoptedAt</c> clean-slate watermark (every inbound
/// message at/before it counts as already-read), plus an explicit set of MessageIds they've marked read
/// AFTER that watermark. Keeping this Graph-free makes it unit-testable; <c>SharePointService</c> calls
/// these helpers with rows + the caller's profile.
/// </summary>
public static class SupplierReadModel
{
    /// <summary>Graph MessageIds are case-SENSITIVE base64 — always compare/store Ordinal.</summary>
    public static readonly StringComparer IdComparer = StringComparer.Ordinal;

    /// <summary>One inbound message row for tallying (deduped across the two source lists by MessageId).</summary>
    public readonly record struct ReadRow(string Rfq, string Supplier, string? MessageId, DateTimeOffset? MsgTime);

    /// <summary>Unread badge tally — total + per-RFQ + per "rfqId|supplier", matching the wire shape.</summary>
    public sealed record UnreadTally(int Total, Dictionary<string, int> ByRfq, Dictionary<string, int> BySupplier);

    /// <summary>A user's two explicit-override sets: messages they marked read, and messages they marked
    /// unread (the latter overrides the watermark so an OLD/pre-watermark message can be made unread).</summary>
    public readonly record struct ReadSets(HashSet<string> Read, HashSet<string> Unread);

    /// <summary>
    /// Whether a user is UNREAD on an inbound message. Explicit overrides win over the clean-slate
    /// watermark: in the unread-set ⇒ unread; in the read-set ⇒ read; otherwise the watermark default
    /// (arrived strictly AFTER it ⇒ unread). A message with no timestamp + no override is treated as read
    /// (old/placeholder data — never spuriously unread).
    /// </summary>
    public static bool IsUnread(DateTimeOffset? msgTime, string? messageId, DateTimeOffset adoptedAt,
        IReadOnlySet<string> readIds, IReadOnlySet<string> unreadIds)
    {
        if (messageId is { Length: > 0 } id)
        {
            if (unreadIds.Contains(id)) return true;    // explicit unread overrides the watermark
            if (readIds.Contains(id))   return false;   // explicit read overrides the watermark
        }
        return msgTime is { } t && t > adoptedAt;        // watermark default
    }

    /// <summary>Parse the stored note (newline-joined MessageIds) into a case-sensitive set.</summary>
    public static HashSet<string> ParseReadIds(string? note)
    {
        var set = new HashSet<string>(IdComparer);
        if (string.IsNullOrWhiteSpace(note)) return set;
        foreach (var line in note.Split('\n'))
        {
            var id = line.Trim();
            if (id.Length > 0) set.Add(id);
        }
        return set;
    }

    /// <summary>Serialize a read-set back to the stored note form (sorted Ordinal for stable diffs).</summary>
    public static string SerializeReadIds(IEnumerable<string> readIds)
        => string.Join('\n', readIds.Where(s => !string.IsNullOrWhiteSpace(s))
                                     .Distinct(IdComparer).OrderBy(s => s, IdComparer));

    /// <summary>
    /// Apply a mark to a read-set and return the NEW set. read=true adds the id (only when it's after the
    /// watermark — at/before the watermark is already covered, so storing it is redundant); read=false
    /// removes it. Also prunes any id at/under the watermark so the set never accumulates covered ids.
    /// </summary>
    public static ReadSets ApplyMark(IReadOnlySet<string> readIds, IReadOnlySet<string> unreadIds, string messageId, bool read)
    {
        var r = new HashSet<string>(readIds, IdComparer);
        var u = new HashSet<string>(unreadIds, IdComparer);
        if (!string.IsNullOrEmpty(messageId))
        {
            if (read) { r.Add(messageId); u.Remove(messageId); }   // explicit read  (overrides a watermark-unread)
            else      { u.Add(messageId); r.Remove(messageId); }   // explicit unread (overrides a watermark-read — old msgs)
        }
        return new ReadSets(r, u);
    }

    /// <summary>
    /// Count unread across inbound message rows for ONE user, deduped by MessageId (Ordinal) across the two
    /// source lists; null/empty-MessageId rows count individually (cannot dedup). Grouped by RFQ and by
    /// "rfqId|supplier" (keys compared OrdinalIgnoreCase, matching the existing badge cascade).
    /// </summary>
    public static UnreadTally Tally(IEnumerable<ReadRow> rows, DateTimeOffset adoptedAt,
        IReadOnlySet<string> readIds, IReadOnlySet<string> unreadIds)
    {
        var seen  = new HashSet<string>(IdComparer);
        var byRfq = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);
        var bySup = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);
        int total = 0;
        foreach (var r in rows)
        {
            if (r.MessageId is { Length: > 0 } id && !seen.Add(id)) continue;   // dedup across both lists
            if (!IsUnread(r.MsgTime, r.MessageId, adoptedAt, readIds, unreadIds)) continue;
            total++;
            byRfq[r.Rfq] = byRfq.GetValueOrDefault(r.Rfq) + 1;
            var sk = $"{r.Rfq}|{r.Supplier}";
            bySup[sk] = bySup.GetValueOrDefault(sk) + 1;
        }
        return new UnreadTally(total, byRfq, bySup);
    }
}
