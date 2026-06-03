using OutlookShredder.Proxy.Models;

namespace OutlookShredder.Proxy.Services;

/// <summary>
/// Computes supplier-confirmation "at-risk" levels for purchase orders (Fulfillment loop -
/// see wip/fulfillment-loop.md). Pure + testable: given a PO's confirmation state, booked time,
/// configured thresholds, and "now", returns green | amber | red. A PO is only at-risk while
/// ConfirmStatus == "Pending". There are two clocks: the supplier-ACK clock (here) and the
/// pay-to-release clock (added with the payment track).
/// </summary>
public static class PoConfirmationMonitor
{
    /// <summary>Tuning, bound from the "PoMonitor" config section. All thresholds configurable.</summary>
    public sealed class Options
    {
        // Supplier acknowledgment clock (runs from PO booked / ReceivedAt until Confirmed).
        public int    AckAmberMinutes { get; set; } = 60;       // amber once unconfirmed this long
        public int    AckRedMinutes   { get; set; } = 120;      // red after this long...
        public string AckRedCutoffEst { get; set; } = "15:30";  // ...or once past this EST wall-clock
        public int    MonitorWindowBusinessDays { get; set; } = 3; // older Pending POs -> "stale", not red

        // Pay-to-release clock (our action; used by the payment track increment).
        public int    PayAmberMinutes { get; set; } = 30;
        public int    PayRedMinutes   { get; set; } = 45;
        public string PayRedCutoffEst { get; set; } = "15:15";
    }

    private static readonly TimeZoneInfo Eastern = ResolveEastern();

    private static TimeZoneInfo ResolveEastern()
    {
        foreach (var id in new[] { "Eastern Standard Time", "America/New_York" })
        {
            try { return TimeZoneInfo.FindSystemTimeZoneById(id); }
            catch { /* try next id */ }
        }
        return TimeZoneInfo.Utc;
    }

    /// <summary>Enriches each PO with its supplier-ack at-risk level (sets AckLevel /
    /// MinutesSincePlaced / AckCutoffPassed). Confirmed (non-Pending) POs come back "green".</summary>
    public static void EnrichAck(IEnumerable<PurchaseOrderRecord> pos, Options opts, DateTimeOffset utcNow)
    {
        foreach (var po in pos)
        {
            var (level, minutes, cutoff) = ComputeAck(po, opts, utcNow);
            po.AckLevel           = level;
            po.MinutesSincePlaced = minutes;
            po.AckCutoffPassed    = cutoff;
        }
    }

    /// <summary>(level, minutesSincePlaced, cutoffPassed) for the supplier-acknowledgment clock.
    /// level = green | amber | red. Only "Pending" POs are scored; others are "green".</summary>
    public static (string Level, int? Minutes, bool CutoffPassed) ComputeAck(
        PurchaseOrderRecord po, Options opts, DateTimeOffset utcNow)
    {
        if (!string.Equals(po.ConfirmStatus, "Pending", StringComparison.OrdinalIgnoreCase))
            return ("green", null, false);

        if (!DateTimeOffset.TryParse(po.ReceivedAt, out var booked))
            return ("green", null, false);   // no clock without a booked time

        var minutes = (int)Math.Max(0, (utcNow - booked).TotalMinutes);

        // Beyond the active window the PO is historical, not actionable -> "stale" (neutral).
        if (IsStale(booked, utcNow, opts.MonitorWindowBusinessDays))
            return ("stale", minutes, false);

        var elapsed = minutes >= opts.AckRedMinutes   ? 2
                    : minutes >= opts.AckAmberMinutes ? 1
                    : 0;

        var cutoffPassed = CutoffPassed(opts.AckRedCutoffEst, utcNow);
        var level = Math.Max(elapsed, cutoffPassed ? 2 : 0);

        return (LevelName(level), minutes, cutoffPassed);
    }

    /// <summary>True once the EST wall-clock is at/after the given HH:mm cutoff (same supplier-day
    /// urgency: past the cutoff, the order won't be processed today, so treat as red).</summary>
    public static bool CutoffPassed(string hhmmEst, DateTimeOffset utcNow)
    {
        if (!TryParseHhmm(hhmmEst, out var cutoff)) return false;
        var estNow = TimeZoneInfo.ConvertTime(utcNow, Eastern);
        return estNow.TimeOfDay >= cutoff;
    }

    /// <summary>True when more than <paramref name="windowBusinessDays"/> business days (Mon-Fri, EST)
    /// have elapsed since the PO was booked - i.e. it is historical, not actively monitored ("stale").</summary>
    public static bool IsStale(DateTimeOffset bookedUtc, DateTimeOffset utcNow, int windowBusinessDays)
    {
        var bookedEst = TimeZoneInfo.ConvertTime(bookedUtc, Eastern).Date;
        var nowEst    = TimeZoneInfo.ConvertTime(utcNow, Eastern).Date;
        // Quick reject: N business days span at most N+2 calendar days (a weekend in between).
        if ((nowEst - bookedEst).TotalDays > windowBusinessDays + 4) return true;
        var bd = 0;
        for (var d = bookedEst.AddDays(1); d <= nowEst; d = d.AddDays(1))
            if (d.DayOfWeek is not (DayOfWeek.Saturday or DayOfWeek.Sunday)) bd++;
        return bd > windowBusinessDays;
    }

    private static string LevelName(int level) => level >= 2 ? "red" : level == 1 ? "amber" : "green";

    private static bool TryParseHhmm(string? s, out TimeSpan t)
    {
        s = (s ?? "").Trim();
        return TimeSpan.TryParseExact(s, "hh\\:mm", System.Globalization.CultureInfo.InvariantCulture, out t)
            || TimeSpan.TryParse(s, System.Globalization.CultureInfo.InvariantCulture, out t);
    }
}
