using System.Collections.Concurrent;
using System.Diagnostics;

namespace OutlookShredder.Proxy.Services;

/// <summary>
/// Process-wide recorder for startup / cache-warm timings, with per-DATASET granularity ('the cache' is
/// many independent datasets). Each warm step records a span — either by wrapping awaited work in
/// <see cref="MeasureAsync{T}"/> (sets row count automatically) or a synchronous <see cref="Track"/>
/// scope. Every span is logged as <c>[Startup] step/dataset Nms rows=R</c> and kept for
/// <c>GET /api/startup</c>, so we can see exactly which dataset is slow and whether it sits on the
/// critical path (blocks Kestrel from listening — i.e. the window where the UI shows the proxy "down").
///
/// Only the FIRST occurrence of each (step, dataset) is recorded — that's the startup warm; the periodic
/// 5-min refreshes that reuse the same methods are ignored, so the snapshot stays a clean startup profile.
/// </summary>
public static class StartupTimings
{
    public static readonly DateTime ProcessStartUtc = DateTime.UtcNow;
    private static readonly long _swStart = Stopwatch.GetTimestamp();
    private static long? _listeningTs;          // Stopwatch ts when Kestrel began serving (ApplicationStarted)
    private static readonly ConcurrentBag<Entry> _entries = new();
    private static readonly HashSet<string> _seen = new();
    private static readonly object _seenLock = new();
    private static Serilog.ILogger? _log;

    public static void UseLogger(Serilog.ILogger log) => _log = log;

    /// <summary>Milliseconds since process start (monotonic).</summary>
    public static double ElapsedMs => Stopwatch.GetElapsedTime(_swStart).TotalMilliseconds;

    /// <summary>Stamp the moment Kestrel began accepting connections (ApplicationStarted). This is the key
    /// "time to serve" number — until now every UI health/version poll fails and the buttons show red.</summary>
    public static void MarkListening()
    {
        _listeningTs ??= Stopwatch.GetTimestamp();
        _log?.Information("[Startup] Kestrel listening at {Ms:0}ms after process start", ElapsedMs);
    }

    /// <summary>Times an awaited warm and records it (row count via <paramref name="rows"/>). Use for the
    /// SP-read cache warms so each dataset gets its own line.</summary>
    public static async Task<T> MeasureAsync<T>(string step, string? dataset, Func<Task<T>> work,
        Func<T, int?>? rows = null, bool critical = false)
    {
        double startMs = ElapsedMs;
        long t0 = Stopwatch.GetTimestamp();
        var result = await work();
        Record(step, dataset, Stopwatch.GetElapsedTime(t0).TotalMilliseconds, rows?.Invoke(result), critical, startMs);
        return result;
    }

    /// <summary>Times an awaited warm with no meaningful row count.</summary>
    public static Task MeasureAsync(string step, string? dataset, Func<Task> work, bool critical = false)
        => MeasureAsync<bool>(step, dataset, async () => { await work(); return true; }, _ => null, critical);

    /// <summary>Synchronous scope: <c>using var t = StartupTimings.Track("step"); … t.Rows = n;</c>.</summary>
    public static Scope Track(string step, string? dataset = null, bool critical = false)
        => new(step, dataset, critical);

    private static void Record(string step, string? dataset, double ms, int? rows, bool critical, double startMs)
    {
        string key = step + "" + (dataset ?? "");
        lock (_seenLock) { if (!_seen.Add(key)) return; }   // first (startup) occurrence only
        _entries.Add(new Entry(step, dataset, ms, rows, critical, startMs));
        _log?.Information("[Startup] {Step}{Ds} {Ms:0}ms{Rows}{Crit}",
            step, dataset is null ? "" : "/" + dataset, ms,
            rows is int r ? $" rows={r}" : "", critical ? " [critical-path]" : "");
    }

    /// <summary>Snapshot for <c>GET /api/startup</c> — ordered by start offset so it reads as a timeline.</summary>
    public static object Snapshot()
    {
        double? listeningMs = _listeningTs is long t
            ? Math.Round(Stopwatch.GetElapsedTime(_swStart, t).TotalMilliseconds) : null;
        var rows = _entries.OrderBy(e => e.StartMs).ToList();
        double criticalTotal = rows.Where(e => e.Critical).Sum(e => e.Ms);
        return new
        {
            processStartUtc = ProcessStartUtc,
            nowMs           = Math.Round(ElapsedMs),
            listeningMs,                                   // time-to-serve (UI stops showing red here)
            criticalPathMs  = Math.Round(criticalTotal),  // sum of blocking warms
            count           = rows.Count,
            entries = rows.Select(e => new
            {
                e.Step, e.Dataset, ms = Math.Round(e.Ms), e.Rows, e.Critical,
                startMs = Math.Round(e.StartMs), endMs = Math.Round(e.StartMs + e.Ms),
            }),
        };
    }

    public sealed record Entry(string Step, string? Dataset, double Ms, int? Rows, bool Critical, double StartMs);

    public sealed class Scope : IDisposable
    {
        private readonly string _step;
        private readonly string? _dataset;
        private readonly bool _critical;
        private readonly double _startMs;
        private readonly long _t0;
        public int? Rows { get; set; }

        internal Scope(string step, string? dataset, bool critical)
        {
            _step = step; _dataset = dataset; _critical = critical;
            _startMs = ElapsedMs; _t0 = Stopwatch.GetTimestamp();
        }

        public void Dispose() => Record(_step, _dataset, Stopwatch.GetElapsedTime(_t0).TotalMilliseconds, Rows, _critical, _startMs);
    }
}
