using System.Globalization;

namespace OutlookShredder.Proxy.Services;

/// <summary>
/// Hourly background job that marks stale RFQ References as complete. Any
/// reference whose creation date is older than <c>DaysThreshold</c> and that
/// is not already Complete gets patched with Complete=true, and a sentinel
/// "[auto-completed after Nd]" line is appended to Notes so the provenance
/// of the state change is visible to humans browsing the list.
///
/// Configuration (<c>appsettings.json</c>):
///   RfqAutoComplete:Enabled        bool    default true
///   RfqAutoComplete:IntervalHours  double  default 1
///   RfqAutoComplete:DaysThreshold  int     default 7
/// </summary>
public sealed class RfqAutoCompleteService : BackgroundService
{
    private readonly SharePointService              _sp;
    private readonly IConfiguration                 _config;
    private readonly ILogger<RfqAutoCompleteService> _log;

    public RfqAutoCompleteService(
        SharePointService                 sp,
        IConfiguration                    config,
        ILogger<RfqAutoCompleteService>   log)
    {
        _sp     = sp;
        _config = config;
        _log    = log;
    }

    protected override async Task ExecuteAsync(CancellationToken ct)
    {
        // Stagger startup so we don't race the SharePoint pre-warm.
        try { await Task.Delay(TimeSpan.FromMinutes(2), ct); }
        catch (OperationCanceledException) { return; }

        while (!ct.IsCancellationRequested)
        {
            try
            {
                if (Enabled)
                    await RunCycleAsync(ct);
                else
                    _log.LogDebug("[AutoComplete] Disabled — skipping cycle");
            }
            catch (Exception ex)
            {
                _log.LogError(ex, "[AutoComplete] Cycle failed");
            }

            try { await Task.Delay(TimeSpan.FromHours(IntervalHours), ct); }
            catch (OperationCanceledException) { break; }
        }
    }

    private async Task RunCycleAsync(CancellationToken ct)
    {
        var days   = DaysThreshold;
        var cutoff = DateTime.UtcNow.AddDays(-days);

        var refs = await _sp.ReadRfqReferencesAsync();
        var scanned   = refs.Count;
        var completed = 0;

        foreach (var r in refs)
        {
            if (ct.IsCancellationRequested) break;

            var rfqId  = r.TryGetValue("RFQ_ID", out var idv) ? idv?.ToString() : null;
            var itemId = r.TryGetValue("Id",     out var iiv) ? iiv?.ToString() : null;
            if (string.IsNullOrEmpty(rfqId) || string.IsNullOrEmpty(itemId)) continue;

            if (IsTrue(r.TryGetValue("Complete", out var cv) ? cv : null)) continue;

            var created = ResolveCreatedUtc(r);
            if (created is null || created > cutoff) continue;

            var existingNotes = r.TryGetValue("Notes", out var nv) ? nv?.ToString() ?? "" : "";
            var suffix        = $"[auto-completed after {days}d]";

            // Skip if the suffix is already present — makes the job idempotent even
            // if Complete was toggled back off for some reason.
            if (existingNotes.Contains(suffix, StringComparison.Ordinal)) continue;

            var newNotes = string.IsNullOrWhiteSpace(existingNotes)
                ? suffix
                : existingNotes.TrimEnd() + "\n" + suffix;

            try
            {
                await _sp.PatchRfqReferenceByItemIdAsync(itemId!, new Dictionary<string, object?>
                {
                    ["Complete"] = true,
                    ["Notes"]    = newNotes,
                });
                completed++;
                _log.LogInformation(
                    "[AutoComplete] Marked {RfqId} complete (age: {Age:F1}d, threshold: {Days}d)",
                    rfqId, (DateTime.UtcNow - created.Value).TotalDays, days);
            }
            catch (Exception ex)
            {
                _log.LogWarning(ex, "[AutoComplete] Failed to complete {RfqId}", rfqId);
            }
        }

        _log.LogInformation(
            "[AutoComplete] Cycle finished — scanned {Scanned}, auto-completed {Completed} (threshold {Days}d)",
            scanned, completed, days);
    }

    // ── Config ────────────────────────────────────────────────────────────────

    private bool   Enabled        => _config.GetValue("RfqAutoComplete:Enabled", true);
    private double IntervalHours  => _config.GetValue("RfqAutoComplete:IntervalHours", 1.0);
    private int    DaysThreshold  => _config.GetValue("RfqAutoComplete:DaysThreshold", 7);

    // ── Helpers ───────────────────────────────────────────────────────────────

    private static bool IsTrue(object? v) => v switch
    {
        bool b                                                            => b,
        string s when bool.TryParse(s, out var parsed)                    => parsed,
        _                                                                 => false,
    };

    private static DateTime? ResolveCreatedUtc(Dictionary<string, object?> r)
    {
        // Prefer the custom DateCreated column (user-specified send date); fall back
        // to the SharePoint item-level Created timestamp.
        var raw = r.TryGetValue("DateCreated", out var dcv) ? dcv?.ToString() :
                  r.TryGetValue("Created",     out var cv)  ? cv?.ToString()  : null;
        if (string.IsNullOrEmpty(raw)) return null;

        if (DateTime.TryParse(raw, CultureInfo.InvariantCulture,
                DateTimeStyles.RoundtripKind | DateTimeStyles.AssumeUniversal, out var dt))
            return dt.ToUniversalTime();

        return null;
    }
}
