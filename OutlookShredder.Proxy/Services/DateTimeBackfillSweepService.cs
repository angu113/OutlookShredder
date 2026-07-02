namespace OutlookShredder.Proxy.Services;

/// <summary>
/// Self-healing bridge for the text-date → native-dateTime migration (wip/datetime-column-migration.md).
/// While the FLEET still runs older proxies that don't dual-write the <c>*Dt</c> columns, any row THOSE
/// proxies create is text-only (null <c>*Dt</c>) and would sort/​filter wrong on this (updated) proxy —
/// worst case, rows read via a server-side <c>$filter=fields/XxxDt ge …</c> (e.g. the PhoneCallLog window)
/// are EXCLUDED entirely until healed.
///
/// This service periodically re-runs the idempotent backfills (they skip rows that already carry the
/// dateTime value), so drift an old proxy introduces is healed within one interval — no manual re-runs.
/// It's a rollout bridge: once the whole fleet is on the dual-write build and this logs 0 healed for a
/// while, disable it (<c>DateTimeBackfill:SweepEnabled=false</c>). Idempotent + read-cheap (a few paginated
/// list reads per cycle); redundant across proxies is harmless (first to patch wins, others skip).
/// </summary>
public sealed class DateTimeBackfillSweepService : BackgroundService
{
    private readonly SharePointService                        _sp;
    private readonly Storage.IInquiryStore                    _inquiries;
    private readonly IConfiguration                           _config;
    private readonly ILogger<DateTimeBackfillSweepService>    _log;

    public DateTimeBackfillSweepService(
        SharePointService sp, Storage.IInquiryStore inquiries,
        IConfiguration config, ILogger<DateTimeBackfillSweepService> log)
    {
        _sp = sp; _inquiries = inquiries; _config = config; _log = log;
    }

    protected override async Task ExecuteAsync(CancellationToken stoppingToken)
    {
        // Opt-out (default on): it's the migration-rollout mitigation. Turn off once the fleet is updated.
        if (!_config.GetValue("DateTimeBackfill:SweepEnabled", true))
        {
            _log.LogInformation("[DtSweep] disabled (DateTimeBackfill:SweepEnabled=false)");
            return;
        }
        var intervalMin = Math.Max(5, _config.GetValue("DateTimeBackfill:SweepIntervalMinutes", 30));
        _log.LogInformation("[DtSweep] enabled — healing text-date drift every {Min} min " +
            "(rows written by fleet proxies not yet on the dual-write build)", intervalMin);

        // Startup delay so the first sweep doesn't collide with SharePoint pre-warm.
        try { await Task.Delay(TimeSpan.FromMinutes(2), stoppingToken); } catch (OperationCanceledException) { return; }

        while (!stoppingToken.IsCancellationRequested)
        {
            try
            {
                int healed = 0;
                async Task Sweep(Task<(int Scanned, int Patched, int Failed)> backfill) => healed += (await backfill).Patched;

                // Both are registry-driven (one entry per migrated column), so a newly-migrated column is
                // swept automatically — nothing to update here.
                await Sweep(_sp.BackfillAllDateTimeColumnsAsync(stoppingToken));    // ErpDocuments x2, PhoneCallLog, Messages
                await Sweep(_inquiries.BackfillDateTimeColumnsAsync(stoppingToken)); // Inquiries / Drafts / Notes / Quotations

                if (healed > 0)
                    _log.LogInformation("[DtSweep] healed {N} drifted row(s) this cycle", healed);
                else
                    _log.LogDebug("[DtSweep] no drift — every row carries its native dateTime column");
            }
            catch (OperationCanceledException) { break; }
            catch (Exception ex) { _log.LogWarning(ex, "[DtSweep] sweep cycle failed — retrying next interval"); }

            try { await Task.Delay(TimeSpan.FromMinutes(intervalMin), stoppingToken); } catch (OperationCanceledException) { break; }
        }
    }
}
