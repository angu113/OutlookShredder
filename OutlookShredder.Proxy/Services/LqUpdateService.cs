namespace OutlookShredder.Proxy.Services;

/// <summary>
/// Background service that periodically recomputes LQ (average $/lb per Metal+Shape)
/// from recent supplier quotes and writes it to the QC SharePoint list.
///
/// Multiple proxy instances can run concurrently — all compute the same 7-day average
/// over the same data, so concurrent writes are idempotent. Random startup jitter
/// (up to half the configured interval) staggers the instances to avoid a stampede.
/// </summary>
public class LqUpdateService : BackgroundService
{
    private readonly SharePointService  _sp;
    private readonly IConfiguration     _config;
    private readonly ILogger<LqUpdateService> _log;

    public LqUpdateService(SharePointService sp, IConfiguration config, ILogger<LqUpdateService> log)
    {
        _sp     = sp;
        _config = config;
        _log    = log;
    }

    protected override async Task ExecuteAsync(CancellationToken ct)
    {
        var intervalMinutes = int.TryParse(_config["QC:LqUpdateIntervalMinutes"], out var m) ? m : 60;
        if (intervalMinutes <= 0) return;  // 0 disables the background updater

        var interval = TimeSpan.FromMinutes(intervalMinutes);

        // Random jitter: 0 to half the interval, so instances don't all fire together
        var jitter = TimeSpan.FromSeconds(Random.Shared.NextDouble() * interval.TotalSeconds / 2);
        _log.LogInformation("[LQ] Background updater starting — interval {Interval}min, startup jitter {Jitter:F0}s",
            intervalMinutes, jitter.TotalSeconds);

        await Task.Delay(jitter, ct);

        while (!ct.IsCancellationRequested)
        {
            try
            {
                _log.LogInformation("[LQ] Running scheduled LQ update");
                var result = await _sp.UpdateQcLqAsync();
                _log.LogInformation("[LQ] Scheduled update complete — {Updated} rows updated, {Misses} misses",
                    result.Updated.Count, result.Misses.Count);
            }
            catch (Exception ex) when (ex is not OperationCanceledException)
            {
                _log.LogError(ex, "[LQ] Scheduled update failed");
            }

            await Task.Delay(interval, ct);
        }
    }
}
