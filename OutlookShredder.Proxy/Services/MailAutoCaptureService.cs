namespace OutlookShredder.Proxy.Services;

/// <summary>
/// Phase 1b-iii: periodically captures + classifies new mirror messages so the workbench fills
/// automatically (instead of the manual Seed/Backfill). Cross-proxy-safe via the claim/processed
/// categories on the wrapper message. Config:
///   MailWorkbench:AutoCapture        — default true
///   MailWorkbench:AutoCaptureSeconds — default 60
/// </summary>
public sealed class MailAutoCaptureService : BackgroundService
{
    private readonly MailWorkbenchService _wb;
    private readonly MailboxBridgeService _bridge;
    private readonly IConfiguration _config;
    private readonly ILogger<MailAutoCaptureService> _log;

    public MailAutoCaptureService(MailWorkbenchService wb, MailboxBridgeService bridge,
        IConfiguration config, ILogger<MailAutoCaptureService> log)
    {
        _wb = wb; _bridge = bridge; _config = config; _log = log;
    }

    protected override async Task ExecuteAsync(CancellationToken ct)
    {
        var enabled = !bool.TryParse(_config["MailWorkbench:AutoCapture"], out var e) || e;   // default ON
        if (!enabled) { _log.LogInformation("[MailAutoCapture] disabled (MailWorkbench:AutoCapture=false)"); return; }
        if (_bridge.MailboxCount == 0) { _log.LogInformation("[MailAutoCapture] no mailboxes configured — idle"); return; }

        var interval = int.TryParse(_config["MailWorkbench:AutoCaptureSeconds"], out var s) && s > 0 ? s : 60;
        _log.LogInformation("[MailAutoCapture] enabled — capturing new mail every {Secs}s", interval);

        try { await Task.Delay(TimeSpan.FromSeconds(20), ct); } catch (OperationCanceledException) { return; }  // let the bridge seed first

        while (!ct.IsCancellationRequested)
        {
            try
            {
                foreach (var mb in _bridge.GetStatuses())
                    await _wb.AutoCaptureCycleAsync(mb.WatchedUpn, ct);
            }
            catch (OperationCanceledException) when (ct.IsCancellationRequested) { break; }
            catch (Exception ex) { _log.LogWarning(ex, "[MailAutoCapture] cycle error"); }

            try { await Task.Delay(TimeSpan.FromSeconds(interval), ct); } catch (OperationCanceledException) { break; }
        }
    }
}
