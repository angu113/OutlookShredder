namespace OutlookShredder.Proxy.Services;

/// <summary>
/// Distributed single-instance coordinator via SharePoint lease.
/// Ensures that exactly one proxy across all machines holds each named lease at a time.
/// Other services (e.g. ZoomCallWatcherService) query IsLeaseHolder before starting.
///
/// Lease duration: 60 s.  Renewal interval: 30 s.
/// If the holder crashes the lease expires in ≤60 s and another proxy claims it.
/// </summary>
public class ProxyLeaseService : BackgroundService
{
    public const  string ServiceName  = "Zoom";
    private const int    LeaseSeconds = 60;
    private const int    RenewEvery   = 30;

    private readonly SharePointService          _sp;
    private readonly ILogger<ProxyLeaseService> _log;
    private readonly string                     _machine = Environment.MachineName;

    private volatile bool _isLeaseHolder;

    /// <summary>True when this proxy currently holds the Zoom lease in SharePoint.</summary>
    public bool IsLeaseHolder => _isLeaseHolder;

    public ProxyLeaseService(SharePointService sp, ILogger<ProxyLeaseService> log)
    {
        _sp  = sp;
        _log = log;
    }

    protected override async Task ExecuteAsync(CancellationToken ct)
    {
        while (!ct.IsCancellationRequested)
        {
            try
            {
                var held = await _sp.AcquireOrRenewLeaseAsync(ServiceName, _machine, LeaseSeconds, ct);
                if (held != _isLeaseHolder)
                {
                    _isLeaseHolder = held;
                    _log.LogInformation("[Lease:{Svc}] {State} on {Machine}",
                        ServiceName, held ? "acquired" : "released", _machine);
                }
            }
            catch (OperationCanceledException) { break; }
            catch (Exception ex)
            {
                _log.LogWarning(ex, "[Lease:{Svc}] Check failed — retaining last known state", ServiceName);
            }

            try { await Task.Delay(TimeSpan.FromSeconds(RenewEvery), ct); }
            catch (OperationCanceledException) { break; }
        }

        _isLeaseHolder = false;
    }
}
