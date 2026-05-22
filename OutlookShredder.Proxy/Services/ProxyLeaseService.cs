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

    /// <summary>
    /// Unconditionally steals the lease from whoever holds it and sets IsLeaseHolder=true
    /// immediately so the caller can start the hook without waiting for the next renewal tick.
    /// Verifies the steal succeeded by reading back after 5 s (another proxy's concurrent
    /// renewal can overwrite us in a race) and re-steals up to two more times if needed.
    /// </summary>
    public async Task StealLeaseAsync(CancellationToken ct)
    {
        for (var attempt = 1; attempt <= 3; attempt++)
        {
            try
            {
                var prev = await _sp.ForceClaimLeaseAsync(ServiceName, _machine, LeaseSeconds, ct);
                _isLeaseHolder = true;
                if (attempt == 1)
                {
                    if (!string.IsNullOrEmpty(prev) && !string.Equals(prev, _machine, StringComparison.OrdinalIgnoreCase))
                        _log.LogInformation("[Lease:{Svc}] Stolen from {Prev} — verifying hold", ServiceName, prev);
                    else
                        _log.LogInformation("[Lease:{Svc}] Claimed on startup (uncontested) — verifying hold", ServiceName);
                }
                else
                {
                    _log.LogInformation("[Lease:{Svc}] Re-steal attempt {N} — verifying hold", ServiceName, attempt);
                }
            }
            catch (Exception ex)
            {
                _log.LogWarning(ex, "[Lease:{Svc}] Steal attempt {N} failed — will wait for normal acquisition", ServiceName, attempt);
                return;
            }

            // Wait 5 s, then renew to verify we still hold it.
            // A concurrent renewal from another proxy can overwrite our steal within milliseconds;
            // this read-back catches that and re-steals if needed.
            try { await Task.Delay(TimeSpan.FromSeconds(5), ct); }
            catch (OperationCanceledException) { return; }

            try
            {
                var held = await _sp.AcquireOrRenewLeaseAsync(ServiceName, _machine, LeaseSeconds, ct);
                if (held)
                {
                    _log.LogInformation("[Lease:{Svc}] Hold confirmed on {Machine} — hook can start", ServiceName, _machine);
                    return;
                }
                _log.LogInformation("[Lease:{Svc}] Steal attempt {N} overwritten by concurrent renewal — retrying", ServiceName, attempt);
            }
            catch (Exception ex)
            {
                _log.LogWarning(ex, "[Lease:{Svc}] Post-steal verification failed", ServiceName);
                return;
            }
        }

        _log.LogWarning("[Lease:{Svc}] Could not confirm hold after 3 steal attempts — will wait for normal acquisition", ServiceName);
        _isLeaseHolder = false;
    }

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
