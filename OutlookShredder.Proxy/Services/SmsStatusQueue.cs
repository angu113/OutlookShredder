using System.Text.Json;
using Azure.Messaging.ServiceBus;
using Azure.Messaging.ServiceBus.Administration;

namespace OutlookShredder.Proxy.Services;

/// <summary>
/// Owns the <c>sms-status-jobs</c> Service Bus QUEUE — carries SignalWire delivery-status callbacks
/// (queued/sent/delivered/failed) from the Azure Function ingress (<c>OutlookShredder.SmsWebhook</c>,
/// which has no SharePoint access) to a proxy, which owns the Messages-list write. Competing-consumers
/// (MaxConcurrentCalls=1) means exactly one proxy applies each status; duplicate detection collapses
/// re-deliveries. No leader, no lease. The processor (<see cref="SmsStatusQueueProcessor"/>) consumes;
/// this class provisions the queue (idempotent — any proxy instance can create it on first use, so the
/// Function doesn't need Service Bus manage rights). Mirrors <see cref="SmsInboundQueue"/>.
/// </summary>
public class SmsStatusQueue
{
    public sealed record Job(string Sid, string Status);

    private readonly ILogger<SmsStatusQueue> _log;

    private readonly ServiceBusClient?               _client;
    private readonly ServiceBusAdministrationClient? _admin;
    private readonly string _queueName;
    private int _ensured;   // 0 = not yet provisioned, 1 = provisioned

    public SmsStatusQueue(IConfiguration config, ILogger<SmsStatusQueue> log)
    {
        _log       = log;
        _queueName = config["ServiceBus:SmsStatusQueueName"] ?? "sms-status-jobs";

        var connStr = config["ServiceBus:ConnectionString"];
        if (!string.IsNullOrWhiteSpace(connStr))
        {
            try
            {
                _client = new ServiceBusClient(connStr);
                _admin  = new ServiceBusAdministrationClient(connStr);
            }
            catch (Exception ex) { _log.LogWarning(ex, "[SMS] status queue client init failed — status updates disabled"); }
        }
    }

    public bool Enabled => _client is not null;
    public string QueueName => _queueName;

    /// <summary>Idempotently provisions the queue with dedup + competing-consumer settings. Safe across proxies.</summary>
    public async Task EnsureQueueAsync(CancellationToken ct = default)
    {
        if (_admin is null) return;
        if (Interlocked.Exchange(ref _ensured, 1) == 1) return;
        try
        {
            if (!await _admin.QueueExistsAsync(_queueName, ct))
            {
                await _admin.CreateQueueAsync(new CreateQueueOptions(_queueName)
                {
                    RequiresDuplicateDetection = false,   // no stable dedup key (status can legitimately repeat: queued -> sent -> delivered, each a distinct event)
                    LockDuration               = TimeSpan.FromMinutes(5),
                    MaxDeliveryCount           = 5,
                    DefaultMessageTimeToLive   = TimeSpan.FromHours(6),
                    AutoDeleteOnIdle           = TimeSpan.MaxValue,   // persistent infra
                    RequiresSession            = false,
                }, ct);
                _log.LogInformation("[SMS] created queue '{Queue}'", _queueName);
            }
        }
        catch (ServiceBusException ex) when (ex.Reason == ServiceBusFailureReason.MessagingEntityAlreadyExists) { /* peer won the race */ }
        catch (Exception ex) { _log.LogWarning(ex, "[SMS] EnsureQueue failed"); Interlocked.Exchange(ref _ensured, 0); }
    }

    public ServiceBusProcessor? CreateProcessor()
        => _client?.CreateProcessor(_queueName, new ServiceBusProcessorOptions
        {
            MaxConcurrentCalls         = 1,                          // exactly-once across all proxies
            AutoCompleteMessages       = false,                      // complete only after processing
            MaxAutoLockRenewalDuration = TimeSpan.FromMinutes(5),
        });
}
