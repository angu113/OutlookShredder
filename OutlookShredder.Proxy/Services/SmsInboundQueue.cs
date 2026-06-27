using System.Text.Json;
using Azure.Messaging.ServiceBus;
using Azure.Messaging.ServiceBus.Administration;

namespace OutlookShredder.Proxy.Services;

/// <summary>
/// Owns the <c>sms-inbound-jobs</c> Service Bus QUEUE — exactly-once handling of inbound SMS across all
/// proxies. Any proxy that RECEIVES a SignalWire inbound webhook (Cloudflare routes to any healthy tunnel
/// replica) enqueues the raw message; the queue's <b>duplicate detection</b> (MessageId = the SignalWire
/// MessageSid) collapses re-deliveries, and <b>competing-consumers</b> (MaxConcurrentCalls=1) means
/// exactly ONE proxy processes it, with auto-failover via lock expiry. No leader, no lease, no held
/// connection. The processor (<see cref="SmsInboundQueueProcessor"/>) consumes; this class provisions +
/// enqueues. Mirrors <see cref="RfqSummaryQueue"/>.
/// </summary>
public class SmsInboundQueue
{
    public sealed record Job(string From, string To, string Body, string? Sid, string? MediaUrls);

    private readonly IConfiguration            _config;
    private readonly ILogger<SmsInboundQueue>  _log;

    private readonly ServiceBusClient?               _client;
    private readonly ServiceBusSender?               _sender;
    private readonly ServiceBusAdministrationClient? _admin;
    private readonly string _queueName;
    private int _ensured;   // 0 = not yet provisioned, 1 = provisioned

    private static readonly JsonSerializerOptions _json = new() { PropertyNamingPolicy = JsonNamingPolicy.CamelCase };

    public SmsInboundQueue(IConfiguration config, ILogger<SmsInboundQueue> log)
    {
        _config    = config;
        _log       = log;
        _queueName = config["ServiceBus:SmsInboundQueueName"] ?? "sms-inbound-jobs";

        var connStr = config["ServiceBus:ConnectionString"];
        if (!string.IsNullOrWhiteSpace(connStr))
        {
            try
            {
                _client = new ServiceBusClient(connStr);
                _sender = _client.CreateSender(_queueName);
                _admin  = new ServiceBusAdministrationClient(connStr);
            }
            catch (Exception ex) { _log.LogWarning(ex, "[SMS] inbound queue client init failed — inbound SMS disabled"); }
        }
    }

    public bool Enabled => _sender is not null;
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
                    RequiresDuplicateDetection          = true,
                    DuplicateDetectionHistoryTimeWindow = TimeSpan.FromMinutes(10),
                    LockDuration                        = TimeSpan.FromMinutes(5),
                    MaxDeliveryCount                    = 5,
                    DefaultMessageTimeToLive            = TimeSpan.FromHours(6),
                    AutoDeleteOnIdle                    = TimeSpan.MaxValue,   // persistent infra
                    RequiresSession                     = false,
                }, ct);
                _log.LogInformation("[SMS] created queue '{Queue}'", _queueName);
            }
        }
        catch (ServiceBusException ex) when (ex.Reason == ServiceBusFailureReason.MessagingEntityAlreadyExists) { /* peer won the race */ }
        catch (Exception ex) { _log.LogWarning(ex, "[SMS] EnsureQueue failed"); Interlocked.Exchange(ref _ensured, 0); }
    }

    /// <summary>Enqueue one inbound SMS for exactly-once processing. MessageId = SignalWire MessageSid so
    /// a re-delivered webhook (or a second proxy receiving the same message) collapses to one job.</summary>
    public async Task EnqueueAsync(Job job, CancellationToken ct = default)
    {
        if (_sender is null) { _log.LogWarning("[SMS] inbound queue disabled — dropping inbound from {From}", job.From); return; }
        await EnsureQueueAsync(ct);
        var msg = new ServiceBusMessage(JsonSerializer.Serialize(job, _json)) { ContentType = "application/json" };
        if (!string.IsNullOrWhiteSpace(job.Sid)) msg.MessageId = job.Sid;   // duplicate-detection key
        await _sender.SendMessageAsync(msg, ct);
    }

    public ServiceBusProcessor? CreateProcessor()
        => _client?.CreateProcessor(_queueName, new ServiceBusProcessorOptions
        {
            MaxConcurrentCalls         = 1,                          // exactly-once across all proxies
            AutoCompleteMessages       = false,                      // complete only after processing
            MaxAutoLockRenewalDuration = TimeSpan.FromMinutes(5),
        });
}
