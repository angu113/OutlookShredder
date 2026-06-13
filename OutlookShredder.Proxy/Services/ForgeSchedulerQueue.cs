using System.Text.Json;
using Azure.Messaging.ServiceBus;
using Azure.Messaging.ServiceBus.Administration;
using OutlookShredder.Proxy.Models;

namespace OutlookShredder.Proxy.Services;

/// <summary>
/// Owns the <c>forge-task-scheduler</c> Service Bus QUEUE.
/// Any proxy can enqueue a task; duplicate detection (MessageId = "{taskName}:{yyyyMMdd}")
/// collapses concurrent enqueues from multiple proxies into one message per task per day.
/// Competing-consumers ensures exactly ONE proxy ever executes a given scheduled run.
/// </summary>
public class ForgeSchedulerQueue
{
    private readonly IConfiguration              _config;
    private readonly ILogger<ForgeSchedulerQueue> _log;

    private readonly ServiceBusClient?               _client;
    private readonly ServiceBusSender?               _sender;
    private readonly ServiceBusAdministrationClient? _admin;
    private readonly string _queueName;
    private int _ensured;

    private static readonly JsonSerializerOptions _json =
        new() { PropertyNamingPolicy = JsonNamingPolicy.CamelCase };

    public ForgeSchedulerQueue(IConfiguration config, ILogger<ForgeSchedulerQueue> log)
    {
        _config    = config;
        _log       = log;
        _queueName = config["ForgeScheduler:TaskQueueName"] ?? "forge-task-scheduler";

        var connStr = config["ServiceBus:ConnectionString"];
        if (!string.IsNullOrWhiteSpace(connStr))
        {
            try
            {
                _client = new ServiceBusClient(connStr);
                _sender = _client.CreateSender(_queueName);
                _admin  = new ServiceBusAdministrationClient(connStr);
            }
            catch (Exception ex)
            {
                _log.LogWarning(ex, "[ForgeScheduler] Queue client init failed — scheduled tasks disabled");
            }
        }
    }

    public bool   Enabled   => _sender is not null;
    public string QueueName => _queueName;

    /// <summary>
    /// Idempotently provisions the queue.  Dup-detection window &gt; 24 h collapses
    /// all enqueues for a given task on the same calendar day to one message.
    /// Safe to call concurrently from many proxies — the peer-wins race is caught.
    /// </summary>
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
                    DuplicateDetectionHistoryTimeWindow = TimeSpan.FromHours(25), // one per calendar day
                    LockDuration                        = TimeSpan.FromMinutes(10),
                    MaxDeliveryCount                    = 2,
                    DefaultMessageTimeToLive            = TimeSpan.FromHours(3),
                    AutoDeleteOnIdle                    = TimeSpan.MaxValue,
                    RequiresSession                     = false,
                }, ct);
                _log.LogInformation("[ForgeScheduler] Created queue '{Queue}'", _queueName);
            }
        }
        catch (ServiceBusException ex) when (ex.Reason == ServiceBusFailureReason.MessagingEntityAlreadyExists)
        {
            /* peer won the race */
        }
        catch (Exception ex)
        {
            _log.LogWarning(ex, "[ForgeScheduler] EnsureQueue failed");
            Interlocked.Exchange(ref _ensured, 0);
        }
    }

    /// <summary>Enqueues a task for the competing-consumer processor.  Duplicate detection
    /// collapses multiple enqueues for the same task on the same UTC day to one message.</summary>
    public async Task EnqueueAsync(string taskName, CancellationToken ct = default)
    {
        if (!Enabled) return;
        try
        {
            await EnsureQueueAsync(ct);
            var body = JsonSerializer.Serialize(new ForgeTaskQueueMessage(taskName), _json);
            await _sender!.SendMessageAsync(new ServiceBusMessage(body)
            {
                MessageId   = $"{taskName}:{DateTime.UtcNow:yyyyMMdd}",
                ContentType = "application/json",
            }, ct);
            _log.LogInformation("[ForgeScheduler] Enqueued task '{Task}'", taskName);
        }
        catch (Exception ex) { _log.LogWarning(ex, "[ForgeScheduler] Enqueue failed for '{Task}'", taskName); }
    }

    public ServiceBusProcessor? CreateProcessor()
        => _client?.CreateProcessor(_queueName, new ServiceBusProcessorOptions
        {
            MaxConcurrentCalls         = 1,
            AutoCompleteMessages       = false,
            MaxAutoLockRenewalDuration = TimeSpan.FromMinutes(15),
        });
}
