using System.Text.Json;
using Azure.Messaging.ServiceBus;
using Azure.Messaging.ServiceBus.Administration;

namespace OutlookShredder.Proxy.Services;

/// <summary>
/// Owns the <c>rfq-summary-jobs</c> Service Bus QUEUE — the race-free generation mechanism for the RFQ
/// state-of-play summary. Any proxy that writes a price-impactful supplier response enqueues a job; the
/// queue's <b>duplicate detection</b> (MessageId = "{rfqId}:{inputsHash}") collapses identical jobs from
/// concurrent proxies, and <b>competing-consumers</b> means exactly ONE proxy ever generates a given
/// summary. The processor (<see cref="RfqSummaryQueueProcessor"/>) consumes; this class provisions the
/// queue + enqueues.
/// </summary>
public class RfqSummaryQueue
{
    public sealed record Job(string RfqId, string InputsHash);

    private readonly RfqStateOfPlayService _state;
    private readonly IConfiguration        _config;
    private readonly ILogger<RfqSummaryQueue> _log;

    private readonly ServiceBusClient?               _client;
    private readonly ServiceBusSender?               _sender;
    private readonly ServiceBusAdministrationClient? _admin;
    private readonly string _queueName;
    private int _ensured;   // 0 = not yet provisioned, 1 = provisioned

    private static readonly JsonSerializerOptions _json = new() { PropertyNamingPolicy = JsonNamingPolicy.CamelCase };

    public RfqSummaryQueue(RfqStateOfPlayService state, IConfiguration config, ILogger<RfqSummaryQueue> log)
    {
        _state = state; _config = config; _log = log;
        _queueName = config["ServiceBus:SummaryQueueName"] ?? "rfq-summary-jobs";

        var connStr = config["ServiceBus:ConnectionString"];
        if (!string.IsNullOrWhiteSpace(connStr))
        {
            try
            {
                _client = new ServiceBusClient(connStr);
                _sender = _client.CreateSender(_queueName);
                _admin  = new ServiceBusAdministrationClient(connStr);
            }
            catch (Exception ex) { _log.LogWarning(ex, "[StateOfPlay] queue client init failed — summary generation disabled"); }
        }
    }

    public bool Enabled => _sender is not null && _config.GetValue("RfqStateOfPlay:Enabled", true);
    public string QueueName => _queueName;

    /// <summary>Idempotently provisions the queue with the race-free settings (dup-detection 10 min,
    /// 5-min lock, 3 deliveries, 1-hour TTL, never auto-delete). Safe to call from many proxies.</summary>
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
                    MaxDeliveryCount                    = 3,
                    DefaultMessageTimeToLive            = TimeSpan.FromHours(1),
                    AutoDeleteOnIdle                    = TimeSpan.MaxValue,   // persistent infra
                    RequiresSession                     = false,
                }, ct);
                _log.LogInformation("[StateOfPlay] created queue '{Queue}'", _queueName);
            }
        }
        catch (Azure.Messaging.ServiceBus.ServiceBusException ex)
            when (ex.Reason == ServiceBusFailureReason.MessagingEntityAlreadyExists) { /* peer won the race */ }
        catch (Exception ex) { _log.LogWarning(ex, "[StateOfPlay] EnsureQueue failed"); Interlocked.Exchange(ref _ensured, 0); }
    }

    /// <summary>Best-effort enqueue of a state-of-play regeneration for an RFQ, if the new rows are
    /// price-impactful. Dedup collapses concurrent identical jobs; the processor gates on ≥2 suppliers.</summary>
    public async Task MaybeEnqueueAsync(string? rfqId, List<Dictionary<string, object?>> rows)
    {
        if (!Enabled || string.IsNullOrWhiteSpace(rfqId)) return;
        if (rfqId is "000000" or "WHOIS" || rfqId.StartsWith("WHOIS", StringComparison.OrdinalIgnoreCase)) return;
        if (!RfqStateOfPlayService.AnyImpactful(rows)) return;   // OOF / no-price → nothing to compare

        try
        {
            await EnsureQueueAsync();
            bool pdfs = _config.GetValue("RfqStateOfPlay:IncludePdfs", false);
            var hash  = _state.ComputeInputsHash(rows, pdfs);
            var body  = JsonSerializer.Serialize(new Job(rfqId, hash), _json);
            await _sender!.SendMessageAsync(new ServiceBusMessage(body)
            {
                MessageId   = $"{rfqId}:{hash}",   // duplicate-detection key — collapses concurrent proxies
                ContentType = "application/json",
            });
        }
        catch (Exception ex) { _log.LogWarning(ex, "[StateOfPlay] enqueue failed for {Rfq}", rfqId); }
    }

    public ServiceBusProcessor? CreateProcessor()
        => _client?.CreateProcessor(_queueName, new ServiceBusProcessorOptions
        {
            MaxConcurrentCalls          = 1,                          // generation is heavy — one at a time per box
            AutoCompleteMessages        = false,                      // complete only after the cache write
            MaxAutoLockRenewalDuration  = TimeSpan.FromMinutes(10),   // a slow AI call never loses the lock
        });
}
