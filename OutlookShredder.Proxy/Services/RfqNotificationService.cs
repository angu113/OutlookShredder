using System.Collections.Concurrent;
using System.Text.Json;
using System.Threading.Channels;
using Azure.Messaging.ServiceBus;
using OutlookShredder.Proxy.Models;

namespace OutlookShredder.Proxy.Services;

/// <summary>
/// Singleton pub/sub bus for Server-Sent Events and Azure Service Bus.
/// When an RFQ row is successfully written to SharePoint the poller (or the
/// add-in taskpane via the extract endpoint) calls <see cref="NotifyRfqProcessed"/>,
/// which fans the event out to every connected SSE client and to the Service Bus
/// topic so that Shredder instances on other machines are also notified.
/// The SSE channel carries "{eventName}\n{dataJson}" strings so clients can parse both fields.
/// </summary>
public class RfqNotificationService
{
    private readonly ConcurrentDictionary<Guid, Channel<string>> _subscribers = new();
    private readonly ServiceBusSender?                           _sbSender;
    private readonly ILogger<RfqNotificationService>             _log;

    public RfqNotificationService(IConfiguration config, ILogger<RfqNotificationService> log)
    {
        _log = log;
        var connStr   = config["ServiceBus:ConnectionString"];
        var topicName = config["ServiceBus:TopicName"] ?? "rfq-updates";

        if (!string.IsNullOrWhiteSpace(connStr))
        {
            try
            {
                var client = new ServiceBusClient(connStr);
                _sbSender  = client.CreateSender(topicName);
                _log.LogInformation("[ServiceBus] Publisher connected to topic '{Topic}'", topicName);
            }
            catch (Exception ex)
            {
                _log.LogWarning(ex, "[ServiceBus] Failed to create sender — cross-machine notifications disabled");
            }
        }
        else
        {
            _log.LogInformation("[ServiceBus] ConnectionString not configured — cross-machine notifications disabled");
        }
    }

    // ── SSE subscription management ───────────────────────────────────────────

    /// <summary>
    /// Registers a new SSE subscriber.  Returns the subscriber ID (needed for
    /// <see cref="Unsubscribe"/>) and the <see cref="ChannelReader{T}"/> to drain.
    /// </summary>
    public (Guid Id, ChannelReader<string> Reader) Subscribe()
    {
        var id = Guid.NewGuid();
        var ch = Channel.CreateUnbounded<string>(
            new UnboundedChannelOptions { SingleWriter = false, SingleReader = true });
        _subscribers[id] = ch;
        return (id, ch.Reader);
    }

    /// <summary>Removes the subscriber and completes its channel.</summary>
    public void Unsubscribe(Guid id)
    {
        if (_subscribers.TryRemove(id, out var ch))
            ch.Writer.TryComplete();
    }

    /// <summary>Broadcasts a minimal <c>rfq-processed</c> event with no payload data.</summary>
    public void NotifyRfqProcessed() => NotifyRfqProcessed(new RfqProcessedNotification());

    /// <summary>
    /// Publishes a "Synonym" event to Service Bus so all Shredder clients update
    /// their local synonym caches without restarting.
    /// </summary>
    public void NotifySynonym(Models.SynonymGroup group) =>
        NotifyRfqProcessed(new RfqProcessedNotification
        {
            EventType    = "Synonym",
            SynonymGroup = group,
        });

    /// <summary>
    /// Broadcasts an <c>rfq-processed</c> SSE event carrying supplier + product data
    /// to all connected SSE clients, and publishes the same payload to Azure Service Bus
    /// so Shredder instances on other machines are notified.
    /// Safe to call from any thread.
    /// </summary>
    public void NotifyRfqProcessed(RfqProcessedNotification notification)
    {
        var json = JsonSerializer.Serialize(notification,
            new JsonSerializerOptions { PropertyNamingPolicy = JsonNamingPolicy.CamelCase });

        // SSE — local clients
        var msg = $"rfq-processed\n{json}";
        foreach (var ch in _subscribers.Values)
            ch.Writer.TryWrite(msg);

        // Service Bus — cross-machine clients (fire and forget)
        if (_sbSender is not null)
            _ = PublishToServiceBusAsync(json);
    }

    private async Task PublishToServiceBusAsync(string json)
    {
        try
        {
            var message = new ServiceBusMessage(json) { ContentType = "application/json" };
            await _sbSender!.SendMessageAsync(message);
        }
        catch (Exception ex)
        {
            _log.LogWarning(ex, "[ServiceBus] Failed to publish rfq-processed message");
        }
    }
}
