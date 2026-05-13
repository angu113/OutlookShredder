using System.Collections.Concurrent;
using System.Text.Json;
using System.Threading.Channels;
using Azure.Messaging.ServiceBus;
using Azure.Messaging.ServiceBus.Administration;
using OutlookShredder.Proxy.Models;

namespace OutlookShredder.Proxy.Services;

/// <summary>
/// Singleton pub/sub bus for Server-Sent Events and Azure Service Bus.
///
/// When an RFQ row is successfully written to SharePoint the poller (or the
/// add-in taskpane via the extract endpoint) calls <see cref="NotifyRfqProcessed"/>,
/// which fans the event out to every connected SSE client and to the Service Bus
/// topic so that Shredder instances on other machines are also notified.
///
/// The proxy also SUBSCRIBES to the same topic via <see cref="StartBusListenerAsync"/>
/// so that when a peer proxy on another machine writes new rows, this proxy can keep
/// its SliCache coherent without waiting for the 5-minute TTL to expire.
///
/// The SSE channel carries "{eventName}\n{dataJson}" strings so clients can parse both fields.
/// </summary>
public class RfqNotificationService
{
    /// <summary>
    /// Stable identity for this proxy process: "{MachineName}:{startupGuid}".
    /// Stamped on every outgoing bus message so receiving proxies can skip their own events.
    /// </summary>
    public static readonly string ProxyId =
        $"{Environment.MachineName}:{Guid.NewGuid():N}";

    private readonly ConcurrentDictionary<Guid, Channel<string>> _subscribers = new();
    private readonly ServiceBusClient?                           _sbClient;
    private readonly ServiceBusSender?                           _sbSender;
    private readonly SliCacheService                             _sliCache;
    private readonly IConfiguration                              _config;
    private readonly ILogger<RfqNotificationService>             _log;

    private ServiceBusProcessor? _sbProcessor;

    private static readonly JsonSerializerOptions _busReadOpts =
        new() { PropertyNameCaseInsensitive = true };

    private static readonly JsonSerializerOptions _busWriteOpts =
        new() { PropertyNamingPolicy = JsonNamingPolicy.CamelCase };

    public RfqNotificationService(
        IConfiguration config,
        SliCacheService sliCache,
        ILogger<RfqNotificationService> log)
    {
        _config   = config;
        _sliCache = sliCache;
        _log      = log;

        var connStr   = config["ServiceBus:ConnectionString"];
        var topicName = config["ServiceBus:TopicName"] ?? "rfq-updates";

        if (!string.IsNullOrWhiteSpace(connStr))
        {
            try
            {
                _sbClient = new ServiceBusClient(connStr);
                _sbSender = _sbClient.CreateSender(topicName);
                _log.LogInformation(
                    "[ServiceBus] Publisher connected to topic '{Topic}' — ProxyId={ProxyId}",
                    topicName, ProxyId);
            }
            catch (Exception ex)
            {
                _log.LogWarning(ex,
                    "[ServiceBus] Failed to create sender — cross-machine notifications disabled");
            }
        }
        else
        {
            _log.LogInformation(
                "[ServiceBus] ConnectionString not configured — cross-machine notifications disabled");
        }
    }

    // ── Proxy bus listener (cross-proxy cache coherency) ─────────────────────

    /// <summary>
    /// Subscribes this proxy to the rfq-updates topic so it can keep its SliCache
    /// coherent when a peer proxy on another machine writes new rows.
    /// Called once from ApplicationStarted alongside PrewarmAsync.
    /// </summary>
    public async Task StartBusListenerAsync(CancellationToken ct = default)
    {
        if (_sbClient is null) return; // no connection string — skip silently

        var connStr   = _config["ServiceBus:ConnectionString"]!;
        var topicName = _config["ServiceBus:TopicName"] ?? "rfq-updates";
        var subName   = $"{Environment.MachineName.ToLowerInvariant()}-proxy";

        try
        {
            var admin = new ServiceBusAdministrationClient(connStr);

            if (!await admin.SubscriptionExistsAsync(topicName, subName, ct))
            {
                await admin.CreateSubscriptionAsync(
                    new CreateSubscriptionOptions(topicName, subName)
                    {
                        // Matches the Shredder client subscription TTL.
                        // Orphaned subscriptions (e.g. machine renamed) auto-delete after a day.
                        AutoDeleteOnIdle = TimeSpan.FromHours(24),
                    }, ct);
            }

            _sbProcessor = _sbClient.CreateProcessor(
                topicName, subName,
                new ServiceBusProcessorOptions { AutoCompleteMessages = false });

            _sbProcessor.ProcessMessageAsync += OnProxyBusMessageAsync;
            _sbProcessor.ProcessErrorAsync   += (args) => Task.CompletedTask;

            await _sbProcessor.StartProcessingAsync(ct);
            _log.LogInformation(
                "[ServiceBus] Proxy listener started on subscription '{Sub}'", subName);
        }
        catch (Exception ex)
        {
            _log.LogWarning(ex,
                "[ServiceBus] Failed to start proxy listener — cross-proxy cache invalidation disabled");
        }
    }

    private async Task OnProxyBusMessageAsync(ProcessMessageEventArgs args)
    {
        try
        {
            var json = args.Message.Body.ToString();
            var msg  = JsonSerializer.Deserialize<RfqProcessedNotification>(json, _busReadOpts);

            // Only act on SR events from OTHER proxies; ignore our own publications.
            if (msg is { EventType: "SR" } &&
                !string.IsNullOrEmpty(msg.RfqId) &&
                msg.ProxyId != ProxyId)
            {
                if (msg.SliRows is { Count: > 0 } rows)
                {
                    _sliCache.MergeRfqRows(msg.RfqId, rows);
                    _log.LogDebug(
                        "[ServiceBus] Proxy received SR from {PeerId} — merged {Count} rows for {RfqId}",
                        msg.ProxyId, rows.Count, msg.RfqId);
                }
                else
                {
                    _sliCache.InvalidateRfq(msg.RfqId);
                    _log.LogDebug(
                        "[ServiceBus] Proxy received SR from {PeerId} — invalidated cache for {RfqId}",
                        msg.ProxyId, msg.RfqId);
                }
            }
        }
        catch (Exception ex)
        {
            _log.LogWarning(ex, "[ServiceBus] Failed to process proxy subscription message");
        }
        finally
        {
            await args.CompleteMessageAsync(args.Message);
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

    /// <summary>Publishes an "IncomingCall" event so Shredder shows a toast with caller and CRM data.</summary>
    public void NotifyIncomingCall(string callerName, string callerPhone,
        string? bpName = null, string? popupMessage = null, string? contactName = null,
        string? callLogSpItemId = null) =>
        NotifyRfqProcessed(new RfqProcessedNotification
        {
            EventType       = "IncomingCall",
            CallerName      = callerName,
            CallerPhone     = callerPhone,
            BpName          = bpName,
            PopupMessage    = popupMessage,
            ContactName     = contactName,
            CallLogSpItemId = callLogSpItemId,
        });

    /// <summary>Publishes a "Message" event so all Shredder clients update their message thread.</summary>
    public void NotifyMessage(Models.MessageRecord msg) =>
        NotifyRfqProcessed(new RfqProcessedNotification
        {
            EventType         = "Message",
            MsgFrom           = msg.From,
            MsgTo             = msg.To,
            MsgBody           = msg.Body,
            MsgConversationId = msg.ConversationId,
            MsgChannel        = msg.Channel,
            MsgDirection      = msg.Direction,
            MsgTimestamp      = msg.TimestampUtc,
        });

    /// <summary>Publishes a "MessageRead" event so other Shredder instances clear the unread badge.</summary>
    public void NotifyMessageRead(string conversationId) =>
        NotifyRfqProcessed(new RfqProcessedNotification
        {
            EventType         = "MessageRead",
            MsgConversationId = conversationId,
        });

    /// <summary>
    /// Publishes an "ERP" event to Service Bus so all Shredder clients update
    /// their local ERP document cache.
    /// </summary>
    public void NotifyErpDocument(Models.ErpBusRecord record) =>
        NotifyRfqProcessed(new RfqProcessedNotification
        {
            EventType   = "ERP",
            ErpDocument = record,
        });

    /// <summary>
    /// Broadcasts an <c>rfq-processed</c> SSE event to all connected SSE clients
    /// and publishes the same payload to Azure Service Bus so Shredder instances on
    /// other machines are notified.  The notification is tagged with <see cref="ProxyId"/>
    /// before dispatch so peer proxies can skip their own reflected events.
    /// Safe to call from any thread.
    /// </summary>
    public void NotifyRfqProcessed(RfqProcessedNotification notification)
    {
        // Tag with this proxy's identity before serialising.
        notification.ProxyId = ProxyId;

        var json = JsonSerializer.Serialize(notification, _busWriteOpts);

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
