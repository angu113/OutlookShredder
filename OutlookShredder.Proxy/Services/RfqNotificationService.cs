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
    private readonly MailCacheService                            _mailCache;
    private readonly Lazy<ForgeTaskService>                      _forgeTasks;
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
        MailCacheService mailCache,
        Lazy<ForgeTaskService> forgeTasks,
        ILogger<RfqNotificationService> log)
    {
        _config     = config;
        _sliCache   = sliCache;
        _mailCache  = mailCache;
        _forgeTasks = forgeTasks;
        _log        = log;

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
            _sbProcessor.ProcessErrorAsync   += args =>
            {
                _log.LogWarning(args.Exception,
                    "[ServiceBus] Proxy listener error ({Source}) on '{Entity}'",
                    args.ErrorSource, args.EntityPath);
                return Task.CompletedTask;
            };

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

            // Refresh the in-memory statements cache when a peer proxy completes the OB export.
            if (msg is { EventType: "TASK_COMPLETE" } && !string.IsNullOrEmpty(msg.TaskName) && msg.ProxyId != ProxyId)
            {
                _ = Task.Run(() => _forgeTasks.Value.HandleTaskCompleteAsync(msg.TaskName!));
                _log.LogDebug("[ServiceBus] Proxy received TASK_COMPLETE '{Task}' from {PeerId}",
                    msg.TaskName, msg.ProxyId);
            }

            // Keep the local MailCache coherent with peer proxies' workbench writes.
            if (msg is { EventType: "Mail" } && msg.ProxyId != ProxyId && !string.IsNullOrEmpty(msg.MailItemId))
            {
                if (string.Equals(msg.MailAction, "Deleted", StringComparison.OrdinalIgnoreCase))
                    _mailCache.ApplyBusDelete(msg.MailItemId);
                else if (msg.MailItem is not null)
                    _mailCache.ApplyBusItem(msg.MailItem);
                _log.LogDebug("[ServiceBus] Proxy received Mail '{Action}' from {PeerId} for {Id}",
                    msg.MailAction, msg.ProxyId, msg.MailItemId);
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
    public void NotifyIncomingCall(string callerName, string? callerPhone,
        string? bpName = null, string? popupMessage = null, string? contactName = null,
        string? callLogSpItemId = null,
        IList<CustomerLookupResult>? allMatches = null) =>
        NotifyRfqProcessed(new RfqProcessedNotification
        {
            EventType       = "IncomingCall",
            CallerName      = callerName,
            CallerPhone     = callerPhone,
            BpName          = bpName,
            PopupMessage    = popupMessage,
            ContactName     = contactName,
            CallLogSpItemId = callLogSpItemId,
            CrmMatches      = allMatches?.Count > 1
                ? allMatches.Select(m => new CrmBusMatchDto(m.BusinessPartner, m.ContactName, m.PopupMessage)).ToList()
                : null,
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
    /// Publishes a "SupplierMsgRead" event so peers update one supplier message's read state and adjust
    /// the SR/RFQ/tab unread badge cascade.
    /// </summary>
    public void NotifySupplierMessageRead(
        string userId, string rfqId, string supplierName, string messageId, bool read, string? readBy) =>
        NotifyRfqProcessed(new RfqProcessedNotification
        {
            EventType    = "SupplierMsgRead",
            RfqId        = rfqId,
            SupplierName = supplierName,
            MessageId    = messageId,
            ConvRead     = read,
            ConvReadBy   = read ? readBy : null,
            ConvReadAt   = read ? DateTimeOffset.UtcNow.ToString("o") : null,
            ConvReadUser = userId,
        });

    /// <summary>
    /// Publishes a "Mail" event so peer proxies sync their MailCache and Shredder Inbox views
    /// refresh. <paramref name="action"/> is Captured | Classified | Amended | Completed | Deleted;
    /// <paramref name="item"/> carries the full snapshot (null on Deleted).
    /// </summary>
    public void NotifyMailItem(string action, string mailItemId, Models.MailBusItem? item) =>
        NotifyRfqProcessed(new RfqProcessedNotification
        {
            EventType  = "Mail",
            MailAction = action,
            MailItemId = mailItemId,
            MailItem   = item,
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

    /// <summary>Publishes an "RFQ_SUMMARY" event so an open RFQ Focus view re-fetches the cached
    /// state-of-play after a queue worker (re)generates it for that RFQ.</summary>
    public void NotifyRfqSummary(string rfqId) =>
        NotifyRfqProcessed(new RfqProcessedNotification { EventType = "RFQ_SUMMARY", RfqId = rfqId });

    /// <summary>Publishes a "TASK_COMPLETE" event so peer proxies refresh their statements cache
    /// from SP without waiting for a manual trigger or the next startup.</summary>
    public void NotifyTaskComplete(string taskName) =>
        NotifyRfqProcessed(new RfqProcessedNotification
        {
            EventType = "TASK_COMPLETE",
            TaskName  = taskName,
        });

    /// <summary>Publishes a "PO_STATUS" event so Trigger Ordered cards re-colour live the moment a
    /// PO's confirm/payment status changes (manual dropdown, auto-confirm, bill match, or receipt).</summary>
    public void NotifyPoStatus(Models.PurchaseOrderRecord r)
    {
        bool hasEmail =
            (string.Equals(r.PaymentStatus, "Required", StringComparison.OrdinalIgnoreCase) && !string.IsNullOrWhiteSpace(r.BillMailItemId)) ||
            (string.Equals(r.PaymentStatus, "Paid",     StringComparison.OrdinalIgnoreCase) && !string.IsNullOrWhiteSpace(r.ReceiptMailItemId));
        NotifyRfqProcessed(new RfqProcessedNotification
        {
            EventType        = "PO_STATUS",
            PoSpItemId       = r.SpItemId,
            PoNumber         = r.PoNumber,
            RfqId            = r.RfqId,
            SupplierName     = r.SupplierName,
            ConfirmStatus    = r.ConfirmStatus ?? "Pending",
            PaymentStatus    = r.PaymentStatus ?? "None",
            HasPaymentEmail  = hasEmail,
            MaterialReceived = !string.IsNullOrWhiteSpace(r.MaterialReceivedAt),
        });
    }

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
