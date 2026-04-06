using System.Collections.Concurrent;
using System.Text.Json;
using System.Threading.Channels;
using OutlookShredder.Proxy.Models;

namespace OutlookShredder.Proxy.Services;

/// <summary>
/// Singleton pub/sub bus for Server-Sent Events.
/// When an RFQ row is successfully written to SharePoint the poller (or the
/// add-in taskpane via the extract endpoint) calls <see cref="NotifyRfqProcessed"/>,
/// which fans the event out to every connected SSE client.
/// The channel carries "{eventName}\n{dataJson}" strings so clients can parse both fields.
/// </summary>
public class RfqNotificationService
{
    private readonly ConcurrentDictionary<Guid, Channel<string>> _subscribers = new();

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
    /// Broadcasts an <c>rfq-processed</c> SSE event carrying supplier + product data.
    /// Safe to call from any thread.
    /// </summary>
    public void NotifyRfqProcessed(RfqProcessedNotification notification)
    {
        var json = JsonSerializer.Serialize(notification,
            new JsonSerializerOptions { PropertyNamingPolicy = JsonNamingPolicy.CamelCase });
        var msg  = $"rfq-processed\n{json}";
        foreach (var ch in _subscribers.Values)
            ch.Writer.TryWrite(msg);
    }
}
