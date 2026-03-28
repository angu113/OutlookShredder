using System.Collections.Concurrent;
using System.Threading.Channels;

namespace OutlookShredder.Proxy.Services;

/// <summary>
/// Singleton pub/sub bus for Server-Sent Events.
/// When an RFQ row is successfully written to SharePoint the poller (or the
/// add-in taskpane via the extract endpoint) calls <see cref="NotifyRfqProcessed"/>,
/// which fans the event out to every connected SSE client.
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

    /// <summary>
    /// Broadcasts an <c>rfq-processed</c> event to every connected SSE client.
    /// Safe to call from any thread.
    /// </summary>
    public void NotifyRfqProcessed()
    {
        foreach (var ch in _subscribers.Values)
            ch.Writer.TryWrite("rfq-processed");
    }
}
