using System.Text.Json;
using Azure.Messaging.ServiceBus;
using Microsoft.Extensions.Hosting;

namespace OutlookShredder.Proxy.Services;

/// <summary>
/// Consumes <c>sms-status-jobs</c> (competing-consumers — exactly one proxy per status event) and applies
/// each SignalWire delivery-status update via <see cref="InquiryService.UpdateMessageStatusAsync"/>, which
/// patches the SharePoint Messages row AND refreshes the sending proxy's in-memory cache + notifies. Without
/// this consumer the Azure Function ingress (<c>OutlookShredder.SmsWebhook</c>, which has no SharePoint
/// access) has nowhere to deliver status callbacks, so outbound messages stay stuck at their initial
/// "queued" placeholder forever. Mirrors <see cref="SmsInboundQueueProcessor"/>.
/// </summary>
public sealed class SmsStatusQueueProcessor : IHostedService, IAsyncDisposable
{
    private readonly SmsStatusQueue  _queue;
    private readonly InquiryService  _inquiry;
    private readonly ILogger<SmsStatusQueueProcessor> _log;
    private ServiceBusProcessor? _processor;

    private static readonly JsonSerializerOptions _json = new() { PropertyNameCaseInsensitive = true };

    public SmsStatusQueueProcessor(SmsStatusQueue queue, InquiryService inquiry,
        ILogger<SmsStatusQueueProcessor> log)
    {
        _queue = queue; _inquiry = inquiry; _log = log;
    }

    public async Task StartAsync(CancellationToken ct)
    {
        if (!_queue.Enabled) { _log.LogInformation("[SMS] status queue processor disabled (no Service Bus)"); return; }
        await _queue.EnsureQueueAsync(ct);
        _processor = _queue.CreateProcessor();
        if (_processor is null) return;
        _processor.ProcessMessageAsync += OnMessageAsync;
        _processor.ProcessErrorAsync   += args =>
        {
            _log.LogWarning(args.Exception, "[SMS] status processor error ({Source}) on '{Entity}'",
                args.ErrorSource, args.EntityPath);
            return Task.CompletedTask;
        };
        await _processor.StartProcessingAsync(ct);
        _log.LogInformation("[SMS] status queue processor listening on '{Queue}'", _queue.QueueName);
    }

    private async Task OnMessageAsync(ProcessMessageEventArgs args)
    {
        try
        {
            var job = JsonSerializer.Deserialize<SmsStatusQueue.Job>(args.Message.Body.ToString(), _json);
            if (job is null || string.IsNullOrWhiteSpace(job.Sid) || string.IsNullOrWhiteSpace(job.Status))
            {
                await args.CompleteMessageAsync(args.Message);
                return;
            }

            // False (no matching row — e.g. a status for a message this proxy hasn't seen yet, or an
            // internal/test message never written to SP) is a normal outcome, not an error: complete either way.
            await _inquiry.UpdateMessageStatusAsync(job.Sid, job.Status, args.CancellationToken);
            await args.CompleteMessageAsync(args.Message);
        }
        catch (Exception ex)
        {
            _log.LogWarning(ex, "[SMS] status job failed — abandoning for retry");
            try { await args.AbandonMessageAsync(args.Message); } catch { /* lock lost / already settled */ }
        }
    }

    public async Task StopAsync(CancellationToken ct)
    {
        if (_processor is not null) { try { await _processor.StopProcessingAsync(ct); } catch { } }
    }

    public async ValueTask DisposeAsync()
    {
        if (_processor is not null) await _processor.DisposeAsync();
    }
}
