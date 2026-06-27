using System.Text.Json;
using Azure.Messaging.ServiceBus;
using Microsoft.Extensions.Hosting;

namespace OutlookShredder.Proxy.Services;

/// <summary>
/// Consumes <c>sms-inbound-jobs</c> (competing-consumers — exactly one proxy per inbound SMS) and processes
/// each message once, with auto-failover (a crashed consumer's lock expires and another proxy picks it up).
/// Phase 0 routes to <see cref="MessagingService.HandleInboundSmsAsync"/> (store + notify); Phase 1 repoints
/// this to <c>InquiryService.IngestInboundAsync</c> (inquiry threading + opt-out + AI draft).
/// Mirrors <see cref="RfqSummaryQueueProcessor"/>.
/// </summary>
public sealed class SmsInboundQueueProcessor : IHostedService, IAsyncDisposable
{
    private readonly SmsInboundQueue  _queue;
    private readonly MessagingService _messaging;
    private readonly ILogger<SmsInboundQueueProcessor> _log;
    private ServiceBusProcessor? _processor;

    private static readonly JsonSerializerOptions _json = new() { PropertyNameCaseInsensitive = true };

    public SmsInboundQueueProcessor(SmsInboundQueue queue, MessagingService messaging,
        ILogger<SmsInboundQueueProcessor> log)
    {
        _queue = queue; _messaging = messaging; _log = log;
    }

    public async Task StartAsync(CancellationToken ct)
    {
        if (!_queue.Enabled) { _log.LogInformation("[SMS] inbound queue processor disabled (no Service Bus)"); return; }
        await _queue.EnsureQueueAsync(ct);
        _processor = _queue.CreateProcessor();
        if (_processor is null) return;
        _processor.ProcessMessageAsync += OnMessageAsync;
        _processor.ProcessErrorAsync   += args =>
        {
            _log.LogWarning(args.Exception, "[SMS] inbound processor error ({Source}) on '{Entity}'",
                args.ErrorSource, args.EntityPath);
            return Task.CompletedTask;
        };
        await _processor.StartProcessingAsync(ct);
        _log.LogInformation("[SMS] inbound queue processor listening on '{Queue}'", _queue.QueueName);
    }

    private async Task OnMessageAsync(ProcessMessageEventArgs args)
    {
        try
        {
            var job = JsonSerializer.Deserialize<SmsInboundQueue.Job>(args.Message.Body.ToString(), _json);
            if (job is null || string.IsNullOrWhiteSpace(job.From))
            {
                await args.CompleteMessageAsync(args.Message);
                return;
            }

            await _messaging.HandleInboundSmsAsync(job.From, job.To ?? "", job.Body ?? "", job.Sid, args.CancellationToken);
            await args.CompleteMessageAsync(args.Message);
        }
        catch (Exception ex)
        {
            _log.LogWarning(ex, "[SMS] inbound job failed — abandoning for retry");
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
