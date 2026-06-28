using System.Text.Json;
using Azure.Messaging.ServiceBus;
using Microsoft.Extensions.Hosting;

namespace OutlookShredder.Proxy.Services;

/// <summary>
/// Consumes <c>sms-inbound-jobs</c> (competing-consumers — exactly one proxy per inbound SMS) and processes
/// each message once, with auto-failover (a crashed consumer's lock expires and another proxy picks it up).
/// Routes to <see cref="InquiryService.IngestInboundAsync"/> (contact consent + opt-out keywords + inquiry
/// threading + live notify). Mirrors <see cref="RfqSummaryQueueProcessor"/>.
/// </summary>
public sealed class SmsInboundQueueProcessor : IHostedService, IAsyncDisposable
{
    private readonly SmsInboundQueue  _queue;
    private readonly InquiryService   _inquiry;
    private readonly ILogger<SmsInboundQueueProcessor> _log;
    private ServiceBusProcessor? _processor;

    private static readonly JsonSerializerOptions _json = new() { PropertyNameCaseInsensitive = true };

    public SmsInboundQueueProcessor(SmsInboundQueue queue, InquiryService inquiry,
        ILogger<SmsInboundQueueProcessor> log)
    {
        _queue = queue; _inquiry = inquiry; _log = log;
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

            await _inquiry.IngestInboundAsync(job.From, job.To ?? "", job.Body ?? "", job.Sid, job.MediaUrls, args.CancellationToken);
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
