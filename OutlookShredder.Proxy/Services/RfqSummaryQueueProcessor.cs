using System.Text.Json;
using Azure.Messaging.ServiceBus;
using Microsoft.Extensions.Hosting;

namespace OutlookShredder.Proxy.Services;

/// <summary>
/// Consumes the <c>rfq-summary-jobs</c> queue (competing-consumers — exactly one proxy per job) and
/// (re)generates the RFQ state-of-play summary: reads the current SLI rows, gates on ≥2 competing
/// suppliers, generates via <see cref="RfqStateOfPlayService"/> when the cache is stale, persists it,
/// and broadcasts an <c>RFQ_SUMMARY</c> topic event so focus views refresh.
/// </summary>
public sealed class RfqSummaryQueueProcessor : IHostedService, IAsyncDisposable
{
    private readonly RfqSummaryQueue        _queue;
    private readonly SharePointService      _sp;
    private readonly RfqStateOfPlayService  _state;
    private readonly RfqNotificationService _notify;
    private readonly IConfiguration         _config;
    private readonly ILogger<RfqSummaryQueueProcessor> _log;
    private ServiceBusProcessor? _processor;

    private static readonly JsonSerializerOptions _json = new() { PropertyNameCaseInsensitive = true };

    public RfqSummaryQueueProcessor(RfqSummaryQueue queue, SharePointService sp, RfqStateOfPlayService state,
        RfqNotificationService notify, IConfiguration config, ILogger<RfqSummaryQueueProcessor> log)
    {
        _queue = queue; _sp = sp; _state = state; _notify = notify; _config = config; _log = log;
    }

    public async Task StartAsync(CancellationToken ct)
    {
        if (!_queue.Enabled) { _log.LogInformation("[StateOfPlay] queue processor disabled"); return; }
        await _queue.EnsureQueueAsync(ct);
        _processor = _queue.CreateProcessor();
        if (_processor is null) return;
        _processor.ProcessMessageAsync += OnMessageAsync;
        _processor.ProcessErrorAsync   += _ => Task.CompletedTask;
        await _processor.StartProcessingAsync(ct);
        _log.LogInformation("[StateOfPlay] queue processor listening on '{Queue}'", _queue.QueueName);
    }

    private async Task OnMessageAsync(ProcessMessageEventArgs args)
    {
        try
        {
            var job = JsonSerializer.Deserialize<RfqSummaryQueue.Job>(args.Message.Body.ToString(), _json);
            var rfqId = job?.RfqId;
            if (string.IsNullOrWhiteSpace(rfqId)) { await args.CompleteMessageAsync(args.Message); return; }

            var rows = await _sp.ReadSupplierItemsByRfqIdAsync(rfqId);
            if (RfqStateOfPlayService.CompetingSuppliers(rows) < 2)   // no comparison yet — nothing to do
            {
                await args.CompleteMessageAsync(args.Message);
                return;
            }

            bool pdfs = _config.GetValue("RfqStateOfPlay:IncludePdfs", false);
            var hash  = _state.ComputeInputsHash(rows, pdfs);
            var cached = await _sp.ReadRfqSummaryAsync(rfqId);
            if (cached?.Summary is { Length: > 0 } && cached.InputsHash == hash)   // already current
            {
                await args.CompleteMessageAsync(args.Message);
                return;
            }

            var result = await _state.GenerateAsync(rfqId, rows, pdfs, null, args.CancellationToken);
            if (result is not null)
            {
                await _sp.WriteRfqSummaryAsync(rfqId, result.Summary, result.InputsHash, result.Model, result.Mode);
                _notify.NotifyRfqSummary(rfqId);
                _log.LogInformation("[StateOfPlay] regenerated summary for {Rfq}", rfqId);
            }
            await args.CompleteMessageAsync(args.Message);
        }
        catch (Exception ex)
        {
            _log.LogWarning(ex, "[StateOfPlay] job failed — abandoning for retry");
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
