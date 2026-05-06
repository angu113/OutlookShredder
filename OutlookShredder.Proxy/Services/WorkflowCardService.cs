using OutlookShredder.Proxy.Models;

namespace OutlookShredder.Proxy.Services;

/// <summary>
/// In-memory cache + SharePoint persistence for workflow cards (picking/shipping slip board).
/// Acts as a hosted service so the cache is pre-loaded before the first HTTP request.
/// Thread-safe: all mutations go through _lock.
/// </summary>
public class WorkflowCardService : IHostedService
{
    private readonly SharePointService             _sp;
    private readonly RfqNotificationService        _notify;
    private readonly ILogger<WorkflowCardService>  _log;

    private readonly List<WorkflowCard> _cache = [];
    private readonly SemaphoreSlim      _lock  = new(1, 1);

    public WorkflowCardService(
        SharePointService sp,
        RfqNotificationService notify,
        ILogger<WorkflowCardService> log)
    {
        _sp     = sp;
        _notify = notify;
        _log    = log;
    }

    public async Task StartAsync(CancellationToken ct)
    {
        try
        {
            await _sp.EnsureWorkflowCardsListAsync(ct);
            var cards = await _sp.ReadWorkflowCardsAsync(ct);
            _cache.AddRange(cards);
            _log.LogInformation("[WF] Loaded {Count} workflow cards", cards.Count);
        }
        catch (Exception ex)
        {
            _log.LogWarning(ex, "[WF] Startup load failed — starting with empty cache");
        }
    }

    public Task StopAsync(CancellationToken ct) => Task.CompletedTask;

    public async Task<List<WorkflowCard>> GetAllAsync()
    {
        await _lock.WaitAsync();
        try { return [.._cache]; }
        finally { _lock.Release(); }
    }

    public async Task<WorkflowCard> CreateAsync(CreateWorkflowCardRequest req, CancellationToken ct)
    {
        await _lock.WaitAsync(ct);
        try
        {
            var maxOrder = _cache
                .Where(c => c.Tab == req.Tab && c.AssignedDate == req.AssignedDate)
                .Select(c => c.SortOrder)
                .DefaultIfEmpty(-1)
                .Max();

            var card = new WorkflowCard
            {
                DocumentNumber = req.DocumentNumber,
                CustomerName   = req.CustomerName,
                DocumentType   = req.DocumentType,
                Tab            = req.Tab,
                AssignedDate   = req.AssignedDate,
                SortOrder      = maxOrder + 1,
                Notes          = req.Notes,
                ErpSpItemId    = req.ErpSpItemId,
            };

            card.SpItemId = await _sp.WriteWorkflowCardAsync(card, ct);
            _cache.Add(card);
            Publish("Created", card);
            return card;
        }
        finally { _lock.Release(); }
    }

    public async Task<WorkflowCard?> UpdateAsync(int spItemId, UpdateWorkflowCardRequest req, CancellationToken ct)
    {
        await _lock.WaitAsync(ct);
        try
        {
            var card = _cache.FirstOrDefault(c => c.SpItemId == spItemId);
            if (card is null) return null;

            if (req.Tab          is not null) card.Tab          = req.Tab;
            if (req.AssignedDate is not null) card.AssignedDate = req.AssignedDate;
            if (req.SortOrder    is not null) card.SortOrder    = req.SortOrder.Value;
            if (req.Notes        is not null) card.Notes        = req.Notes;

            await _sp.UpdateWorkflowCardAsync(spItemId, req, ct);
            Publish("Updated", card);
            return card;
        }
        finally { _lock.Release(); }
    }

    public async Task DeleteAsync(int spItemId, CancellationToken ct)
    {
        await _lock.WaitAsync(ct);
        try
        {
            var card = _cache.FirstOrDefault(c => c.SpItemId == spItemId);
            if (card is null) return;
            _cache.Remove(card);
            await _sp.DeleteWorkflowCardAsync(spItemId, ct);
            _notify.NotifyRfqProcessed(new RfqProcessedNotification
            {
                EventType         = "WorkflowCard",
                WorkflowAction    = "Deleted",
                WorkflowDeletedId = spItemId,
            });
        }
        finally { _lock.Release(); }
    }

    /// <summary>Called by FileWatcherService after a PickingSlip is written to SP.</summary>
    public async Task AutoCreateFromPickingSlipAsync(ErpExtraction extraction, string? erpSpItemId, CancellationToken ct)
    {
        if (extraction.DocumentType != "PickingSlip") return;

        var today = DateOnly.FromDateTime(DateTime.Today).ToString("yyyy-MM-dd");

        // Processing: always for picking slips
        await CreateAsync(new CreateWorkflowCardRequest
        {
            DocumentNumber = extraction.DocumentNumber ?? "",
            CustomerName   = extraction.CustomerName,
            DocumentType   = extraction.DocumentType,
            Tab            = "Processing",
            AssignedDate   = today,
            ErpSpItemId    = erpSpItemId,
        }, ct);

        // Delivery: only when delivery method is not pickup
        var dm = extraction.DeliveryMethod?.Trim();
        if (dm is not null && !dm.Equals("Pickup", StringComparison.OrdinalIgnoreCase))
        {
            await CreateAsync(new CreateWorkflowCardRequest
            {
                DocumentNumber = extraction.DocumentNumber ?? "",
                CustomerName   = extraction.CustomerName,
                DocumentType   = extraction.DocumentType,
                Tab            = "Delivery",
                AssignedDate   = today,
                ErpSpItemId    = erpSpItemId,
            }, ct);
        }
    }

    private void Publish(string action, WorkflowCard card) =>
        _notify.NotifyRfqProcessed(new RfqProcessedNotification
        {
            EventType      = "WorkflowCard",
            WorkflowAction = action,
            WorkflowCard   = card,
        });
}
