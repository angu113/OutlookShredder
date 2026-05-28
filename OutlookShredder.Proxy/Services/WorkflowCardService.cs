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
    private Timer?                      _refreshTimer;

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
            await _sp.EnsureDeliveryServicesListAsync(ct);
            var cards = await _sp.ReadWorkflowCardsAsync(ct);
            _cache.AddRange(cards);
            _log.LogInformation("[WF] Loaded {Count} workflow cards", cards.Count);
        }
        catch (Exception ex)
        {
            _log.LogWarning(ex, "[WF] Startup load failed — starting with empty cache");
        }

        // Refresh from SP every 60 s so cards created by other proxy instances are picked up.
        _refreshTimer = new Timer(_ => _ = RefreshFromSpAsync(), null,
            TimeSpan.FromSeconds(60), TimeSpan.FromSeconds(60));
    }

    public Task StopAsync(CancellationToken ct)
    {
        _refreshTimer?.Dispose();
        return Task.CompletedTask;
    }

    private async Task RefreshFromSpAsync()
    {
        try
        {
            using var cts  = new CancellationTokenSource(TimeSpan.FromSeconds(30));
            var freshCards = await _sp.ReadWorkflowCardsAsync(cts.Token);

            List<WorkflowCard> newCards;
            await _lock.WaitAsync();
            try
            {
                var existingIds = _cache.Select(c => c.SpItemId).ToHashSet();
                newCards = freshCards.Where(c => !existingIds.Contains(c.SpItemId)).ToList();
                _cache.Clear();
                _cache.AddRange(freshCards);
            }
            finally { _lock.Release(); }

            // Publish cards that arrived via another proxy — notifies live Trigger clients.
            foreach (var card in newCards)
            {
                _log.LogInformation("[WF] SP refresh found new card {Id} ({Doc}) — publishing to bus", card.SpItemId, card.DocumentNumber);
                Publish("Created", card);
            }
        }
        catch (Exception ex)
        {
            _log.LogWarning(ex, "[WF] Periodic SP refresh failed");
        }
    }

    public async Task<List<WorkflowCard>> GetAllAsync()
    {
        await _lock.WaitAsync();
        try { return [.._cache.Where(c => !c.IsCompleted)]; }
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
                DocumentNumber  = req.DocumentNumber,
                CustomerName    = req.CustomerName,
                DocumentType    = req.DocumentType,
                Tab             = req.Tab,
                AssignedDate    = req.AssignedDate,
                SortOrder       = maxOrder + 1,
                Notes           = req.Notes,
                ErpSpItemId     = req.ErpSpItemId,
                DeliveryAddress = req.DeliveryAddress,
                DeliveryMethod  = req.DeliveryMethod,
                RagStatus       = req.RagStatus,
                DeliveryService = req.DeliveryService,
                WasAutoCreated  = req.WasAutoCreated,
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

            if (req.Tab             is not null) card.Tab             = req.Tab;
            if (req.AssignedDate    is not null) card.AssignedDate    = req.AssignedDate;
            if (req.SortOrder       is not null) card.SortOrder       = req.SortOrder.Value;
            if (req.Notes           is not null) card.Notes           = req.Notes;
            if (req.IsCompleted     is not null) card.IsCompleted     = req.IsCompleted.Value;
            if (req.RagStatus       is not null) card.RagStatus       = req.RagStatus       == "" ? null : req.RagStatus;
            if (req.DeliveryService is not null) card.DeliveryService = req.DeliveryService == "" ? null : req.DeliveryService;
            if (req.WasAutoCreated  is not null) card.WasAutoCreated  = req.WasAutoCreated.Value;

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

    /// <summary>
    /// Called by FileWatcherService after a PickingSlip is written to SP.
    /// Routes into the Trigger "Prioritize" intake (AssignedDate="") so the user can schedule it.
    ///   Processing: created only when one or more configured shop-operation keywords matched in B: comments.
    ///   Delivery:   created only when DeliveryMethod is literally "Delivery" (case-insensitive).
    /// </summary>
    public async Task AutoCreateFromPickingSlipAsync(
        ErpExtraction extraction,
        string? erpSpItemId,
        IReadOnlyList<string> processOps,
        CancellationToken ct)
    {
        if (extraction.DocumentType != "PickingSlip") return;

        var docNum = extraction.DocumentNumber ?? "";
        if (string.IsNullOrEmpty(docNum)) return;

        // Guard: skip if a non-completed card already exists for this doc + tab (in any column, incl. Prioritize)
        bool hasProcessing, hasDelivery;
        await _lock.WaitAsync(ct);
        try
        {
            hasProcessing = _cache.Any(c => c.DocumentNumber == docNum && c.Tab == "Processing" && !c.IsCompleted);
            hasDelivery   = _cache.Any(c => c.DocumentNumber == docNum && c.Tab == "Delivery"   && !c.IsCompleted);
        }
        finally { _lock.Release(); }

        if (processOps.Count > 0 && !hasProcessing)
            await CreateAsync(new CreateWorkflowCardRequest
            {
                DocumentNumber  = docNum,
                CustomerName    = extraction.CustomerName,
                DocumentType    = extraction.DocumentType,
                Tab             = "Processing",
                AssignedDate    = "",
                ErpSpItemId     = erpSpItemId,
                DeliveryAddress = extraction.DeliveryAddress,
                DeliveryMethod  = extraction.DeliveryMethod,
                WasAutoCreated  = true,
            }, ct);

        var dm = extraction.DeliveryMethod?.Trim();
        bool isLiteralDelivery = !string.IsNullOrEmpty(dm) &&
                                 dm.Equals("Delivery", StringComparison.OrdinalIgnoreCase);
        if (isLiteralDelivery && !hasDelivery)
        {
            await CreateAsync(new CreateWorkflowCardRequest
            {
                DocumentNumber  = docNum,
                CustomerName    = extraction.CustomerName,
                DocumentType    = extraction.DocumentType,
                Tab             = "Delivery",
                AssignedDate    = "",
                ErpSpItemId     = erpSpItemId,
                DeliveryAddress = extraction.DeliveryAddress,
                DeliveryMethod  = extraction.DeliveryMethod,
                WasAutoCreated  = true,
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
