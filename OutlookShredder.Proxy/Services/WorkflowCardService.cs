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
                OwnerUser       = req.OwnerUser,
                ProcessOps      = req.ProcessOps,
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
    ///   Worklist lane (Tab="Worklist"): created for EVERY picking slip so nothing is missed; any matched
    ///     shop-operation keywords ride along on the card's Notes for context.
    ///   Delivery: created for any delivery method that isn't a customer pickup / will-call
    ///     (see <see cref="IsDeliveryMethod"/>) — e.g. "Our Truck", "UPS Ground", "Delivery".
    /// Both are deduped against existing non-completed cards for the same doc + tab.
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

        // Owner = the doc's sales rep ("Customer Rep:" / "Requested By:"), else the importer
        // (the proxy host user that scanned the slip — mirrors ErpDocumentRecord.SourceUser).
        var owner = !string.IsNullOrWhiteSpace(extraction.SalesRep)
            ? extraction.SalesRep!.Trim()
            : Environment.UserName;

        // Guard: skip if a non-completed card already exists for this doc + tab (in any column, incl. Prioritize)
        bool hasWorklist, hasDelivery;
        await _lock.WaitAsync(ct);
        try
        {
            hasWorklist = _cache.Any(c => c.DocumentNumber == docNum && c.Tab == "Worklist" && !c.IsCompleted);
            hasDelivery   = _cache.Any(c => c.DocumentNumber == docNum && c.Tab == "Delivery"   && !c.IsCompleted);
        }
        finally { _lock.Release(); }

        // Worklist lane (Tab="Worklist"): every picking slip lands in Prioritize so nothing is missed.
        // Any matched shop ops ride along on ProcessOps (rendered as fab-type chips on the card); Notes is
        // left empty for the user. (Previously ops were dumped into Notes — moved out so the notes box is the
        // user's own, and the ops drive structured chips instead.)
        if (!hasWorklist)
            await CreateAsync(new CreateWorkflowCardRequest
            {
                DocumentNumber  = docNum,
                CustomerName    = extraction.CustomerName,
                DocumentType    = extraction.DocumentType,
                Tab             = "Worklist",
                AssignedDate    = "",
                ErpSpItemId     = erpSpItemId,
                DeliveryAddress = extraction.DeliveryAddress,
                DeliveryMethod  = extraction.DeliveryMethod,
                ProcessOps      = processOps.Count > 0 ? string.Join(", ", processOps) : null,
                WasAutoCreated  = true,
                OwnerUser       = owner,
            }, ct);

        // Delivery lane: any non-pickup delivery method (Our Truck / a carrier / plain "Delivery").
        bool createDelivery = IsDeliveryMethod(extraction.DeliveryMethod) && !hasDelivery;
        if (createDelivery)
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
                OwnerUser       = owner,
            }, ct);

        _log.LogInformation(
            "[WF] Auto-create {Doc}: fabrication={Fab} delivery={Del} (method='{Method}', ops={Ops})",
            docNum, !hasWorklist, createDelivery, extraction.DeliveryMethod ?? "", processOps.Count);
    }

    /// <summary>
    /// True when a slip's "Delivery Method:" means it leaves on a vehicle (our truck, a carrier, or a
    /// plain "Delivery") rather than a customer pickup. The ERP free-texts this field ("Our Truck",
    /// "UPS Ground", "Pickup", "Will Call", "Delivery", …), so we exclude the pickup / will-call variants
    /// and treat everything else non-empty as a delivery — keying only on the literal word "Delivery"
    /// missed "Our Truck", the most common real delivery value. Empty/unknown → not a delivery.
    /// </summary>
    internal static bool IsDeliveryMethod(string? deliveryMethod)
    {
        var dm = deliveryMethod?.Trim();
        if (string.IsNullOrEmpty(dm)) return false;
        var lower = dm.ToLowerInvariant();
        string[] pickupMarkers = ["pickup", "pick up", "pick-up", "will call", "will-call", "willcall"];
        return !pickupMarkers.Any(m => lower.Contains(m));
    }

    private void Publish(string action, WorkflowCard card) =>
        _notify.NotifyRfqProcessed(new RfqProcessedNotification
        {
            EventType      = "WorkflowCard",
            WorkflowAction = action,
            WorkflowCard   = card,
        });
}
