using OutlookShredder.Proxy.Models;

namespace OutlookShredder.Proxy.Services;

/// <summary>
/// In-memory cache + SharePoint persistence for the Stock Needed scratchpad.
/// Items are logically deleted by writing a RfqId; GET returns only active items (RfqId == null).
/// </summary>
public class StockNeededService : IHostedService
{
    private readonly SharePointService          _sp;
    private readonly ILogger<StockNeededService> _log;

    private readonly List<StockNeededItem> _cache = [];
    private readonly SemaphoreSlim         _lock  = new(1, 1);
    private Timer?                         _refreshTimer;

    public StockNeededService(SharePointService sp, ILogger<StockNeededService> log)
    {
        _sp  = sp;
        _log = log;
    }

    public async Task StartAsync(CancellationToken ct)
    {
        try
        {
            await _sp.EnsureStockNeededListAsync(ct);
            var items = await _sp.ReadStockNeededItemsAsync(ct);
            _cache.AddRange(items);
            _log.LogInformation("[SN] Loaded {Count} stock-needed items", items.Count);
        }
        catch (Exception ex)
        {
            _log.LogWarning(ex, "[SN] Startup load failed — starting with empty cache");
        }

        _refreshTimer = new Timer(_ => _ = RefreshFromSpAsync(), null,
            TimeSpan.FromSeconds(60), TimeSpan.FromSeconds(60));
    }

    public Task StopAsync(CancellationToken ct)
    {
        _refreshTimer?.Dispose();
        return Task.CompletedTask;
    }

    public async Task<List<StockNeededItem>> GetActiveAsync()
    {
        await _lock.WaitAsync();
        try { return _cache.Where(i => i.RfqId == null).OrderBy(i => i.ProductName).ToList(); }
        finally { _lock.Release(); }
    }

    public async Task<StockNeededItem> CreateAsync(CreateStockNeededItemRequest req, CancellationToken ct = default)
    {
        var item = await _sp.WriteStockNeededItemAsync(req, ct);
        await _lock.WaitAsync(ct);
        try { _cache.Add(item); }
        finally { _lock.Release(); }
        return item;
    }

    public async Task<bool> PatchAsync(int spItemId, PatchStockNeededItemRequest req, CancellationToken ct = default)
    {
        await _sp.PatchStockNeededItemAsync(spItemId, req, ct);
        await _lock.WaitAsync(ct);
        try
        {
            var existing = _cache.FirstOrDefault(i => i.SpItemId == spItemId);
            if (existing is null) return false;
            if (req.ProductName      is not null) existing.ProductName      = req.ProductName;
            if (req.ProductSearchKey is not null) existing.ProductSearchKey = req.ProductSearchKey;
            if (req.Category         is not null) existing.Category         = req.Category;
            if (req.Shape            is not null) existing.Shape            = req.Shape;
            if (req.QuantityNeeded   is not null) existing.QuantityNeeded   = req.QuantityNeeded;
            if (req.SizeRequested    is not null) existing.SizeRequested    = req.SizeRequested;
            if (req.Notes            is not null) existing.Notes            = req.Notes;
            if (req.RfqId            is not null) existing.RfqId            = req.RfqId;
            return true;
        }
        finally { _lock.Release(); }
    }

    public async Task<bool> DeleteAsync(int spItemId, CancellationToken ct = default)
    {
        await _sp.DeleteStockNeededItemAsync(spItemId, ct);
        await _lock.WaitAsync(ct);
        try
        {
            var i = _cache.FindIndex(x => x.SpItemId == spItemId);
            if (i < 0) return false;
            _cache.RemoveAt(i);
            return true;
        }
        finally { _lock.Release(); }
    }

    private async Task RefreshFromSpAsync()
    {
        try
        {
            var fresh = await _sp.ReadStockNeededItemsAsync();
            await _lock.WaitAsync();
            try { _cache.Clear(); _cache.AddRange(fresh); }
            finally { _lock.Release(); }
        }
        catch (Exception ex) { _log.LogWarning(ex, "[SN] Refresh failed"); }
    }
}
