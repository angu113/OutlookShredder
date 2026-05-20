using OutlookShredder.Proxy.Models;

namespace OutlookShredder.Proxy.Services;

/// <summary>
/// CRUD service for ShredderTodos SP list + bus event fan-out.
/// Maintains an in-memory cache pre-loaded at startup (same pattern as WorkflowCardService).
/// </summary>
public class TodoService : IHostedService
{
    private readonly SharePointService        _sp;
    private readonly RfqNotificationService   _notify;
    private readonly ILogger<TodoService>     _log;

    private readonly List<ShredderTodo> _cache = [];
    private readonly SemaphoreSlim      _lock  = new(1, 1);

    public TodoService(
        SharePointService      sp,
        RfqNotificationService notify,
        ILogger<TodoService>   log)
    {
        _sp     = sp;
        _notify = notify;
        _log    = log;
    }

    public async Task StartAsync(CancellationToken ct)
    {
        try
        {
            DateTime? since = null;
            try
            {
                var cfg = await _sp.GetShredderConfigAsync("RfqDataStartDate");
                if (cfg.HasValue && DateTime.TryParse(cfg.Value.Value, out var d))
                    since = d;
            }
            catch { /* config not set — load all */ }

            var todos = await _sp.ReadTodosAsync(since, ct);
            _cache.AddRange(todos);
            _log.LogInformation("[Todo] Loaded {Count} todos (since {Since})", todos.Count, since?.ToString("yyyy-MM-dd") ?? "all");

            // Backfill any todos that pre-date the TodoId column
            const string chars = "ABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789";
            foreach (var todo in todos.Where(t => string.IsNullOrEmpty(t.TodoId) && t.SpItemId is not null))
            {
                var newId = "TD" + new string(Enumerable.Range(0, 6).Select(_ => chars[Random.Shared.Next(chars.Length)]).ToArray());
                await _sp.PatchTodoIdAsync(todo.SpItemId!, newId, ct);
                todo.TodoId = newId;
                _log.LogInformation("[Todo] Backfilled TodoId={Id} on {SpId}", newId, todo.SpItemId);
            }
        }
        catch (Exception ex)
        {
            _log.LogWarning(ex, "[Todo] Cache pre-load failed — will serve from SP on first request");
        }
    }

    public Task StopAsync(CancellationToken ct) => Task.CompletedTask;

    public async Task<List<ShredderTodo>> GetAllAsync()
    {
        await _lock.WaitAsync();
        try { return [.._cache]; }
        finally { _lock.Release(); }
    }

    public async Task<ShredderTodo> CreateAsync(CreateTodoRequest req, CancellationToken ct)
    {
        var todo = await _sp.CreateTodoAsync(req, ct);
        await _lock.WaitAsync(ct);
        try { _cache.Add(todo); }
        finally { _lock.Release(); }
        Publish("Created", todo);
        _log.LogInformation("[Todo] Created '{Title}' by {User}", todo.Title, todo.CreatedBy);
        return todo;
    }

    public async Task<ShredderTodo?> UpdateAsync(string spItemId, UpdateTodoRequest req, CancellationToken ct)
    {
        var todo = await _sp.UpdateTodoAsync(spItemId, req, ct);
        if (todo is null) return null;
        await _lock.WaitAsync(ct);
        try
        {
            var idx = _cache.FindIndex(t => t.SpItemId == spItemId);
            if (idx >= 0) _cache[idx] = todo;
            else          _cache.Add(todo);
        }
        finally { _lock.Release(); }
        Publish("Updated", todo);
        return todo;
    }

    public async Task<bool> DeleteAsync(string spItemId, CancellationToken ct)
    {
        await _lock.WaitAsync(ct);
        var found = _cache.Any(t => t.SpItemId == spItemId);
        try { _cache.RemoveAll(t => t.SpItemId == spItemId); }
        finally { _lock.Release(); }
        if (!found) return false;
        await _sp.DeleteTodoAsync(spItemId, ct);
        _notify.NotifyRfqProcessed(new RfqProcessedNotification
        {
            EventType     = "Todo",
            TodoAction    = "Deleted",
            TodoDeletedId = spItemId,
        });
        _log.LogInformation("[Todo] Deleted {Id}", spItemId);
        return true;
    }

    private void Publish(string action, ShredderTodo todo) =>
        _notify.NotifyRfqProcessed(new RfqProcessedNotification
        {
            EventType  = "Todo",
            TodoAction = action,
            TodoItem   = todo,
        });
}
