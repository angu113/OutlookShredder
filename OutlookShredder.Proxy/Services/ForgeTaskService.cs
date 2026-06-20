using System.Text.Json;
using OutlookShredder.Proxy.Models;

namespace OutlookShredder.Proxy.Services;

/// <summary>
/// Hosted service that owns the Forge scheduled-task infrastructure:
///  - Loads the ForgeTasks SP list at startup; populates the in-memory statements cache
///    if today's successful run is already stored.
///  - Runs a 1-minute timer that enqueues <c>customer-statements-export</c> at 7 pm EST.
///    Azure Service Bus duplicate detection (MessageId = "task:yyyyMMdd") guarantees that
///    only ONE proxy ever executes a given run even when several proxies enqueue simultaneously.
///  - Processes queue messages: triggers the Steve OB export, parses the CSV, stores the
///    full structured result in SP, updates the in-memory cache, and publishes a TASK_COMPLETE
///    topic event so all other proxies refresh from SP.
///  - Handles TASK_COMPLETE events from peer proxies by refreshing the cache from SP.
/// </summary>
public class ForgeTaskService : BackgroundService
{
    private static readonly TimeZoneInfo _est =
        TimeZoneInfo.FindSystemTimeZoneById("Eastern Standard Time");

    private static readonly JsonSerializerOptions _jsonOpts =
        new() { PropertyNamingPolicy = JsonNamingPolicy.CamelCase, PropertyNameCaseInsensitive = true };

    private const string TaskName = "customer-statements-export";
    private const string ImmediateOverdueEmailTask = "immediate-overdue-email";   // 4pm EST

    private readonly ForgeSchedulerQueue      _queue;
    private readonly SharePointService        _sp;
    private readonly RfqNotificationService   _notify;
    private readonly MailService              _mail;
    private readonly IConfiguration           _config;
    private readonly ILogger<ForgeTaskService> _log;

    // In-memory cache — written from one thread at a time (queue processor or bus handler);
    // read from many HTTP threads.  Plain reference assignment is atomic on x64.
    private List<CustomerStatementDto>? _statements;
    private List<string>?               _customerNames;
    private DateTime?                   _asOf;
    private string                      _status    = "none";
    private string?                    _lastRunMessage;
    private DateTime?                   _lastRunEstDate;   // EST date of last successful run
    private int                        _running;          // 0 = idle, 1 = a manual/queued run is in progress

    public ForgeTaskService(
        ForgeSchedulerQueue      queue,
        SharePointService        sp,
        RfqNotificationService   notify,
        MailService              mail,
        IConfiguration           config,
        ILogger<ForgeTaskService> log)
    {
        _queue  = queue;
        _sp     = sp;
        _notify = notify;
        _mail   = mail;
        _config = config;
        _log    = log;
    }

    // ── Public accessors for StatementsController ─────────────────────────────

    public string                      Status         => _status;
    public DateTime?                   AsOf           => _asOf;
    public string?                     LastRunMessage => _lastRunMessage;
    public bool                        IsRunning      => Volatile.Read(ref _running) == 1;
    public List<string>?               GetCustomerNames()  => _customerNames;
    public List<CustomerStatementDto>? GetStatements()     => _statements;

    /// <summary>
    /// Durable health of the statements task: the SharePoint record (truthful across proxies — a
    /// peer's failed nightly run is visible here, since failures don't broadcast on the bus) merged
    /// with this proxy's in-memory running flag + cache state.  Falls back to in-memory if SP is
    /// unreachable.  For a home/system health badge — distinct from the hot-path in-memory-only
    /// <c>GET /api/statements/status</c>.
    /// </summary>
    public async Task<ForgeTaskStatus> GetTaskStatusAsync(CancellationToken ct = default)
    {
        ForgeTaskRecord? record = null;
        try { record = await _sp.GetForgeTaskAsync(TaskName, ct); }
        catch (Exception ex) { _log.LogWarning(ex, "[ForgeTask] task-status SP read failed — using in-memory"); }

        var running = IsRunning;
        var status  = record?.LastRunStatus  ?? _status;
        var lastAt  = record?.LastRunAt       ?? _asOf;
        var lastMsg = record?.LastRunMessage  ?? _lastRunMessage;

        return new ForgeTaskStatus(
            TaskName,
            record?.Enabled ?? true,
            record?.ScheduleTime,
            record?.TaskType,
            status,
            lastAt,
            lastMsg,
            record?.LastRunBy,
            running,
            _statements is not null,
            DeriveHealth(status, lastAt, running));
    }

    /// <summary>
    /// In-memory-only health snapshot (no SharePoint read) for the cheap, frequently-polled
    /// <c>GET /api/health</c> dashboard aggregation.  Reflects this proxy's last run + peer
    /// successes (refreshed via TASK_COMPLETE) + the startup SP load — but not a peer's *failure*
    /// (those don't broadcast).  Use <see cref="GetTaskStatusAsync"/> for the cross-proxy-truthful read.
    /// </summary>
    public ForgeTaskStatus GetTaskStatusInMemory() => new(
        TaskName,
        Enabled:        true,
        ScheduleTime:   null,
        TaskType:       null,
        LastRunStatus:  _status,
        LastRunAt:      _asOf,
        LastRunMessage: _lastRunMessage,
        LastRunBy:      null,
        Running:        IsRunning,
        CacheLoaded:    _statements is not null,
        Health:         DeriveHealth(_status, _asOf, IsRunning));

    /// <summary>Derives the badge health: ok (success today) / stale (prior-day success) / fail / running / unknown.</summary>
    private string DeriveHealth(string? status, DateTime? lastAt, bool running)
    {
        if (running) return "running";
        if (status == "success")
        {
            var fresh = false;
            if (lastAt.HasValue)
            {
                var estRun = TimeZoneInfo.ConvertTimeFromUtc(lastAt.Value, _est);
                var estNow = TimeZoneInfo.ConvertTimeFromUtc(DateTime.UtcNow, _est);
                fresh = estRun.Date == estNow.Date;
            }
            return fresh ? "ok" : "stale";
        }
        return status == "failed" ? "fail" : "unknown";
    }

    /// <summary>
    /// Manually runs the statements export on THIS proxy immediately, bypassing the Service Bus
    /// queue and the 7pm schedule/dedup guards.  Intended for admin/testing/recovery — e.g. after a
    /// failed nightly run, where the queue's 25h duplicate-detection window would otherwise swallow a
    /// re-enqueue of the same <c>task:yyyyMMdd</c> message.  Runs in the background and returns
    /// immediately; poll <see cref="Status"/> for the outcome.  Returns false if a run is already
    /// in progress (manual or queued).
    /// </summary>
    public bool TryTriggerNow(string? taskName = null)
    {
        if (Interlocked.CompareExchange(ref _running, 1, 0) != 0)
            return false;

        var task = taskName ?? TaskName;
        _ = Task.Run(async () =>
        {
            try { await ExecuteTaskAsync(task, CancellationToken.None); }
            catch { /* ExecuteTaskAsync already logged + set _status = "failed" */ }
            finally { Volatile.Write(ref _running, 0); }
        });
        return true;
    }

    // ── BackgroundService ─────────────────────────────────────────────────────

    protected override async Task ExecuteAsync(CancellationToken stoppingToken)
    {
        // Ensure the SB queue exists before starting the processor.
        try { await _queue.EnsureQueueAsync(stoppingToken); }
        catch (Exception ex) { _log.LogWarning(ex, "[ForgeTask] EnsureQueue failed — queue may not be available"); }

        // Load today's cached result from SP (best-effort — SP may not be ready yet).
        try { await LoadFromSpAsync(stoppingToken); }
        catch (Exception ex) { _log.LogWarning(ex, "[ForgeTask] SP load on startup failed"); }

        // Start the competing-consumer queue processor.
        var processor = _queue.CreateProcessor();
        if (processor is not null)
        {
            processor.ProcessMessageAsync += OnQueueMessageAsync;
            processor.ProcessErrorAsync   += args =>
            {
                _log.LogWarning(args.Exception, "[ForgeTask] Queue processor error");
                return Task.CompletedTask;
            };
            await processor.StartProcessingAsync(stoppingToken);
            _log.LogInformation("[ForgeTask] Queue processor started on '{Queue}'", _queue.QueueName);
        }
        else
        {
            _log.LogInformation("[ForgeTask] Service Bus not configured — scheduled tasks disabled");
        }

        // 1-minute schedule check loop.
        using var timer = new PeriodicTimer(TimeSpan.FromMinutes(1));
        while (await timer.WaitForNextTickAsync(stoppingToken))
        {
            try { await CheckScheduleAsync(stoppingToken); }
            catch (Exception ex) { _log.LogWarning(ex, "[ForgeTask] Schedule check failed"); }
        }

        if (processor is not null)
        {
            try { await processor.StopProcessingAsync(); }
            catch { /* best-effort stop */ }
        }
    }

    private async Task CheckScheduleAsync(CancellationToken ct)
    {
        var estNow = TimeZoneInfo.ConvertTimeFromUtc(DateTime.UtcNow, _est);

        // 4pm EST — Immediate-overdue email. The queue's MessageId dedup ({task}:{yyyyMMdd}) guarantees a
        // single send per day even if several proxies enqueue at 16:00.
        if (estNow is { Hour: 16, Minute: 0 })
        {
            _log.LogInformation("[ForgeTask] 4pm EST — enqueueing {Task}", ImmediateOverdueEmailTask);
            await _queue.EnqueueAsync(ImmediateOverdueEmailTask, ct);
        }

        // 7pm EST — statements export.
        if (estNow is { Hour: 19, Minute: 0 })
        {
            // Guard: skip if the last successful load was within 23 hours (e.g. manual trigger earlier today).
            // 23 (not 24) leaves margin so a run completing a few minutes after 19:00 doesn't skip the next
            // night's 19:00 run (_asOf is the completion time).
            if (_asOf.HasValue && (DateTime.UtcNow - _asOf.Value).TotalHours < 23)
            {
                _log.LogInformation(
                    "[ForgeTask] Skipping 7pm enqueue — last run was {Hours:F1}h ago (< 23h threshold)",
                    (DateTime.UtcNow - _asOf.Value).TotalHours);
                return;
            }

            // Guard: skip if today's EST date already has a recorded run in cache.
            if (_lastRunEstDate.HasValue && _lastRunEstDate.Value.Date == estNow.Date) return;

            _log.LogInformation("[ForgeTask] 7pm EST — enqueueing {Task}", TaskName);
            await _queue.EnqueueAsync(TaskName, ct);
        }
    }

    // ── Queue message handler ─────────────────────────────────────────────────

    private async Task OnQueueMessageAsync(Azure.Messaging.ServiceBus.ProcessMessageEventArgs args)
    {
        // Mutual exclusion with a manual TriggerNow run on this proxy — abandon so SB redelivers
        // once the in-flight run finishes (rather than starting two concurrent Steve exports).
        if (Interlocked.CompareExchange(ref _running, 1, 0) != 0)
        {
            _log.LogInformation("[ForgeTask] Queue message arrived while a run is in progress — abandoning for redelivery");
            await args.AbandonMessageAsync(args.Message);
            return;
        }

        try
        {
            var json = args.Message.Body.ToString();
            var msg  = JsonSerializer.Deserialize<ForgeTaskQueueMessage>(json, _jsonOpts);
            if (msg is not null)
                await ExecuteTaskAsync(msg.TaskName, args.CancellationToken);
            await args.CompleteMessageAsync(args.Message);
        }
        catch (Exception ex)
        {
            _log.LogError(ex, "[ForgeTask] Queue message handler failed");
            await args.AbandonMessageAsync(args.Message);
        }
        finally
        {
            Volatile.Write(ref _running, 0);
        }
    }

    private async Task ExecuteTaskAsync(string taskName, CancellationToken ct)
    {
        // Lightweight tasks that don't run the Steve export (and must not touch the statements status).
        if (taskName == ImmediateOverdueEmailTask)
        {
            await RunOverdueEmailAsync(TermsBucket.Immediate, send: true, ct);
            return;
        }

        _log.LogInformation("[ForgeTask] Starting '{Task}' on {Machine}", taskName, Environment.MachineName);

        _status         = "running";
        _lastRunMessage = "Export in progress…";

        try
        {
            await _sp.UpdateForgeTaskStatusAsync(taskName, "running", null, Environment.MachineName, ct);
        }
        catch (Exception ex)
        {
            _log.LogWarning(ex, "[ForgeTask] Could not mark task running in SP — proceeding anyway");
        }

        try
        {
            // Trigger Steve OB export.
            SteveState.ClearExportResult();
            SteveState.SetPending("sales-invoice-export");

            // Poll for the CSV (5 min timeout, 5 s interval).
            string? csvPath = null;
            var deadline = DateTime.UtcNow.AddMinutes(5);
            while (DateTime.UtcNow < deadline && !ct.IsCancellationRequested)
            {
                await Task.Delay(5000, ct);
                var path = SteveState.GetExportResult();
                if (!string.IsNullOrEmpty(path) && File.Exists(path)) { csvPath = path; break; }
            }

            if (string.IsNullOrEmpty(csvPath))
                throw new Exception("Steve export timed out — no CSV received within 5 minutes. " +
                    "Ensure OpenBravo is open in a browser with the Shredder extension active.");

            // Parse CSV into structured statements.
            var csvContent = await File.ReadAllTextAsync(csvPath, ct);
            var statements = StatementCsvParser.Parse(csvContent);
            _log.LogInformation("[ForgeTask] Parsed {Count} customers from '{File}'",
                statements.Count, Path.GetFileName(csvPath));

            // Store in SP (clears old result, writes new).
            var resultJson    = JsonSerializer.Serialize(statements, _jsonOpts);
            var customersJson = JsonSerializer.Serialize(
                statements.Select(s => s.CustomerName).ToList(), _jsonOpts);
            await _sp.StoreForgeTaskResultAsync(taskName, resultJson, customersJson, ct);
            await _sp.UpdateForgeTaskStatusAsync(
                taskName, "success", $"{statements.Count} customers", Environment.MachineName, ct);

            // Update in-memory cache.
            var estNow      = TimeZoneInfo.ConvertTimeFromUtc(DateTime.UtcNow, _est);
            _statements     = statements;
            _customerNames  = statements.Select(s => s.CustomerName).ToList();
            _asOf           = DateTime.UtcNow;
            _status         = "success";
            _lastRunMessage = $"{statements.Count} customers";
            _lastRunEstDate = estNow;

            // Notify peer proxies to refresh their caches.
            _notify.NotifyTaskComplete(taskName);
            _log.LogInformation("[ForgeTask] '{Task}' complete — {Count} customers stored",
                taskName, statements.Count);

            // Right after the 7pm load: email the Net-terms overdue customers. Runs ONLY on the proxy that
            // executed the export (peers get TASK_COMPLETE -> HandleTaskCompleteAsync, which does not send),
            // so exactly one Net email goes out. A mail failure must not fail the export (already succeeded).
            try { await RunOverdueEmailAsync(TermsBucket.Net, send: true, ct); }
            catch (Exception ex) { _log.LogWarning(ex, "[ForgeTask] post-load Net-overdue email failed"); }
        }
        catch (Exception ex)
        {
            _log.LogError(ex, "[ForgeTask] Task '{Task}' failed", taskName);
            _status         = "failed";
            _lastRunMessage = ex.Message;
            try
            {
                await _sp.UpdateForgeTaskStatusAsync(
                    taskName, "failed", ex.Message, Environment.MachineName, ct);
            }
            catch { /* best-effort */ }
            throw; // re-throw so the message is abandoned
        }
        finally
        {
            SteveState.ClearPending();
            SteveState.ClearExportResult();
        }
    }

    // ── Peer notification handler (called from RfqNotificationService) ────────

    public async Task HandleTaskCompleteAsync(string taskName, CancellationToken ct = default)
    {
        _log.LogInformation("[ForgeTask] Peer completed '{Task}' — refreshing from SP", taskName);
        try
        {
            var resultJson = await _sp.GetForgeTaskResultAsync(taskName, ct);
            if (string.IsNullOrEmpty(resultJson)) return;

            var statements = JsonSerializer.Deserialize<List<CustomerStatementDto>>(resultJson, _jsonOpts);
            if (statements is null) return;

            var estNow      = TimeZoneInfo.ConvertTimeFromUtc(DateTime.UtcNow, _est);
            _statements     = statements;
            _customerNames  = statements.Select(s => s.CustomerName).ToList();
            _asOf           = DateTime.UtcNow;
            _status         = "success";
            _lastRunEstDate = estNow;
            _log.LogInformation("[ForgeTask] Cache refreshed from peer — {Count} customers", statements.Count);
        }
        catch (Exception ex) { _log.LogWarning(ex, "[ForgeTask] Cache refresh from SP failed"); }
    }

    // ── ShadowCat overdue-customer emails ─────────────────────────────────────

    /// <summary>
    /// Computes the overdue customers for one terms bucket (Immediate / Net), renders the list to a PDF,
    /// and — when <paramref name="send"/> is true — emails it to the AR address. Uses the in-memory
    /// statements snapshot; at 4pm (before tonight's 7pm run) that may be empty, so it falls back to the last
    /// stored SP result (prior night). Sends a "none overdue" note when the bucket is clear; does nothing
    /// only when there's no data at all. With <paramref name="send"/>=false it renders without emailing (the
    /// PDF is returned) — used by the admin preview endpoint so the live AR inbox isn't hit during testing.
    /// </summary>
    public async Task<OverdueEmailResult> RunOverdueEmailAsync(TermsBucket bucket, bool send, CancellationToken ct = default)
    {
        var asOf  = TimeZoneInfo.ConvertTimeFromUtc(DateTime.UtcNow, _est).Date;
        var label = bucket == TermsBucket.Immediate ? "Immediate" : "Net";

        var statements = _statements;
        if (statements is null)
        {
            // 4pm path: today's export hasn't run; use the most recent stored snapshot (prior night).
            try
            {
                var json = await _sp.GetForgeTaskResultAsync(TaskName, ct);
                if (!string.IsNullOrEmpty(json))
                    statements = JsonSerializer.Deserialize<List<CustomerStatementDto>>(json, _jsonOpts);
            }
            catch (Exception ex) { _log.LogWarning(ex, "[ForgeTask] overdue email: SP snapshot load failed"); }
        }
        if (statements is null)
        {
            _log.LogWarning("[ForgeTask] {Label}-overdue email skipped — no statements data available", label);
            return new OverdueEmailResult(label, asOf, 0, 0m, false, null);
        }

        var rows  = StatementOverdue.OverdueRows(statements, bucket, asOf);
        var total = rows.Sum(r => r.Overdue);
        byte[]? pdf = rows.Count > 0 ? StatementOverdueDocument.Render(rows, label, asOf) : null;

        if (send)
        {
            var to      = _config["ShadowCat:OverdueEmailTo"] ?? "hackensack@metalsupermarkets.com";
            var subject = $"ShadowCat: {label} Customers Overdue on {asOf:MMMM d, yyyy}";
            var body    = rows.Count > 0
                ? $"{rows.Count} {label}-terms customer(s) overdue as of {asOf:MMMM d, yyyy}. " +
                  $"Total overdue {total:C0}. See the attached PDF."
                : $"No {label}-terms customers are overdue as of {asOf:MMMM d, yyyy}.";

            await _mail.SendSupplierInquiryAsync(
                to, subject, body,
                attachmentName:        pdf is null ? null : $"ShadowCat-{label}-Overdue-{asOf:yyyy-MM-dd}.pdf",
                attachmentBytes:       pdf,
                attachmentContentType: pdf is null ? null : "application/pdf");

            _log.LogInformation("[ForgeTask] Sent {Label}-overdue email to {To}: {Count} customer(s), asOf {AsOf:yyyy-MM-dd}",
                label, to, rows.Count, asOf);
        }

        return new OverdueEmailResult(label, asOf, rows.Count, total, send, send ? null : pdf);
    }

    // ── Startup SP load ───────────────────────────────────────────────────────

    private async Task LoadFromSpAsync(CancellationToken ct)
    {
        var record = await _sp.GetForgeTaskAsync(TaskName, ct);
        if (record is null)
        {
            _log.LogInformation("[ForgeTask] No SP record for '{Task}' — first run", TaskName);
            return;
        }

        _status         = record.LastRunStatus ?? "none";
        _lastRunMessage = record.LastRunMessage;
        // NB: _asOf stays success-only (set in the success path below) — CheckScheduleAsync's 23h
        // guard treats it as "last successful run", so a failed/running record must not populate it.

        if (_status != "success" || record.LastRunAt is null || string.IsNullOrEmpty(record.ResultData))
            return;

        // Only load into memory if the cached run is from today (EST).
        var runEst   = TimeZoneInfo.ConvertTimeFromUtc(record.LastRunAt.Value, _est);
        var todayEst = TimeZoneInfo.ConvertTimeFromUtc(DateTime.UtcNow, _est);
        if (runEst.Date != todayEst.Date)
        {
            _log.LogInformation("[ForgeTask] Cached run is from {Date:yyyy-MM-dd} — not today; skipping pre-load",
                runEst.Date);
            return;
        }

        var statements = JsonSerializer.Deserialize<List<CustomerStatementDto>>(record.ResultData, _jsonOpts);
        if (statements is null) return;

        _statements     = statements;
        _customerNames  = statements.Select(s => s.CustomerName).ToList();
        _asOf           = record.LastRunAt;
        _lastRunEstDate = runEst;
        _log.LogInformation("[ForgeTask] Pre-loaded {Count} customers from SP cache (run {Time:HH:mm} EST)",
            statements.Count, runEst);
    }
}

/// <summary>Result of a ShadowCat overdue-email run. <see cref="Pdf"/> is populated only on a preview
/// (send=false) so the admin endpoint can return it without hitting the live AR inbox.</summary>
public sealed record OverdueEmailResult(
    string Bucket, DateTime AsOf, int Count, decimal TotalOverdue, bool Sent, byte[]? Pdf);
