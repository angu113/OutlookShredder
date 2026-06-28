using Microsoft.AspNetCore.Mvc;
using OutlookShredder.Proxy.Services;

namespace OutlookShredder.Proxy.Controllers;

/// <summary>
/// Per-service health snapshot consumed by Shredder's Home tab dashboard.
/// Lightweight: reads cached state on in-process singletons rather than hitting Graph live.
/// </summary>
[ApiController]
[Route("api/health")]
public class HealthController : ControllerBase
{
    private readonly IConfiguration _config;
    private readonly ProductCatalogService _catalog;
    private readonly AiServiceFactory _aiFactory;
    private readonly FileWatcherService _fileWatcher;
    private readonly MailboxBridgeService _mailboxBridge;
    private readonly ForgeTaskService _forgeTask;
    private readonly CustomerCacheService _customers;

    public HealthController(
        IConfiguration config,
        ProductCatalogService catalog,
        AiServiceFactory aiFactory,
        FileWatcherService fileWatcher,
        MailboxBridgeService mailboxBridge,
        ForgeTaskService forgeTask,
        CustomerCacheService customers)
    {
        _config = config;
        _catalog = catalog;
        _aiFactory = aiFactory;
        _fileWatcher = fileWatcher;
        _mailboxBridge = mailboxBridge;
        _forgeTask = forgeTask;
        _customers = customers;
    }

    [HttpGet]
    public IActionResult Get()
    {
        var services = new List<ServiceHealth>
        {
            CheckSecrets(),
            CheckSharePoint(),
            CheckCustomers(),
            CheckMail(),
            CheckServiceBus(),
            CheckClaude(),
            CheckGemini(),
            CheckAiRouting(),
            CheckFileWatcher(),
            CheckMailboxBridge(),
            CheckStatementsTask(),
            CheckSignalWire(),
        };
        return Ok(new HealthReport(services));
    }

    // SMS (SignalWire) — "ok" + the customer-facing FromNumber once the channel is fully configured (Pulse's
    // Home-tab bubble surfaces the number from this); "disabled" until then. Mirrors SignalWireService.IsConfigured.
    private ServiceHealth CheckSignalWire()
    {
        var configured =
            !string.IsNullOrWhiteSpace(_config["SignalWire:ProjectId"])  &&
            !string.IsNullOrWhiteSpace(_config["SignalWire:ApiToken"])   &&
            !string.IsNullOrWhiteSpace(_config["SignalWire:FromNumber"]) &&
            !string.IsNullOrWhiteSpace(_config["SignalWire:SpaceUrl"]);
        return configured
            ? new("signalwire", "Messaging", "ok", _config["SignalWire:FromNumber"] ?? "")
            : new("signalwire", "Messaging", "disabled", "SMS not configured");
    }

    // WS1 — where the proxy's secrets came from this launch (Key Vault vs the cleartext file fallback).
    private ServiceHealth CheckSecrets()
        => SecretsBootstrap.Source == "keyvault"
            ? new("secrets", "Secrets", "ok", $"Key Vault ({SecretsBootstrap.Detail})")
            : new("secrets", "Secrets", "degraded", $"Local file — {SecretsBootstrap.Detail}");

    private ServiceHealth CheckSharePoint()
    {
        if (string.IsNullOrWhiteSpace(_config["SharePoint:TenantId"]) ||
            string.IsNullOrWhiteSpace(_config["SharePoint:ClientId"]) ||
            string.IsNullOrWhiteSpace(_config["SharePoint:ClientSecret"]))
        {
            return new("sharepoint", "SharePoint", "fail", "Graph credentials not configured");
        }

        if (_catalog.LastRefreshAt is null)
        {
            return _catalog.LastError is null
                ? new("sharepoint", "SharePoint", "degraded", "Catalog cache not yet loaded")
                : new("sharepoint", "SharePoint", "fail", $"Catalog refresh error: {_catalog.LastError}");
        }

        var age = DateTime.UtcNow - _catalog.LastRefreshAt.Value;
        var count = _catalog.CachedNames.Count;
        var ageText = FormatAge(age);
        // Refresh cadence is 1h; flag as degraded once we're well past two cycles.
        if (_catalog.LastError is not null)
            return new("sharepoint", "SharePoint", "degraded",
                $"Stale cache ({count} items, {ageText}) — last refresh failed: {_catalog.LastError}");
        if (age > TimeSpan.FromHours(3))
            return new("sharepoint", "SharePoint", "degraded",
                $"Cache stale ({count} items, refreshed {ageText})");
        return new("sharepoint", "SharePoint", "ok",
            $"Catalog cached ({count} items, refreshed {ageText})");
    }

    private ServiceHealth CheckMail()
    {
        var mailbox = _config["Mail:MailboxAddress"];
        if (string.IsNullOrWhiteSpace(mailbox))
            return new("mail", "Mailbox", "fail", "Mail:MailboxAddress not configured");
        return new("mail", "Mailbox", "ok", $"Polling {mailbox}");
    }

    private ServiceHealth CheckServiceBus()
    {
        var conn = _config["ServiceBus:ConnectionString"];
        if (string.IsNullOrWhiteSpace(conn))
            return new("serviceBus", "Service Bus", "fail", "ServiceBus:ConnectionString not configured");

        // All Service Bus entities the proxy uses — not just the main rfq-updates topic.
        var topic  = _config["ServiceBus:TopicName"] ?? "rfq-updates";
        var forgeQ = _config["ForgeScheduler:TaskQueueName"] ?? "forge-task-scheduler";
        var sumQ   = _config["ServiceBus:SummaryQueueName"] ?? "rfq-summary-jobs";
        return new("serviceBus", "Service Bus", "ok",
            $"topic '{topic}'  •  queues '{forgeQ}', '{sumQ}'");
    }

    private ServiceHealth CheckCustomers()
    {
        if (_customers.CacheBuiltUtc is null)
            return new("customers", "Customers / CRM", "degraded", "CRM cache not yet loaded");
        var age = DateTime.UtcNow - _customers.CacheBuiltUtc.Value;
        return new("customers", "Customers / CRM", "ok",
            $"CRM cached ({_customers.ItemCount} records, {FormatAge(age)})");
    }

    private ServiceHealth CheckClaude()
    {
        var hasKey = !string.IsNullOrWhiteSpace(_config["Anthropic:ApiKey"]);
        var model = _config["Claude:Model"] ?? "claude-sonnet-4-6";
        return hasKey
            ? new("claude", "Claude", "ok", $"API key configured ({model})")
            : new("claude", "Claude", "disabled", "Anthropic:ApiKey not set");
    }

    private ServiceHealth CheckGemini()
    {
        var hasKey = !string.IsNullOrWhiteSpace(_config["Google:ApiKey"]);
        var model = _config["Gemini:Model"] ?? "gemini-2.5-flash";
        return hasKey
            ? new("gemini", "Gemini", "ok", $"API key configured ({model})")
            : new("gemini", "Gemini", "disabled", "Google:ApiKey not set");
    }

    private ServiceHealth CheckAiRouting()
    {
        try
        {
            var svc = _aiFactory.GetService();
            return new("aiRouting", "AI Routing", "ok", svc.ProviderName);
        }
        catch (Exception ex)
        {
            return new("aiRouting", "AI Routing", "fail", ex.Message);
        }
    }

    private ServiceHealth CheckFileWatcher()
    {
        var fw = _fileWatcher.GetHealthStatus();

        if (!fw.Enabled)
            return new("fileWatcher", "File Watcher", "disabled",
                "FileWatcher:Enabled = false");

        if (!fw.WatchPathExists)
            return new("fileWatcher", "File Watcher", "fail",
                $"Watch path not found: {fw.WatchPath ?? "(none)"}");

        if (!fw.FswActive)
            return new("fileWatcher", "File Watcher", "ok",
                $"Initializing… path: {fw.WatchPath}");

        var dir = Path.GetFileName(fw.WatchPath?.TrimEnd(Path.DirectorySeparatorChar));
        return new("fileWatcher", "File Watcher", "ok",
            $"Watching {dir}  •  {fw.ProcessedCount} file{(fw.ProcessedCount == 1 ? "" : "s")} processed");
    }

    private ServiceHealth CheckMailboxBridge()
    {
        var statuses = _mailboxBridge.GetStatuses();
        if (statuses.Count == 0)
            return new("mailboxBridge", "Mailbox Bridge", "disabled", "No mailboxes configured");

        var failed = statuses.Where(m => !m.PollSucceeded && m.LastError is not null).ToList();
        if (failed.Count > 0)
            return new("mailboxBridge", "Mailbox Bridge", "fail",
                $"Poll failed: {string.Join(", ", failed.Select(m => $"{m.WatchedUpn} ({m.LastError})"))}");

        var notYetPolled = statuses.Where(m => m.LastPollAt is null).ToList();
        if (notYetPolled.Count == statuses.Count)
            return new("mailboxBridge", "Mailbox Bridge", "degraded", "Initializing… no poll completed yet");

        var stale = statuses.Where(m => m.LastPollAt is { } t && DateTime.UtcNow - t.UtcDateTime > TimeSpan.FromMinutes(5)).ToList();
        if (stale.Count > 0)
            return new("mailboxBridge", "Mailbox Bridge", "degraded",
                $"Poll stale: {string.Join(", ", stale.Select(m => m.WatchedUpn))}");

        var total = statuses.Sum(m => m.MessageCount);
        return new("mailboxBridge", "Mailbox Bridge", "ok",
            $"{statuses.Count} mailbox(es), {total} message(s) cached");
    }

    private ServiceHealth CheckStatementsTask()
    {
        // In-memory only — /api/health must not hit Graph live (see class summary).  The richer
        // cross-proxy-truthful read is GET /api/forge/task-status.
        var s = _forgeTask.GetTaskStatusInMemory();
        return s.Health switch
        {
            "ok"      => new("statements", "Statements", "ok",
                             $"{s.LastRunMessage} (exported {FormatAge(DateTime.UtcNow - s.LastRunAt!.Value)})"),
            "running" => new("statements", "Statements", "degraded", "Export in progress…"),
            "stale"   => new("statements", "Statements", "degraded",
                             s.LastRunAt.HasValue
                                ? $"Last success {FormatAge(DateTime.UtcNow - s.LastRunAt.Value)} — not today; use Fetch"
                                : "Last success was not today; use Fetch"),
            "fail"    => new("statements", "Statements", "fail",
                             $"Last export failed: {s.LastRunMessage}"),
            _         => new("statements", "Statements", "disabled", "No export yet today"),
        };
    }

    private static string FormatAge(TimeSpan age)
    {
        if (age.TotalMinutes < 1) return "<1m ago";
        if (age.TotalMinutes < 60) return $"{(int)age.TotalMinutes}m ago";
        if (age.TotalHours < 24) return $"{(int)age.TotalHours}h ago";
        return $"{(int)age.TotalDays}d ago";
    }

    public record ServiceHealth(string Id, string Label, string Status, string Detail);
    public record HealthReport(List<ServiceHealth> Services);
}
