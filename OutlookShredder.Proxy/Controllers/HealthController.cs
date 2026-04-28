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

    public HealthController(
        IConfiguration config,
        ProductCatalogService catalog,
        AiServiceFactory aiFactory,
        FileWatcherService fileWatcher)
    {
        _config = config;
        _catalog = catalog;
        _aiFactory = aiFactory;
        _fileWatcher = fileWatcher;
    }

    [HttpGet]
    public IActionResult Get()
    {
        var services = new List<ServiceHealth>
        {
            CheckSharePoint(),
            CheckMail(),
            CheckServiceBus(),
            CheckClaude(),
            CheckGemini(),
            CheckAiRouting(),
            CheckFileWatcher(),
        };
        return Ok(new HealthReport(services));
    }

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
        var topic = _config["ServiceBus:TopicName"] ?? "rfq-updates";
        if (string.IsNullOrWhiteSpace(conn))
            return new("serviceBus", "Service Bus", "fail", "ServiceBus:ConnectionString not configured");
        return new("serviceBus", "Service Bus", "ok", $"Topic '{topic}' configured");
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
            return new("fileWatcher", "File Watcher", "degraded",
                $"Watcher starting… path: {fw.WatchPath}");

        var dir = Path.GetFileName(fw.WatchPath?.TrimEnd(Path.DirectorySeparatorChar));
        return new("fileWatcher", "File Watcher", "ok",
            $"Watching {dir}  •  {fw.ProcessedCount} file{(fw.ProcessedCount == 1 ? "" : "s")} processed");
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
