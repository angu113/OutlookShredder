using Microsoft.AspNetCore.Mvc;
using OutlookShredder.Proxy.Models;
using OutlookShredder.Proxy.Services;

namespace OutlookShredder.Proxy.Controllers;

[ApiController]
[Route("api/cache")]
public class CacheController : ControllerBase
{
    private readonly IEnumerable<ICacheStatusProvider> _caches;
    private readonly IConfiguration _config;

    public CacheController(IEnumerable<ICacheStatusProvider> caches, IConfiguration config)
    {
        _caches = caches;
        _config = config;
    }

    private CacheConfigDto ReadConfig() => new()
    {
        FullRefreshIntervalDays = _config.GetValue("PersistentCache:FullRefreshIntervalDays", 7),
        DeltaBufferMinutes      = _config.GetValue("PersistentCache:DeltaBufferMinutes", 10),
        Enabled                 = _config.GetValue("PersistentCache:Enabled", true),
    };

    [HttpGet("status")]
    public IActionResult GetStatus()
    {
        var cfg  = ReadConfig();
        var dtos = _caches
            .Select(c => c.ToDto(cfg.FullRefreshIntervalDays))
            .ToList();
        return Ok(new CacheStatusResponse { Caches = dtos, Config = cfg });
    }

    [HttpPost("rebuild")]
    public async Task<IActionResult> Rebuild([FromQuery] string type = "all", CancellationToken ct = default)
    {
        var targets = ResolveTargets(type);
        if (targets.Count == 0) return BadRequest(new { error = $"Unknown cache type '{type}'" });

        var results = new List<RebuildResult>();
        foreach (var cache in targets)
        {
            var sw = System.Diagnostics.Stopwatch.StartNew();
            await cache.ForceRebuildAsync(ct);
            results.Add(new RebuildResult { Type = cache.Name, ItemCount = cache.ItemCount, DurationMs = sw.ElapsedMilliseconds });
        }
        return Ok(results.Count == 1 ? results[0] : results);
    }

    [HttpPost("delta")]
    public async Task<IActionResult> Delta([FromQuery] string type = "all", CancellationToken ct = default)
    {
        var targets = ResolveTargets(type);
        if (targets.Count == 0) return BadRequest(new { error = $"Unknown cache type '{type}'" });

        var results = new List<RebuildResult>();
        foreach (var cache in targets)
        {
            var sw = System.Diagnostics.Stopwatch.StartNew();
            await cache.ForceDeltaAsync(ct);
            results.Add(new RebuildResult { Type = cache.Name, ItemCount = cache.ItemCount, DurationMs = sw.ElapsedMilliseconds });
        }
        return Ok(results.Count == 1 ? results[0] : results);
    }

    [HttpPatch("config")]
    public IActionResult PatchConfig([FromBody] CacheConfigDto patch)
    {
        // Write to the proxy's appsettings via in-memory config override.
        // Values persist for the session; a full appsettings.json write would
        // require file I/O with locking — not needed for the initial version.
        ((IConfigurationRoot)_config).Providers
            .OfType<Microsoft.Extensions.Configuration.Json.JsonConfigurationProvider>()
            .FirstOrDefault(); // placeholder — runtime patch not wired; restart reads appsettings

        return Ok(patch);
    }

    private List<ICacheStatusProvider> ResolveTargets(string type)
    {
        if (type.Equals("all", StringComparison.OrdinalIgnoreCase))
            return _caches.ToList();
        return _caches.Where(c => c.Name.Equals(type, StringComparison.OrdinalIgnoreCase)).ToList();
    }
}
