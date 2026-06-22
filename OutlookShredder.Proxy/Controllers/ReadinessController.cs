using Microsoft.AspNetCore.Mvc;
using OutlookShredder.Proxy.Services;

namespace OutlookShredder.Proxy.Controllers;

/// <summary>
/// Startup readiness snapshot for the client. <c>GET /api/ready</c> is auth-exempt (like /api/health) so
/// the client can poll it before the launch-token race resolves. Live transitions are pushed as
/// <c>cache-ready</c> / <c>all-ready</c> SSE events on /api/events (see <see cref="ReadinessService"/>).
/// </summary>
[ApiController]
public class ReadinessController : ControllerBase
{
    private readonly ReadinessService _readiness;

    public ReadinessController(ReadinessService readiness) => _readiness = readiness;

    [HttpGet("/api/ready")]
    public IActionResult Get()
    {
        var services = _readiness.Snapshot();
        return Ok(new
        {
            proxyReady = true,
            allReady = services.All(s => s.Ready),
            services = services.Select(s => new { id = s.Id, label = s.Label, ready = s.Ready, itemCount = s.ItemCount }),
        });
    }
}
