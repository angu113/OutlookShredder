using Microsoft.AspNetCore.Mvc;
using System.Reflection;

namespace OutlookShredder.Proxy.Controllers;

[ApiController]
[Route("api/version")]
public class VersionController : ControllerBase
{
    // GET /api/version  ->  { "version": "2026.297" }
    [HttpGet]
    public IActionResult Get()
    {
        var version = typeof(VersionController).Assembly
            .GetCustomAttribute<AssemblyInformationalVersionAttribute>()
            ?.InformationalVersion ?? "unknown";
        return Ok(new { version });
    }
}
