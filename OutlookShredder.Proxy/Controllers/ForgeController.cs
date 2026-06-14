using Microsoft.AspNetCore.Mvc;
using OutlookShredder.Proxy.Services;

namespace OutlookShredder.Proxy.Controllers;

/// <summary>
/// Health of the Forge scheduled tasks.  Reads the durable SharePoint record (so a peer proxy's
/// failed nightly run is visible here too) merged with this proxy's in-memory running/cache state.
/// Intended for a home/system health badge — distinct from <c>GET /api/statements/status</c>, which
/// is the in-memory-only hot path the Generate Statements tab polls.
/// </summary>
[ApiController]
[Route("api/forge")]
public class ForgeController(ForgeTaskService forgeTask) : ControllerBase
{
    /// <summary>
    /// Per-task health (currently the single <c>customer-statements-export</c> task).  Returns a
    /// <c>tasks</c> array so the shape is stable once ForgeTasks supports multiple scheduled tasks.
    /// </summary>
    [HttpGet("task-status")]
    public async Task<IActionResult> TaskStatus(CancellationToken ct)
        => Ok(new { tasks = new[] { await forgeTask.GetTaskStatusAsync(ct) } });
}
