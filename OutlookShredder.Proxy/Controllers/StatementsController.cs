using Microsoft.AspNetCore.Mvc;
using OutlookShredder.Proxy.Services;

namespace OutlookShredder.Proxy.Controllers;

/// <summary>
/// Customer-statements data endpoints — backed by the ForgeTaskService in-memory cache
/// (populated by the nightly Steve OB export, stored in the ForgeTasks SP list, and shared
/// across all proxies via the TASK_COMPLETE Service Bus event).
///
/// All endpoints serve from the in-memory cache; no real-time SP reads on the hot path.
/// The cache is populated at 7pm EST (scheduled run) or on proxy startup when today's data exists.
///
/// Available for general reuse across the Forge suite.
/// </summary>
[ApiController]
[Route("api/statements")]
public class StatementsController(ForgeTaskService forgeTask, SharePointService sp) : ControllerBase
{
    /// <summary>Returns the status of the last (or current) statements run.</summary>
    [HttpGet("status")]
    public IActionResult GetStatus() => Ok(new
    {
        status         = forgeTask.Status,
        asOf           = forgeTask.AsOf,
        customerCount  = forgeTask.GetCustomerNames()?.Count ?? 0,
        running        = forgeTask.IsRunning,
        lastRunMessage = forgeTask.LastRunMessage,
    });

    /// <summary>
    /// Manually triggers the customer-statements export on this proxy, bypassing the 7pm schedule
    /// and the Service Bus dedup window.  For admin/testing/recovery (e.g. re-running after a failed
    /// nightly export).  Returns 202 immediately; the run continues in the background — poll
    /// <c>GET /api/statements/status</c> for completion.  Returns 409 if a run is already in progress.
    /// Requires OpenBravo to be open in a browser with the Shredder extension active.
    /// </summary>
    [HttpPost("trigger")]
    public IActionResult Trigger()
    {
        if (!forgeTask.TryTriggerNow())
            return Conflict(new { started = false, message = "A statements run is already in progress." });
        return Accepted(new { started = true, message = "Statements export triggered. Poll /api/statements/status for completion." });
    }

    /// <summary>Returns the sorted list of customer names with outstanding balances.</summary>
    [HttpGet("customers")]
    public IActionResult GetCustomers()
    {
        var names = forgeTask.GetCustomerNames();
        if (names is null)
            return Ok(new { ready = false, customers = Array.Empty<string>() });
        return Ok(new { ready = true, customers = names });
    }

    /// <summary>
    /// Returns full structured statement data for all customers with outstanding balances.
    /// Each item includes customer name, payment terms, and a list of unpaid invoices with
    /// invoice date, document number, due date, and outstanding amount.
    /// For general reuse across the Forge suite.
    /// </summary>
    [HttpGet("data")]
    public IActionResult GetData()
    {
        var statements = forgeTask.GetStatements();
        if (statements is null)
            return Ok(new { ready = false, statements = Array.Empty<object>() });
        return Ok(new { ready = true, statements });
    }

    /// <summary>
    /// Provisions the ForgeTasks SharePoint list with all required columns and seeds the
    /// customer-statements-export task record.  Idempotent — safe to call multiple times.
    /// </summary>
    [HttpPost("setup")]
    public async Task<IActionResult> Setup(CancellationToken ct)
    {
        try
        {
            await sp.EnsureForgeTasksListAsync(ct);
            return Ok(new { success = true, message = "ForgeTasks list ensured and task record seeded." });
        }
        catch (Exception ex)
        {
            return StatusCode(500, new { success = false, error = ex.Message });
        }
    }
}
