using Microsoft.AspNetCore.Mvc;
using OutlookShredder.Proxy.Services;

namespace OutlookShredder.Proxy.Controllers;

[ApiController]
[Route("api/phone")]
public class PhoneController : ControllerBase
{
    private readonly SharePointService _sp;
    private readonly ILogger<PhoneController> _log;

    public PhoneController(SharePointService sp, ILogger<PhoneController> log)
    {
        _sp  = sp;
        _log = log;
    }

    /// <summary>Returns the most recent call log entries, newest first.</summary>
    [HttpGet("call-log")]
    public async Task<IActionResult> GetCallLog(
        [FromQuery] int top = 200,
        CancellationToken ct = default)
    {
        try
        {
            var records = await _sp.ReadPhoneCallLogAsync(top, ct);
            return Ok(records);
        }
        catch (Exception ex)
        {
            _log.LogWarning(ex, "[Phone] Failed to read call log");
            return StatusCode(500, new { error = ex.Message });
        }
    }
}
