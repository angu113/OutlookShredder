using Microsoft.AspNetCore.Mvc;
using OutlookShredder.Proxy.Services;

namespace OutlookShredder.Proxy.Controllers;

[ApiController]
[Route("api/qc")]
public class QcController : ControllerBase
{
    private readonly SharePointService _sp;

    public QcController(SharePointService sp) => _sp = sp;

    /// <summary>
    /// Returns the QC SharePoint list as { columns: [...], rows: [[...], ...] }.
    /// Uses the same app-only credentials as the rest of the proxy.
    /// </summary>
    [HttpGet]
    public async Task<IActionResult> GetAsync()
    {
        var result = await _sp.ReadQcListAsync();
        return Ok(result);
    }
}
