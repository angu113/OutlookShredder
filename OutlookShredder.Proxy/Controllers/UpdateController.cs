using Microsoft.AspNetCore.Mvc;
using OutlookShredder.Proxy.Services;

namespace OutlookShredder.Proxy.Controllers;

[ApiController]
[Route("api/update")]
public class UpdateController : ControllerBase
{
    private readonly SharePointService         _sp;
    private readonly ILogger<UpdateController> _log;

    public UpdateController(SharePointService sp, ILogger<UpdateController> log)
    {
        _sp  = sp;
        _log = log;
    }

    // GET /api/update/version?channel=dev|prod  →  { "version": "2026.158.dev" }
    [HttpGet("version")]
    public async Task<ActionResult> GetVersion([FromQuery] string? channel = null)
    {
        try
        {
            var version = await _sp.GetPublishVersionAsync(channel);
            return Ok(new { version });
        }
        catch (Exception ex)
        {
            _log.LogError(ex, "[Update] Failed to read publish version from SharePoint");
            return StatusCode(502, new { error = ex.Message });
        }
    }

    // GET /api/update/package?channel=dev|prod  →  streams a ZIP of the publish folder
    [HttpGet("package")]
    public async Task GetPackage([FromQuery] string? channel = null)
    {
        _log.LogInformation("[Update] Package download started (channel={Channel})", channel ?? "prod");
        Response.ContentType = "application/zip";
        Response.Headers.ContentDisposition = "attachment; filename=\"ShredderUpdate.zip\"";

        try
        {
            await _sp.WritePublishPackageZipAsync(Response.Body, channel);
            _log.LogInformation("[Update] Package download complete");
        }
        catch (Exception ex)
        {
            _log.LogError(ex, "[Update] Package download failed");
            // Headers already sent — can't return an error status at this point
        }
    }
}
