using Microsoft.AspNetCore.Mvc;
using OutlookShredder.Proxy.Services;

namespace OutlookShredder.Proxy.Controllers;

[ApiController]
[Route("api/shredder-config")]
public class ShredderConfigController(
    SharePointService sp,
    ILogger<ShredderConfigController> log) : ControllerBase
{
    /// <summary>
    /// GET /api/shredder-config/{name}
    /// Returns the named config entry from the ShredderConfig SP list.
    /// 404 if the row does not exist.
    /// </summary>
    [HttpGet("{name}")]
    public async Task<IActionResult> Get(string name)
    {
        try
        {
            var result = await sp.GetShredderConfigAsync(name);
            if (result is null) return NotFound();
            return Ok(new { name, value = result.Value.Value, comments = result.Value.Comments });
        }
        catch (Exception ex)
        {
            log.LogError(ex, "[Config] GET '{Name}' failed", name);
            return StatusCode(500, ex.Message);
        }
    }

    /// <summary>
    /// PUT /api/shredder-config
    /// Creates or updates a named config entry in the ShredderConfig SP list (upsert).
    /// Body: { "name": "...", "value": "...", "comments": "..." }
    /// </summary>
    [HttpPut]
    public async Task<IActionResult> Put([FromBody] ShredderConfigRequest req)
    {
        if (string.IsNullOrWhiteSpace(req.Name))
            return BadRequest("name is required.");

        try
        {
            await sp.UpsertShredderConfigAsync(req.Name, req.Value ?? "", req.Comments ?? "");
            return Ok(new { req.Name, req.Value });
        }
        catch (Exception ex)
        {
            log.LogError(ex, "[Config] PUT '{Name}' failed", req.Name);
            return StatusCode(500, ex.Message);
        }
    }
}

public sealed class ShredderConfigRequest
{
    public string?  Name     { get; set; }
    public string?  Value    { get; set; }
    public string?  Comments { get; set; }
}
