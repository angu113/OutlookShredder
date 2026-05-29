using Microsoft.AspNetCore.Mvc;
using OutlookShredder.Proxy.Models;
using OutlookShredder.Proxy.Services;

namespace OutlookShredder.Proxy.Controllers;

[ApiController]
[Route("api/stock-needed")]
public class StockNeededController : ControllerBase
{
    private readonly StockNeededService _sn;
    public StockNeededController(StockNeededService sn) => _sn = sn;

    [HttpGet]
    public async Task<IActionResult> GetActive() =>
        Ok(await _sn.GetActiveAsync());

    [HttpPost]
    public async Task<IActionResult> Create([FromBody] CreateStockNeededItemRequest req)
    {
        var item = await _sn.CreateAsync(req);
        return Ok(item);
    }

    [HttpPatch("{spItemId:int}")]
    public async Task<IActionResult> Patch(int spItemId, [FromBody] PatchStockNeededItemRequest req)
    {
        var found = await _sn.PatchAsync(spItemId, req);
        return found ? Ok() : NotFound();
    }

    [HttpDelete("{spItemId:int}")]
    public async Task<IActionResult> Delete(int spItemId)
    {
        await _sn.DeleteAsync(spItemId);
        return NoContent();
    }
}
