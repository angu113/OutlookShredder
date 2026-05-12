using Microsoft.AspNetCore.Mvc;
using OutlookShredder.Proxy.Models;
using OutlookShredder.Proxy.Services;

namespace OutlookShredder.Proxy.Controllers;

[ApiController, Route("api/workflow")]
public class WorkflowController : ControllerBase
{
    private readonly WorkflowCardService _wf;
    private readonly SharePointService   _sp;

    public WorkflowController(WorkflowCardService wf, SharePointService sp)
    {
        _wf = wf;
        _sp = sp;
    }

    [HttpGet("delivery-services")]
    public async Task<IActionResult> GetDeliveryServices(CancellationToken ct) =>
        Ok(await _sp.ReadDeliveryServicesAsync(ct));

    [HttpGet("cards")]
    public async Task<IActionResult> GetCards() =>
        Ok(await _wf.GetAllAsync());

    [HttpPost("cards")]
    public async Task<IActionResult> CreateCard(
        [FromBody] CreateWorkflowCardRequest req, CancellationToken ct)
    {
        var card = await _wf.CreateAsync(req, ct);
        return Ok(card);
    }

    [HttpPatch("cards/{spItemId:int}")]
    public async Task<IActionResult> UpdateCard(
        int spItemId,
        [FromBody] UpdateWorkflowCardRequest req,
        CancellationToken ct)
    {
        var card = await _wf.UpdateAsync(spItemId, req, ct);
        return card is null ? NotFound() : Ok(card);
    }

    [HttpDelete("cards/{spItemId:int}")]
    public async Task<IActionResult> DeleteCard(int spItemId, CancellationToken ct)
    {
        await _wf.DeleteAsync(spItemId, ct);
        return NoContent();
    }
}
