using Microsoft.AspNetCore.Mvc;
using OutlookShredder.Proxy.Models;
using OutlookShredder.Proxy.Services;

namespace OutlookShredder.Proxy.Controllers;

[ApiController, Route("api/todos")]
public class TodosController : ControllerBase
{
    private readonly TodoService _todos;

    public TodosController(TodoService todos) => _todos = todos;

    [HttpGet]
    public async Task<IActionResult> GetAll() =>
        Ok(await _todos.GetAllAsync());

    [HttpPost]
    public async Task<IActionResult> Create(
        [FromBody] CreateTodoRequest req, CancellationToken ct)
    {
        if (string.IsNullOrWhiteSpace(req.Title))
            return BadRequest(new { error = "Title is required" });
        var todo = await _todos.CreateAsync(req, ct);
        return Ok(todo);
    }

    [HttpPatch("{spItemId}")]
    public async Task<IActionResult> Update(
        string spItemId,
        [FromBody] UpdateTodoRequest req,
        CancellationToken ct)
    {
        var todo = await _todos.UpdateAsync(spItemId, req, ct);
        return todo is null ? NotFound() : Ok(todo);
    }

    [HttpDelete("{spItemId}")]
    public async Task<IActionResult> Delete(string spItemId, CancellationToken ct)
    {
        var deleted = await _todos.DeleteAsync(spItemId, ct);
        return deleted ? NoContent() : NotFound();
    }
}
