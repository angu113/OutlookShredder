using Microsoft.AspNetCore.Mvc;
using OutlookShredder.Proxy.Services;

namespace OutlookShredder.Proxy.Controllers;

[ApiController]
public class MailStatusController : ControllerBase
{
    private readonly MailPollerService _poller;

    public MailStatusController(MailPollerService poller) => _poller = poller;

    /// <summary>GET /api/mail/status — live snapshot of poller, reprocess batch, rate limiter, and in-flight messages.</summary>
    [HttpGet("/api/mail/status")]
    public IActionResult GetStatus() => Ok(_poller.GetStatus());
}
