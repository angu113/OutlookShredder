using Microsoft.AspNetCore.Mvc;

namespace OutlookShredder.Proxy.Controllers;

[ApiController]
[Route("api/service-bus")]
public class ServiceBusController : ControllerBase
{
    private readonly IConfiguration _config;

    public ServiceBusController(IConfiguration config) => _config = config;

    /// <summary>
    /// Returns the Service Bus connection string and topic name so Shredder
    /// instances can read them from the proxy rather than keeping their own copy.
    /// Response: { configured: bool, connectionString: string|null, topicName: string }
    /// </summary>
    [HttpGet("config")]
    public IActionResult GetConfig()
    {
        var connStr = _config["ServiceBus:ConnectionString"];
        var topic   = _config["ServiceBus:TopicName"] ?? "rfq-updates";
        var configured = !string.IsNullOrWhiteSpace(connStr);
        return Ok(new { configured, connectionString = configured ? connStr : null, topicName = topic });
    }
}
