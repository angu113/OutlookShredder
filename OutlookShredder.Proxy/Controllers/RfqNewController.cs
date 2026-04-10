using Microsoft.AspNetCore.Mvc;
using OutlookShredder.Proxy.Services;

namespace OutlookShredder.Proxy.Controllers;

[ApiController]
[Route("api/rfq-new")]
public class RfqNewController(MailService mail, ILogger<RfqNewController> log) : ControllerBase
{
    /// <summary>
    /// POST /api/rfq-new/send-email
    /// Sends an RFQ email via Microsoft Graph (Mail.Send app permission required).
    /// From / Reply-To are read from appsettings Mail:FromAddress / Mail:ReplyToAddress.
    /// </summary>
    [HttpPost("send-email")]
    public async Task<IActionResult> SendEmail([FromBody] SendRfqEmailRequest req)
    {
        if (string.IsNullOrWhiteSpace(req.Subject) || string.IsNullOrWhiteSpace(req.Body))
            return BadRequest("Subject and Body are required.");

        try
        {
            await mail.SendRfqEmailAsync(req.Subject, req.Body, req.BccAddresses);
            return Ok();
        }
        catch (Exception ex)
        {
            log.LogError(ex, "[RFQ Send] Failed to send '{Subject}'", req.Subject);
            return StatusCode(500, ex.Message);
        }
    }
}

public sealed class SendRfqEmailRequest
{
    public string         Subject      { get; set; } = "";
    public string         Body         { get; set; } = "";
    public List<string>   BccAddresses { get; set; } = [];
}
