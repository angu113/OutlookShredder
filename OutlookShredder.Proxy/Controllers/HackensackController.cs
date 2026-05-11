using MailKit;
using MailKit.Net.Imap;
using Microsoft.AspNetCore.Mvc;
using OutlookShredder.Proxy.Services;

namespace OutlookShredder.Proxy.Controllers;

[ApiController]
[Route("api/hackensack")]
public class HackensackController : ControllerBase
{
    private const string MailboxAddress  = "hackensack@metalsupermarkets.com";
    private const string ProcessedFlag   = "RFQ-Processed";

    private readonly DelegatedTokenProvider         _tokens;
    private readonly ILogger<HackensackController>  _log;

    public HackensackController(DelegatedTokenProvider tokens, ILogger<HackensackController> log)
    {
        _tokens = tokens;
        _log    = log;
    }

    /// <summary>GET /api/hackensack/inbox?top=50</summary>
    [HttpGet("inbox")]
    public async Task<IActionResult> GetInboxAsync([FromQuery] int top = 50, CancellationToken ct = default)
    {
        if (!_tokens.IsConfigured)
            return StatusCode(503, new { error = "auth_not_configured", message = "No MSAL cache — complete device-code sign-in first." });

        string token;
        try
        {
            token = await _tokens.GetAccessTokenAsync(ct);
        }
        catch (Exception ex)
        {
            _log.LogWarning(ex, "Hackensack token unavailable");
            return StatusCode(503, new { error = "auth_unavailable", message = ex.Message });
        }

        try
        {
            using var imap = new ImapClient();
            await imap.ConnectAsync("outlook.office365.com", 993,
                MailKit.Security.SecureSocketOptions.SslOnConnect, ct);
            await imap.AuthenticateAsync(
                new MailKit.Security.SaslMechanismOAuth2(MailboxAddress, token), ct);

            var inbox = imap.Inbox;
            await inbox.OpenAsync(FolderAccess.ReadOnly, ct);

            // Fetch the last N messages by sequence range (most recent last in IMAP).
            var count = inbox.Count;
            if (count == 0)
            {
                await imap.DisconnectAsync(true, ct);
                return Ok(Array.Empty<MailMessageSummary>());
            }

            var start = Math.Max(0, count - top);
            var summaries = await inbox.FetchAsync(
                start, -1,
                MessageSummaryItems.UniqueId
                | MessageSummaryItems.Envelope
                | MessageSummaryItems.Flags,  // includes custom keywords
                ct);

            var messages = summaries
                .OrderByDescending(s => s.Date)
                .Take(top)
                .Select(s => new MailMessageSummary(
                    s.UniqueId.ToString(),
                    s.Envelope.Subject ?? "(no subject)",
                    s.Envelope.From.Mailboxes.FirstOrDefault()?.Address ?? "",
                    s.Envelope.From.Mailboxes.FirstOrDefault()?.Name ?? "",
                    s.Date.ToString("o"),
                    (s.Flags & MessageFlags.Seen) != 0,
                    false,  // HasAttachments: not fetched in summary for performance
                    s.Keywords?.Contains(ProcessedFlag, StringComparer.OrdinalIgnoreCase) ?? false
                ))
                .ToArray();

            await imap.DisconnectAsync(true, ct);
            return Ok(messages);
        }
        catch (Exception ex)
        {
            _log.LogWarning(ex, "IMAP inbox fetch failed");
            return StatusCode(503, new { error = "imap_error", message = ex.Message });
        }
    }
}

public sealed record MailMessageSummary(
    string Id,
    string Subject,
    string FromAddress,
    string FromName,
    string ReceivedAt,
    bool   IsRead,
    bool   HasAttachments,
    bool   IsProcessed
);
