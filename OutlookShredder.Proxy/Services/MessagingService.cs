using OutlookShredder.Proxy.Models;

namespace OutlookShredder.Proxy.Services;

public class MessagingService
{
    private readonly SharePointService         _sp;
    private readonly SignalWireService         _sw;
    private readonly RfqNotificationService    _notify;
    private readonly MailService               _mail;
    private readonly IConfiguration            _config;
    private readonly ILogger<MessagingService> _log;

    public MessagingService(SharePointService sp, SignalWireService sw,
        RfqNotificationService notify, MailService mail, IConfiguration config,
        ILogger<MessagingService> log)
    {
        _sp     = sp;
        _sw     = sw;
        _notify = notify;
        _mail   = mail;
        _config = config;
        _log    = log;
    }

    public string[] KnownUsers =>
        _config.GetSection("Messaging:KnownUsers").Get<string[]>() ?? [];

    public async Task<bool> SendInternalAsync(string from, string to, string body,
        string? subject = null, CancellationToken ct = default)
    {
        var record = new MessageRecord
        {
            From           = from,
            To             = to,
            Channel        = "internal",
            Direction      = "out",
            Subject        = subject,
            Body           = body,
            ConversationId = InternalConvId(from, to),
            TimestampUtc   = DateTimeOffset.UtcNow.ToString("o"),
            IsRead         = false,
        };

        try { await _sp.WriteMessageAsync(record, ct); }
        catch (Exception ex) { _log.LogWarning(ex, "[Messaging] SP write failed for internal message"); }

        _notify.NotifyMessage(record);
        return true;
    }

    public async Task<bool> SendEmailAsync(string from, string to, string subject, string body,
        CancellationToken ct = default)
    {
        if (to.EndsWith("@mithrilmetals.com", StringComparison.OrdinalIgnoreCase))
        {
            _log.LogWarning("[Messaging] Email send rejected — internal recipient {To}", to);
            return false;
        }

        try { await _mail.SendSupplierInquiryAsync(to, subject, body); }
        catch (Exception ex)
        {
            _log.LogWarning(ex, "[Messaging] Email send via MailService failed to {To}", to);
            return false;
        }

        var convId = EmailConvId(to);
        var record = new MessageRecord
        {
            From           = from,
            To             = to,
            Channel        = "email",
            Direction      = "out",
            Subject        = subject,
            Body           = body,
            ConversationId = convId,
            TimestampUtc   = DateTimeOffset.UtcNow.ToString("o"),
            IsRead         = true,
        };

        try { await _sp.WriteMessageAsync(record, ct); }
        catch (Exception ex) { _log.LogWarning(ex, "[Messaging] SP write failed for outbound email"); }

        _notify.NotifyMessage(record);
        return true;
    }

    public async Task<bool> SendSmsAsync(string from, string to, string body, CancellationToken ct = default)
    {
        if (!_sw.IsConfigured)
        {
            _log.LogWarning("[Messaging] SignalWire not configured — SMS unavailable");
            return false;
        }

        var sid = await _sw.SendSmsAsync(to, body, ct);

        var record = new MessageRecord
        {
            From           = from,
            To             = to,
            Channel        = "sms",
            Direction      = "out",
            Body           = body,
            ConversationId = SmsConvId(to),
            TimestampUtc   = DateTimeOffset.UtcNow.ToString("o"),
            IsRead         = true,
            ExternalId     = sid,
        };

        try { await _sp.WriteMessageAsync(record, ct); }
        catch (Exception ex) { _log.LogWarning(ex, "[Messaging] SP write failed for outbound SMS"); }

        _notify.NotifyMessage(record);
        return sid is not null;
    }

    public async Task HandleInboundSmsAsync(string from, string to, string body, string? sid, CancellationToken ct = default)
    {
        var record = new MessageRecord
        {
            From           = from,
            To             = to,
            Channel        = "sms",
            Direction      = "in",
            Body           = body,
            ConversationId = SmsConvId(from),
            TimestampUtc   = DateTimeOffset.UtcNow.ToString("o"),
            IsRead         = false,
            ExternalId     = sid,
        };

        try { await _sp.WriteMessageAsync(record, ct); }
        catch (Exception ex) { _log.LogWarning(ex, "[Messaging] SP write failed for inbound SMS"); }

        _notify.NotifyMessage(record);
    }

    public async Task MarkReadAsync(string conversationId, CancellationToken ct = default)
    {
        try { await _sp.MarkConversationReadAsync(conversationId, ct); }
        catch (Exception ex) { _log.LogWarning(ex, "[Messaging] SP MarkRead failed for '{Id}'", conversationId); }

        _notify.NotifyMessageRead(conversationId);
    }

    public static string InternalConvId(string a, string b)
    {
        var parts = new[] { a.ToLowerInvariant(), b.ToLowerInvariant() };
        Array.Sort(parts);
        return $"internal:{parts[0]}|{parts[1]}";
    }

    public static string SmsConvId(string phone) =>
        $"sms:{NormalizePhone(phone)}";

    public static string EmailConvId(string email) =>
        $"email:{email.ToLowerInvariant()}";

    private static string NormalizePhone(string phone) =>
        "+" + new string(phone.Where(char.IsDigit).ToArray());
}
