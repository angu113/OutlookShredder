using OutlookShredder.Proxy.Models;
using OutlookShredder.Proxy.Services.Storage;

namespace OutlookShredder.Proxy.Services;

/// <summary>
/// Orchestrates the inbound side of the SMS customer-inquiry pipeline (Phase 1): contact consent/opt-out,
/// keyword handling (STOP/HELP/START), threading an inbound message into the customer's latest inquiry
/// (reopening a Closed one or minting a new CINQ), and publishing live updates. Depends only on the storage
/// seams (<see cref="IInquiryStore"/> / <see cref="IMessageStore"/>) and the notification/SMS services — never
/// on the SharePoint connection — so the whole pipeline ports to another store by swapping the DAO
/// registration. Channel is an attribute of the message (SMS now; email later), so nothing here is
/// SMS-specific beyond the carrier keyword set + the outbound HELP reply.
/// </summary>
public sealed class InquiryService : IHostedService
{
    private readonly IInquiryStore           _store;
    private readonly IMessageStore           _messages;
    private readonly RfqNotificationService  _notify;
    private readonly SignalWireService       _sw;
    private readonly IConfiguration          _config;
    private readonly ILogger<InquiryService> _log;

    // Stub default — real copy is tracked in wip/customer-experience-sms-inquiry.md ("Content still needed").
    private const string DefaultHelpReply =
        "Mithril Metals: text us your question and our team will help. Msg & data rates may apply. " +
        "Reply STOP to opt out.";

    public InquiryService(IInquiryStore store, IMessageStore messages, RfqNotificationService notify,
        SignalWireService sw, IConfiguration config, ILogger<InquiryService> log)
    {
        _store    = store;
        _messages = messages;
        _notify   = notify;
        _sw       = sw;
        _config   = config;
        _log      = log;
    }

    public async Task StartAsync(CancellationToken ct)
    {
        // Provision the lists/tables up front so the first inbound isn't slowed by creation and so the
        // index-at-construction invariant holds.
        try
        {
            await _store.EnsureProvisionedAsync(ct);
            await _messages.EnsureProvisionedAsync(ct);
            _log.LogInformation("[Inquiry] storage provisioned");
        }
        catch (Exception ex) { _log.LogWarning(ex, "[Inquiry] startup provisioning failed — will retry lazily"); }
    }

    public Task StopAsync(CancellationToken ct) => Task.CompletedTask;

    /// <summary>
    /// Ingests one inbound SMS (called once per message by <see cref="SmsInboundQueueProcessor"/> — the
    /// dedup queue guarantees exactly-once). Upserts the contact + consent, handles carrier keywords, then
    /// either records a compliance/info message or threads a real customer message into an inquiry.
    /// </summary>
    public async Task IngestInboundAsync(string from, string to, string body, string? sid, CancellationToken ct = default)
    {
        var phone   = InquiryRules.NormalizeE164(from);
        var now     = DateTimeOffset.UtcNow.ToString("o");
        var keyword = InquiryRules.ClassifyKeyword(body);

        // 1. Upsert the contact + apply consent transitions.
        var contact = await _store.GetContactAsync(phone, ct)
                      ?? new MessagingContact { Phone = phone, ConsentCapturedAt = now, ConsentMethod = "inbound-sms" };
        switch (keyword)
        {
            case InquiryRules.Keyword.OptOut: contact.OptOut = true;  contact.OptOutAt = now;  break;
            case InquiryRules.Keyword.OptIn:  contact.OptOut = false; contact.OptOutAt = null; break;
        }
        try { await _store.UpsertContactAsync(contact, ct); }
        catch (Exception ex) { _log.LogWarning(ex, "[Inquiry] contact upsert failed for {Phone}", phone); }

        var inquiries = await _store.GetInquiriesByPhoneAsync(phone, ct);
        var latest    = inquiries.FirstOrDefault();   // store returns most-recent first

        // 2. Carrier keyword (STOP/HELP/START family): a compliance/info signal, not a sales question — do
        //    NOT mint a CINQ or bump unread. Record it against the existing thread (if any) for audit, and
        //    answer HELP unless opted out.
        if (keyword != InquiryRules.Keyword.None)
        {
            var kwMsg = await AppendMessageAsync(from, to, body, sid, latest?.Id, now, ct);
            if (keyword == InquiryRules.Keyword.Help && !contact.OptOut)
                await SendHelpReplyAsync(from, ct);
            if (latest is not null) _notify.NotifyInquiryMessage(latest.Id, kwMsg);
            _log.LogInformation("[Inquiry] {Keyword} from {Phone} (optOut={OptOut})", keyword, phone, contact.OptOut);
            return;
        }

        // 3. Normal customer message → thread it.
        var action = InquiryRules.DecideThread(latest);
        Inquiry inquiry;
        bool isNew = false;
        if (action == InquiryRules.ThreadAction.CreateNew)
        {
            inquiry = new Inquiry
            {
                Id            = await GenerateCinqIdAsync(ct),
                CustomerPhone = phone,
                Status        = InquiryStatus.Open,
                CreatedAt     = now,
                UpdatedAt     = now,
                LastMessageAt = now,
                UnreadCount   = 1,
            };
            await _store.CreateInquiryAsync(inquiry, ct);
            isNew = true;
        }
        else
        {
            inquiry = latest!;
            if (action == InquiryRules.ThreadAction.Reopen) inquiry.Status = InquiryStatus.Open;
            inquiry.LastMessageAt = now;
            inquiry.UpdatedAt     = now;
            inquiry.UnreadCount  += 1;
            await _store.UpdateInquiryAsync(inquiry, ct);
        }

        var msg = await AppendMessageAsync(from, to, body, sid, inquiry.Id, now, ct);

        _notify.NotifyInquiry(isNew ? "Created" : "Updated", inquiry);
        _notify.NotifyInquiryMessage(inquiry.Id, msg);
        _log.LogInformation("[Inquiry] {Action} {Id} from {Phone} (unread={Unread})",
            action, inquiry.Id, phone, inquiry.UnreadCount);
    }

    /// <summary>Updates an outbound message's delivery status by SID (SignalWire status callback).</summary>
    public Task<bool> UpdateMessageStatusAsync(string sid, string status, CancellationToken ct = default)
        => _messages.UpdateStatusBySidAsync(sid, status, ct);

    private async Task<MessageRecord> AppendMessageAsync(
        string from, string to, string body, string? sid, string? inquiryId, string now, CancellationToken ct)
    {
        var msg = new MessageRecord
        {
            From           = from,
            To             = to,
            Channel        = "sms",
            Direction      = "in",
            Body           = body,
            ConversationId = MessagingService.SmsConvId(from),
            TimestampUtc   = now,
            IsRead         = false,
            ExternalId     = sid,
            InquiryId      = inquiryId,
        };
        try { await _messages.AppendAsync(msg, ct); }
        catch (Exception ex) { _log.LogWarning(ex, "[Inquiry] message append failed for {From}", from); }
        return msg;
    }

    private async Task SendHelpReplyAsync(string to, CancellationToken ct)
    {
        if (!_sw.IsConfigured) { _log.LogWarning("[Inquiry] HELP received but SignalWire not configured"); return; }
        var reply = _config["SignalWire:HelpReply"] is { Length: > 0 } cfg ? cfg : DefaultHelpReply;
        try { await _sw.SendSmsAsync(to, reply, ct); }
        catch (Exception ex) { _log.LogWarning(ex, "[Inquiry] HELP auto-reply failed to {To}", to); }
    }

    private async Task<string> GenerateCinqIdAsync(CancellationToken ct)
    {
        for (int i = 0; i < 20; i++)
        {
            var candidate = InquiryRules.RandomCinqId();
            if (await _store.GetInquiryByIdAsync(candidate, ct) is null) return candidate;
        }
        throw new InvalidOperationException("CINQ id generation exhausted its retry budget");
    }
}
