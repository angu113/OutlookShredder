using OutlookShredder.Proxy.Models;
using OutlookShredder.Proxy.Services.Sms;
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
    private readonly ISmsGateway             _sms;
    private readonly InquiryDraftService     _drafts;
    private readonly CustomerCacheService    _crm;
    private readonly IConfiguration          _config;
    private readonly ILogger<InquiryService> _log;

    // Stub default — real copy is tracked in wip/customer-experience-sms-inquiry.md ("Content still needed").
    private const string DefaultHelpReply =
        "Mithril Metals: text us your question and our team will help. Msg & data rates may apply. " +
        "Reply STOP to opt out.";

    public InquiryService(IInquiryStore store, IMessageStore messages, RfqNotificationService notify,
        ISmsGateway sms, InquiryDraftService drafts, CustomerCacheService crm, IConfiguration config,
        ILogger<InquiryService> log)
    {
        _store    = store;
        _messages = messages;
        _notify   = notify;
        _sms      = sms;
        _drafts   = drafts;
        _crm      = crm;
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

        // 3. Normal customer message → thread it. Resolve the customer from CRM (denormalised for the list +
        //    "first-time caller" detection); inbound always leaves us owing a reply (AwaitingReply).
        var crm = _crm.LookupByPhone(from);
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
                CustomerName  = crm?.BusinessPartner,
                ContactName   = crm?.ContactName,
                CreatedAt     = now,
                UpdatedAt     = now,
                LastMessageAt = now,
                UnreadCount   = 1,
                AwaitingReply = true,
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
            inquiry.AwaitingReply = true;
            inquiry.CustomerName ??= crm?.BusinessPartner;   // backfill if not resolved before
            inquiry.ContactName  ??= crm?.ContactName;
            await _store.UpdateInquiryAsync(inquiry, ct);
        }

        var msg = await AppendMessageAsync(from, to, body, sid, inquiry.Id, now, ct);

        _notify.NotifyInquiry(isNew ? "Created" : "Updated", inquiry);
        _notify.NotifyInquiryMessage(inquiry.Id, msg);
        _log.LogInformation("[Inquiry] {Action} {Id} from {Phone} (unread={Unread})",
            action, inquiry.Id, phone, inquiry.UnreadCount);

        // Phase 2: AI reply suggestion — async, never auto-sent, and detached from the queue consumer's
        // token so a slow Claude call neither blocks ingest nor is cancelled when the SB message completes.
        _ = GenerateDraftAsync(inquiry.Id, body, sid);
    }

    /// <summary>Builds + persists an AI reply suggestion for the inquiry and pushes it live. Fire-and-forget:
    /// it owns its own timeout and swallows all errors (a draft is a non-critical suggestion).</summary>
    private async Task GenerateDraftAsync(string inquiryId, string inboundBody, string? triggeringSid)
    {
        try
        {
            using var cts = new CancellationTokenSource(TimeSpan.FromMinutes(2));
            var ct = cts.Token;

            // Prior transcript = the thread minus the just-appended inbound (passed separately as "latest").
            var history    = await _messages.GetByInquiryAsync(inquiryId, 12, ct);
            var prior      = history.Count > 0 ? history.Take(history.Count - 1).ToList() : history;
            var transcript = InquiryDraftPrompt.BuildTranscript(prior);

            // Linked HSK# / notes arrive in Phase 3 (quotation linking + notes) — empty for now.
            var result = await _drafts.DraftAsync(
                new InquiryDraftInput(inboundBody, transcript, Array.Empty<string>(), null), ct);
            if (result is null) return;

            var draft = new InquiryDraft
            {
                InquiryId           = inquiryId,
                TriggeringMessageId = triggeringSid,
                Source              = DraftSource.Ai,
                Body                = result.Reply,
                SuggestedIntent     = result.Intent,
                SuggestedUrgency    = result.Urgency,
                NeedsQuote          = result.NeedsQuote,
                Status              = DraftStatus.Pending,
                CreatedAt           = DateTimeOffset.UtcNow.ToString("o"),
            };
            await _store.CreateDraftAsync(draft, ct);
            _notify.NotifyInquiryDraft(draft);
            _log.LogInformation("[Inquiry] AI draft for {Id} (intent={Intent} urgency={Urgency} needsQuote={NeedsQuote})",
                inquiryId, draft.SuggestedIntent, draft.SuggestedUrgency, draft.NeedsQuote);
        }
        catch (Exception ex) { _log.LogWarning(ex, "[Inquiry] AI draft generation failed for {Id}", inquiryId); }
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
        if (!_sms.IsConfigured) { _log.LogWarning("[Inquiry] HELP received but SMS gateway not configured"); return; }
        var reply = _config["SignalWire:HelpReply"] is { Length: > 0 } cfg ? cfg : DefaultHelpReply;
        try { await _sms.SendAsync(to, reply, ct: ct); }
        catch (Exception ex) { _log.LogWarning(ex, "[Inquiry] HELP auto-reply failed to {To}", to); }
    }

    // ── Phase 3 operator actions (called by InquiriesController) ──────────────────────────────────

    public Task<IReadOnlyList<Inquiry>> ListAsync(string? status, string? query, CancellationToken ct = default)
        => _store.GetInquiriesAsync(status, query, ct);

    public async Task<InquiryDetail?> GetDetailAsync(string inquiryId, CancellationToken ct = default)
    {
        var inquiry = await _store.GetInquiryByIdAsync(inquiryId, ct);
        if (inquiry is null) return null;
        var messages   = await _messages.GetByInquiryAsync(inquiryId, 200, ct);
        var notes      = await _store.GetNotesByInquiryAsync(inquiryId, ct);
        var quotations = await _store.GetQuotationsByInquiryAsync(inquiryId, ct);
        var drafts     = await _store.GetDraftsByInquiryAsync(inquiryId, ct);
        var contact    = await _store.GetContactAsync(inquiry.CustomerPhone, ct);

        var crm  = _crm.LookupByPhone(inquiry.CustomerPhone);
        var card = new CustomerCard(
            crm?.BusinessPartner ?? inquiry.CustomerName,
            crm?.ContactName ?? inquiry.ContactName,
            crm?.PopupMessage,
            IsFirstTime: crm is null);
        return new InquiryDetail(inquiry, [.. messages], [.. notes], [.. quotations], [.. drafts], contact, card);
    }

    /// <summary>Sends an operator reply to the customer (suppressed if opted out), records the outbound
    /// message, advances the inquiry, optionally marks the source draft Used, and pushes live updates.
    /// Throws <see cref="InvalidOperationException"/> when the contact opted out or no gateway is configured.</summary>
    public async Task<MessageRecord?> SendOperatorReplyAsync(
        string inquiryId, string body, int? fromDraftSpItemId, string? operatorUser, CancellationToken ct = default)
    {
        var inquiry = await _store.GetInquiryByIdAsync(inquiryId, ct);
        if (inquiry is null) return null;

        var contact = await _store.GetContactAsync(inquiry.CustomerPhone, ct);
        if (contact?.OptOut == true)
            throw new InvalidOperationException("Contact has opted out — outbound suppressed.");
        if (!_sms.IsConfigured)
            throw new InvalidOperationException("SMS gateway not configured.");

        var now = DateTimeOffset.UtcNow.ToString("o");
        var sid = await _sms.SendAsync(inquiry.CustomerPhone, body, StatusCallbackUrl(), ct);

        var msg = new MessageRecord
        {
            From           = _sms.FromNumber ?? "",
            To             = inquiry.CustomerPhone,
            Channel        = "sms",
            Direction      = "out",
            Body           = body,
            ConversationId = MessagingService.SmsConvId(inquiry.CustomerPhone),
            TimestampUtc   = now,
            IsRead         = true,
            ExternalId     = sid,
            InquiryId      = inquiry.Id,
            Status         = sid is null ? "failed" : "queued",
        };
        await _messages.AppendAsync(msg, ct);

        inquiry.LastMessageAt = now;   // outbound advances the thread but never adds unread
        inquiry.UpdatedAt     = now;
        inquiry.AwaitingReply = false; // we've replied — no longer owe the customer
        // Auto-assign on first response: the first person to reply (or claim) owns it; stealable later.
        if (string.IsNullOrWhiteSpace(inquiry.AssignedTo) && !string.IsNullOrWhiteSpace(operatorUser))
            inquiry.AssignedTo = operatorUser;
        await _store.UpdateInquiryAsync(inquiry, ct);

        if (fromDraftSpItemId is int dsid)
        {
            try { await _store.UpdateDraftStatusAsync(dsid, DraftStatus.Used, ct); }
            catch (Exception ex) { _log.LogWarning(ex, "[Inquiry] mark draft {Id} Used failed", dsid); }
        }

        _notify.NotifyInquiry("Updated", inquiry);
        _notify.NotifyInquiryMessage(inquiry.Id, msg);
        _log.LogInformation("[Inquiry] outbound reply on {Id} by {User} (sid={Sid})", inquiry.Id, operatorUser ?? "?", sid);
        return msg;
    }

    /// <summary>Accepts an AI draft: sends its body to the customer and marks it Used. Returns null when the
    /// inquiry/draft isn't found.</summary>
    public async Task<MessageRecord?> AcceptDraftAsync(string inquiryId, int draftSpItemId, string? operatorUser, CancellationToken ct = default)
    {
        var drafts = await _store.GetDraftsByInquiryAsync(inquiryId, ct);
        var draft  = drafts.FirstOrDefault(d => d.SpItemId == draftSpItemId);
        if (draft is null) return null;
        return await SendOperatorReplyAsync(inquiryId, draft.Body, draftSpItemId, operatorUser, ct);
    }

    public Task DismissDraftAsync(int draftSpItemId, CancellationToken ct = default)
        => _store.UpdateDraftStatusAsync(draftSpItemId, DraftStatus.Dismissed, ct);

    public async Task<InquiryNote?> AddNoteAsync(string inquiryId, string author, string body, CancellationToken ct = default)
    {
        if (await _store.GetInquiryByIdAsync(inquiryId, ct) is null) return null;
        var note = new InquiryNote { InquiryId = inquiryId, Author = author, Body = body, CreatedAt = DateTimeOffset.UtcNow.ToString("o") };
        await _store.CreateNoteAsync(note, ct);
        return note;
    }

    /// <summary>Links an HSK# quotation to the inquiry (deduped per inquiry) and advances a non-closed
    /// inquiry to Quoted.</summary>
    public async Task<InquiryQuotation?> LinkQuotationAsync(string inquiryId, string hskNumber, string linkedBy, CancellationToken ct = default)
    {
        var inquiry = await _store.GetInquiryByIdAsync(inquiryId, ct);
        if (inquiry is null) return null;

        var hsk = hskNumber.Trim();
        var existing = await _store.GetQuotationsByInquiryAsync(inquiryId, ct);
        var quotation = new InquiryQuotation
        {
            InquiryId = inquiryId, HskNumber = hsk,
            LinkedAt  = DateTimeOffset.UtcNow.ToString("o"), LinkedBy = linkedBy,
        };
        if (existing.Any(e => string.Equals(e.HskNumber, hsk, StringComparison.OrdinalIgnoreCase)))
            return quotation;   // already linked — idempotent

        await _store.CreateQuotationAsync(quotation, ct);
        if (!string.Equals(inquiry.Status, InquiryStatus.Closed, StringComparison.OrdinalIgnoreCase) &&
            !string.Equals(inquiry.Status, InquiryStatus.Quoted, StringComparison.OrdinalIgnoreCase))
        {
            inquiry.Status    = InquiryStatus.Quoted;
            inquiry.UpdatedAt = DateTimeOffset.UtcNow.ToString("o");
            await _store.UpdateInquiryAsync(inquiry, ct);
            _notify.NotifyInquiry("Updated", inquiry);
        }
        return quotation;
    }

    public async Task<Inquiry?> UpdateInquiryAsync(string inquiryId, string? status, string? assignedTo, CancellationToken ct = default)
    {
        var inquiry = await _store.GetInquiryByIdAsync(inquiryId, ct);
        if (inquiry is null) return null;
        if (status is not null)     inquiry.Status     = status;
        if (assignedTo is not null) inquiry.AssignedTo = assignedTo.Length == 0 ? null : assignedTo;
        inquiry.UpdatedAt = DateTimeOffset.UtcNow.ToString("o");
        await _store.UpdateInquiryAsync(inquiry, ct);
        _notify.NotifyInquiry("Updated", inquiry);
        return inquiry;
    }

    public async Task<Inquiry?> MarkReadAsync(string inquiryId, CancellationToken ct = default)
    {
        var inquiry = await _store.GetInquiryByIdAsync(inquiryId, ct);
        if (inquiry is null) return null;
        if (inquiry.UnreadCount != 0)
        {
            inquiry.UnreadCount = 0;
            inquiry.UpdatedAt   = DateTimeOffset.UtcNow.ToString("o");
            await _store.UpdateInquiryAsync(inquiry, ct);
            _notify.NotifyInquiry("Updated", inquiry);
        }
        return inquiry;
    }

    private string? StatusCallbackUrl()
    {
        var b = _config["SignalWire:WebhookBaseUrl"];
        return string.IsNullOrWhiteSpace(b) ? null : b.TrimEnd('/') + "/api/sms/status";
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
