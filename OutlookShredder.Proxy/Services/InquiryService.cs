using System.Text;
using System.Text.Json;
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
    private readonly OutboundMediaBlobService _blob;   // ephemeral MMS media egress (Blob+SAS)
    private readonly PdfRasterService        _pdf;     // PDF -> per-page JPEGs for MMS
    private readonly InquiryDraftService     _drafts;
    private readonly CustomerCacheService    _crm;
    private readonly SmsInquiryCacheService  _cache;
    private readonly ProductCatalogService   _catalog;
    private readonly IConfiguration          _config;
    private readonly ILogger<InquiryService> _log;

    // Default copy (override per key in appsettings). Straight apostrophes only — see DefaultOptInReply note.
    private const string DefaultHelpReply =
        "Mithril Metals Corp., Authorized Metal Supermarkets Franchisee (Hackensack). For assistance, call " +
        "(201) 957-7955 or email hackensack@metalsupermarkets.com. Reply STOP to unsubscribe. Msg & data rates may apply.";
    private const string DefaultOptOutReply =
        "Mithril Metals Corp., Authorized Metal Supermarkets Franchisee (Hackensack). You have been successfully " +
        "unsubscribed and will no longer receive SMS messages from us. Reply START to resubscribe.";
    // Straight apostrophe (not curly) keeps the message GSM-7, not UCS-2 — UCS-2 cuts the segment size to 67
    // chars and would turn this ~290-char message into ~5 billable segments instead of ~2.
    private const string DefaultOptInReply =
        "Mithril Metals Corp., Authorized Metal Supermarkets Franchisee (Hackensack). You're now subscribed to " +
        "receive SMS replies regarding quotes, orders, and store inquiries. Message frequency varies based on your " +
        "inquiries. Msg & data rates may apply. Reply HELP for help, STOP to unsubscribe.";

    public InquiryService(IInquiryStore store, IMessageStore messages, RfqNotificationService notify,
        ISmsGateway sms, OutboundMediaBlobService blob, PdfRasterService pdf, InquiryDraftService drafts,
        CustomerCacheService crm, SmsInquiryCacheService cache, ProductCatalogService catalog,
        IConfiguration config, ILogger<InquiryService> log)
    {
        _store    = store;
        _messages = messages;
        _notify   = notify;
        _sms      = sms;
        _blob     = blob;
        _pdf      = pdf;
        _drafts   = drafts;
        _cache    = cache;
        _catalog  = catalog;
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
    public async Task IngestInboundAsync(string from, string to, string body, string? sid, string? mediaJson = null,
        CancellationToken ct = default)
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

        var inquiries        = await _store.GetInquiriesByPhoneAsync(phone, ct);
        var (action, latest) = InquiryRules.DecideThread(inquiries);   // latest = most-recent ACTIVE inquiry (null if none)

        // 2. Carrier keyword (STOP/HELP/START family): a compliance/info signal, not a sales question — do
        //    NOT mint a CINQ or bump unread. Record it against the existing thread (if any) for audit, and
        //    answer HELP unless opted out.
        if (keyword != InquiryRules.Keyword.None)
        {
            var kwMsg = await AppendMessageAsync(from, to, body, sid, latest?.Id, now, null, ct);
            await SendComplianceReplyAsync(keyword, from, contact.OptOut, ct);
            if (latest is not null) { _notify.NotifyInquiryMessage(latest.Id, kwMsg); _ = _cache.RefreshOneAsync(latest.Id); }
            _log.LogInformation("[Inquiry] {Keyword} from {Phone} (optOut={OptOut})", keyword, phone, contact.OptOut);
            return;
        }

        // 3. Normal customer message → thread it. Resolve the customer from CRM (denormalised for the list +
        //    "first-time caller" detection); inbound always leaves us owing a reply (AwaitingReply).
        var crm = _crm.LookupByPhone(from);
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
        else   // Append (active/Spam thread) OR Reopen (all the customer's threads were Closed) — continue the one thread
        {
            inquiry = latest!;
            if (action == InquiryRules.ThreadAction.Reopen)
                inquiry.Status = InquiryStatus.Open;   // a returning customer reopens their closed thread (one thread per customer)
            inquiry.LastMessageAt = now;
            inquiry.UpdatedAt     = now;
            inquiry.UnreadCount  += 1;
            inquiry.AwaitingReply = true;
            inquiry.CustomerName ??= crm?.BusinessPartner;   // backfill if not resolved before
            inquiry.ContactName  ??= crm?.ContactName;
            await _store.UpdateInquiryAsync(inquiry, ct);
        }

        // Inbound media (MMS attachments; email attachments later): download from the carrier, store durably,
        // promote a text/plain caption to an empty body, and collect image/PDF parts for the AI draft.
        var (effectiveBody, mediaStored, aiAttachments) =
            await ProcessInboundMediaAsync(inquiry.Id, body, sid, mediaJson, ct);

        var msg = await AppendMessageAsync(from, to, effectiveBody, sid, inquiry.Id, now, mediaStored, ct);

        _notify.NotifyInquiry(isNew ? "Created" : "Updated", inquiry);
        _notify.NotifyInquiryMessage(inquiry.Id, msg);
        _ = _cache.RefreshOneAsync(inquiry.Id);   // full populate (contact + new message); off the read path
        _log.LogInformation("[Inquiry] {Action} {Id} from {Phone} (unread={Unread}, media={Media})",
            action, inquiry.Id, phone, inquiry.UnreadCount, msg.Media.Count);

        // Phase 2: AI reply suggestion — async, never auto-sent, and detached from the queue consumer's
        // token so a slow Claude call neither blocks ingest nor is cancelled when the SB message completes.
        _ = GenerateDraftAsync(inquiry.Id, effectiveBody, sid, aiAttachments);
    }

    /// <summary>Builds + persists an AI reply suggestion for the inquiry and pushes it live. Fire-and-forget:
    /// it owns its own timeout and swallows all errors (a draft is a non-critical suggestion).</summary>
    private async Task GenerateDraftAsync(string inquiryId, string inboundBody, string? triggeringSid,
        IReadOnlyList<DraftAttachment>? attachments = null)
    {
        try
        {
            using var cts = new CancellationTokenSource(TimeSpan.FromMinutes(2));
            var ct = cts.Token;

            // Prior transcript = the thread minus the just-appended inbound (passed separately as "latest").
            var history    = await _messages.GetByInquiryAsync(inquiryId, 12, ct);
            var prior      = history.Count > 0 ? history.Take(history.Count - 1).ToList() : history;
            var transcript = InquiryDraftPrompt.BuildTranscript(prior);

            // Phase 6: reuse the RFQ catalog token-matcher for the product heavy-lifting — feed the AI the
            // closest catalog families so its clarifier compares against real products (not prompt heuristics).
            // The terse-dim expansion turns "2 box"/"2 angle" into a square/equal cross-section for the match.
            string? catalogContext = null;
            try
            {
                var candidates = _catalog.TopCandidates(InquiryRules.ExpandTerseDims(inboundBody), 6);
                if (candidates.Count > 0)
                    catalogContext = string.Join("\n", candidates.Select(c => "- " + c.Name));
            }
            catch (Exception ex) { _log.LogWarning(ex, "[Inquiry] catalog match failed for {Id}", inquiryId); }

            // Linked HSK# / notes arrive in Phase 3 (quotation linking + notes) — empty for now. Image/PDF
            // attachments (when present) are fed to the model so it can read a sketch / spec sheet.
            var result = await _drafts.DraftAsync(
                new InquiryDraftInput(inboundBody, transcript, Array.Empty<string>(), null, attachments, catalogContext), ct);
            if (result is null) return;

            // Each new message re-clarifies with the updated conversation context, so this fresh suggestion
            // SUPERSEDES any prior pending draft — only the latest stays Pending (the cached suggestion always
            // reflects the most recent message, moving the conversation toward a quote).
            foreach (var d in (await _store.GetDraftsByInquiryAsync(inquiryId, ct))
                              .Where(x => string.Equals(x.Status, DraftStatus.Pending, StringComparison.OrdinalIgnoreCase) && x.SpItemId is int))
            {
                try { await _store.UpdateDraftStatusAsync(d.SpItemId!.Value, DraftStatus.Dismissed, ct);
                      _cache.SetDraftStatus(inquiryId, d.SpItemId!.Value, DraftStatus.Dismissed); }
                catch (Exception ex) { _log.LogWarning(ex, "[Inquiry] supersede prior draft {Id} failed", d.SpItemId); }
            }

            var draft = new InquiryDraft
            {
                InquiryId           = inquiryId,
                TriggeringMessageId = triggeringSid,
                Source              = DraftSource.Ai,
                Body                = result.Reply,
                SuggestedIntent     = result.Intent,
                SuggestedUrgency    = result.Urgency,
                NeedsQuote          = result.NeedsQuote,
                OptionsJson         = result.Options.Count > 0 ? JsonSerializer.Serialize(result.Options) : null,
                Status              = DraftStatus.Pending,
                CreatedAt           = DateTimeOffset.UtcNow.ToString("o"),
            };
            await _store.CreateDraftAsync(draft, ct);
            _cache.ApplyDraft(draft);
            _notify.NotifyInquiryDraft(draft);
            _log.LogInformation("[Inquiry] AI draft for {Id} (intent={Intent} urgency={Urgency} needsQuote={NeedsQuote})",
                inquiryId, draft.SuggestedIntent, draft.SuggestedUrgency, draft.NeedsQuote);
        }
        catch (Exception ex) { _log.LogWarning(ex, "[Inquiry] AI draft generation failed for {Id}", inquiryId); }
    }

    /// <summary>Updates an outbound message's delivery status by SID (SignalWire status callback).</summary>
    public Task<bool> UpdateMessageStatusAsync(string sid, string status, CancellationToken ct = default)
        => _messages.UpdateStatusBySidAsync(sid, status, ct);

    /// <summary>Operator-triggered: dismiss any prior pending drafts and generate a FRESH AI suggestion for the
    /// latest inbound message (so a stale/early draft can be replaced on demand). False if no inbound exists.</summary>
    public async Task<bool> RegenerateDraftAsync(string inquiryId, CancellationToken ct = default)
    {
        var messages    = await _messages.GetByInquiryAsync(inquiryId, 12, ct);
        var lastInbound = messages.LastOrDefault(m => string.Equals(m.Direction, "in", StringComparison.OrdinalIgnoreCase));
        if (lastInbound is null) return false;
        await GenerateDraftAsync(inquiryId, lastInbound.Body, lastInbound.ExternalId, null);  // supersedes prior pending
        return true;
    }

    private async Task<MessageRecord> AppendMessageAsync(
        string from, string to, string body, string? sid, string? inquiryId, string now,
        string? mediaJson, CancellationToken ct)
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
            MediaJson      = mediaJson,
        };
        try { await _messages.AppendAsync(msg, ct); }
        catch (Exception ex) { _log.LogWarning(ex, "[Inquiry] message append failed for {From}", from); }
        return msg;
    }

    private static readonly JsonSerializerOptions _mediaJsonOpts = new() { PropertyNameCaseInsensitive = true };

    /// <summary>The descriptor the SMS webhook (Azure Function) forwards for each inbound media part.</summary>
    private sealed record InboundMediaPart(string? Url, string? ContentType);

    /// <summary>An operator-attached outbound file (image or PDF) to send as MMS.</summary>
    public sealed record OutboundAttachment(string FileName, string ContentType, byte[] Bytes);

    /// <summary>Content types SignalWire MMS accepts as images (a PDF is rasterized to these first).</summary>
    private static readonly HashSet<string> _mmsImageTypes = new(StringComparer.OrdinalIgnoreCase)
        { "image/jpeg", "image/jpg", "image/png", "image/gif", "image/bmp" };

    /// <summary>Stores one attachment's bytes durably under the inquiry (SharePoint InquiryMedia) and records a
    /// <see cref="MessageMedia"/> reference — the permanent copy the app renders, independent of the blob/SAS.</summary>
    private async Task SaveDurableAsync(string inquiryId, string name, byte[] bytes, string contentType, string kind,
        List<MessageMedia> into, CancellationToken ct)
    {
        await _messages.SaveMediaAsync(inquiryId, name, bytes, ct);
        into.Add(new MessageMedia { Name = name, ContentType = contentType, Kind = kind });
    }

    /// <summary>Greedily packs outbound media into MMS-sized batches (&lt;=8 items and ~1 MB each — the carrier
    /// caps a message near 1.2 MB local). Yields (batch, isFirst) so the text body rides only the first send.</summary>
    private static IEnumerable<(List<(string Name, string ContentType, byte[] Bytes)> Batch, bool IsFirst)>
        PackForMms(List<(string Name, string ContentType, byte[] Bytes)> items)
    {
        const int  MaxItems = 8;
        const long MaxBytes = 1_000_000;
        var  cur      = new List<(string Name, string ContentType, byte[] Bytes)>();
        long curBytes = 0;
        var  first    = true;
        foreach (var it in items)
        {
            if (cur.Count > 0 && (cur.Count >= MaxItems || curBytes + it.Bytes.Length > MaxBytes))
            {
                yield return (cur, first); first = false;
                cur = new(); curBytes = 0;
            }
            cur.Add(it); curBytes += it.Bytes.Length;
        }
        if (cur.Count > 0) yield return (cur, first);
    }

    /// <summary>
    /// Downloads each inbound media part from the carrier and stores the binary parts durably under the
    /// inquiry (so previews + AI survive the carrier's short retention). Returns the effective body (a
    /// text/plain part is promoted when the body is empty — MMS delivers the caption as media), the stored
    /// media JSON for the message row, and the image/PDF attachments to feed the AI draft. SMIL/HTML layout
    /// parts are dropped. Best-effort: a failed part is skipped, never fatal to ingest.
    /// </summary>
    private async Task<(string Body, string? MediaJson, IReadOnlyList<DraftAttachment> AiAttachments)>
        ProcessInboundMediaAsync(string inquiryId, string body, string? sid, string? mediaJson, CancellationToken ct)
    {
        if (string.IsNullOrWhiteSpace(mediaJson)) return (body, null, []);

        List<InboundMediaPart>? parts;
        try { parts = JsonSerializer.Deserialize<List<InboundMediaPart>>(mediaJson, _mediaJsonOpts); }
        catch (Exception ex) { _log.LogWarning(ex, "[Inquiry] bad media payload"); return (body, null, []); }
        if (parts is null || parts.Count == 0) return (body, null, []);

        var stored   = new List<MessageMedia>();
        var ai       = new List<DraftAttachment>();
        string? caption = null;
        var index = 0;

        foreach (var part in parts)
        {
            var i = index++;
            if (string.IsNullOrWhiteSpace(part.Url)) continue;
            var kind = InquiryRules.ClassifyMedia(part.ContentType, null);
            if (kind == InquiryRules.MediaKind.Ignore) continue;

            var dl = await _sms.DownloadMediaAsync(part.Url!, ct);
            if (dl is null) continue;
            var (servedType, bytes) = dl.Value;
            var contentType = string.IsNullOrWhiteSpace(part.ContentType) ? servedType : part.ContentType!;

            if (kind == InquiryRules.MediaKind.Caption)
            {
                try { var txt = Encoding.UTF8.GetString(bytes).Trim(); if (caption is null && txt.Length > 0) caption = txt; }
                catch { /* not decodable as text — skip */ }
                continue;
            }

            var ext  = InquiryRules.ExtForContentType(contentType);
            var name = $"{(string.IsNullOrEmpty(sid) ? "msg" : sid)}-{i}.{ext}";
            try { await _messages.SaveMediaAsync(inquiryId, name, bytes, ct); }
            catch (Exception ex) { _log.LogWarning(ex, "[Inquiry] media store failed for {Name}", name); continue; }

            var kindStr = kind switch
            {
                InquiryRules.MediaKind.Image => "image",
                InquiryRules.MediaKind.Pdf   => "pdf",
                InquiryRules.MediaKind.Cad   => "cad",
                _                            => "file",
            };
            stored.Add(new MessageMedia { Name = name, ContentType = contentType, Kind = kindStr });

            // Feed only vision-readable parts to the model, and only mime types it accepts.
            var mime = kind == InquiryRules.MediaKind.Pdf ? "application/pdf" : contentType.ToLowerInvariant();
            if (mime == "image/jpg") mime = "image/jpeg";
            if (kind == InquiryRules.MediaKind.Pdf || mime is "image/jpeg" or "image/png" or "image/gif" or "image/webp")
                ai.Add(new DraftAttachment(mime, Convert.ToBase64String(bytes), name));
        }

        var effectiveBody = !string.IsNullOrWhiteSpace(body) ? body : (caption ?? "");
        var storedJson    = stored.Count > 0 ? JsonSerializer.Serialize(stored, _mediaJsonOpts) : null;
        return (effectiveBody, storedJson, ai);
    }

    /// <summary>Serves a stored inquiry media file (preview / download) from the local disk cache (SP fallback
    /// inside the cache). Name is an opaque key we minted.</summary>
    public async Task<(string ContentType, byte[] Bytes)?> GetMediaAsync(string inquiryId, string fileName, CancellationToken ct = default)
        => InquiryRules.IsSafeMediaName(fileName) ? await _cache.GetMediaAsync(inquiryId, fileName, ct) : null;

    // ── Operator-initiated "New SMS" (start a conversation to a chosen number) ──────────────────────

    /// <summary>Canonical US E.164 key ("+1" + 10 digits) matching what inbound stores, or null if the input
    /// isn't a 10-digit US number. (NormalizeE164 alone wouldn't re-add the country code.)</summary>
    private static string? NormalizeUsE164(string? raw)
    {
        var ten = CustomerImportService.NormalizePhone(raw);
        return ten is null ? null : "+1" + ten;
    }

    /// <summary>The customer's existing one-thread inquiry id for this (raw) phone, or null if none. Read-only —
    /// lets the client open an existing conversation instead of starting a duplicate.</summary>
    public async Task<string?> FindInquiryIdByPhoneAsync(string rawPhone, CancellationToken ct = default)
    {
        var e164 = NormalizeUsE164(rawPhone);
        if (e164 is null) return null;
        var inquiries = await _store.GetInquiriesByPhoneAsync(e164, ct);
        return InquiryRules.DecideThread(inquiries).Target?.Id;
    }

    /// <summary>Get-or-create-or-reopen the one-thread inquiry for a number (operator-initiated New SMS). Resolves
    /// CRM for the denormalized name, reopens a Closed thread, ensures a MessagingContact, and warms the cache.
    /// Returns the inquiry, or null when the phone isn't a valid US number. Creating happens here (on first send).</summary>
    public async Task<Inquiry?> StartInquiryAsync(string rawPhone, CancellationToken ct = default)
    {
        var e164 = NormalizeUsE164(rawPhone);
        if (e164 is null) return null;
        var now = DateTimeOffset.UtcNow.ToString("o");

        var inquiries        = await _store.GetInquiriesByPhoneAsync(e164, ct);
        var (action, latest) = InquiryRules.DecideThread(inquiries);
        var crm              = _crm.LookupByPhone(e164);

        if (action != InquiryRules.ThreadAction.CreateNew && latest is not null)
        {
            if (action == InquiryRules.ThreadAction.Reopen)   // returning customer's closed thread
            {
                latest.Status    = InquiryStatus.Open;
                latest.UpdatedAt = now;
                await _store.UpdateInquiryAsync(latest, ct);
            }
            _cache.ApplyInquiry(latest);
            _notify.NotifyInquiry("Updated", latest);
            return latest;
        }

        var contact = await _store.GetContactAsync(e164, ct)
                      ?? new MessagingContact { Phone = e164, ConsentCapturedAt = now, ConsentMethod = "operator-initiated" };
        try { await _store.UpsertContactAsync(contact, ct); }
        catch (Exception ex) { _log.LogWarning(ex, "[Inquiry] contact upsert failed for {Phone}", e164); }

        var inquiry = new Inquiry
        {
            Id            = await GenerateCinqIdAsync(ct),
            CustomerPhone = e164,
            Status        = InquiryStatus.Open,
            CustomerName  = crm?.BusinessPartner,
            ContactName   = crm?.ContactName,
            CreatedAt     = now,
            UpdatedAt     = now,
            LastMessageAt = now,
            UnreadCount   = 0,       // we initiated — nothing unread from the customer
            AwaitingReply = false,   // and we're not waiting on our own reply
        };
        await _store.CreateInquiryAsync(inquiry, ct);
        try { await _cache.RefreshOneAsync(inquiry.Id, ct); } catch { /* the send + detail load re-warm it */ }
        _notify.NotifyInquiry("Created", inquiry);
        _log.LogInformation("[Inquiry] operator-started {Id} for {Phone}", inquiry.Id, e164);
        return inquiry;
    }

    /// <summary>Media backfill/recovery: re-runs media processing for a known message SID against the supplied
    /// carrier media descriptors and patches the existing row's body + media. Repairs messages that ingested
    /// before media handling existed (or whose media download failed). False if no row matched the SID.</summary>
    public async Task<bool> BackfillMessageMediaAsync(string inquiryId, string sid, string mediaJson, CancellationToken ct = default)
    {
        var (body, stored, _) = await ProcessInboundMediaAsync(inquiryId, "", sid, mediaJson, ct);
        var ok = await _messages.PatchBodyMediaBySidAsync(sid, body, stored, ct);
        if (ok) _log.LogInformation("[Inquiry] backfilled media for sid {Sid} on {Id}", sid, inquiryId);
        return ok;
    }

    /// <summary>Sends the carrier-style confirmation for a STOP/START/HELP keyword. APP-OWNED while
    /// <c>SignalWire:AppHelpReply</c> is true (the default) — when SignalWire's 10DLC campaign is registered to
    /// auto-send the templates, set it false to hand compliance back to the carrier (avoids double-texting).
    /// Sent directly via the gateway (not the opt-out-suppressed reply path) because the STOP confirmation is
    /// the one message a carrier permits after opt-out; a HELP to an already-opted-out number is suppressed.</summary>
    private async Task SendComplianceReplyAsync(InquiryRules.Keyword keyword, string to, bool optedOut, CancellationToken ct)
    {
        if (!_config.GetValue("SignalWire:AppHelpReply", true)) return;   // SignalWire's campaign owns it when false
        if (keyword == InquiryRules.Keyword.Help && optedOut) return;      // they asked us to stop — stay silent
        if (!_sms.IsConfigured) { _log.LogWarning("[Inquiry] {Kw} received but SMS gateway not configured", keyword); return; }

        var reply = keyword switch
        {
            InquiryRules.Keyword.Help   => _config["SignalWire:HelpReply"]   is { Length: > 0 } h ? h : DefaultHelpReply,
            InquiryRules.Keyword.OptOut => _config["SignalWire:OptOutReply"] is { Length: > 0 } o ? o : DefaultOptOutReply,
            InquiryRules.Keyword.OptIn  => _config["SignalWire:OptInReply"]  is { Length: > 0 } i ? i : DefaultOptInReply,
            _ => null,
        };
        if (reply is null) return;

        try { await _sms.SendAsync(to, reply, ct: ct); }
        catch (Exception ex) { _log.LogWarning(ex, "[Inquiry] {Kw} confirmation failed to {To}", keyword, to); }
    }

    // ── Phase 3 operator actions (called by InquiriesController) ──────────────────────────────────

    // Reads are served from the in-memory cache (active inquiries; Closed/Spam fall through to SP inside it).
    public Task<IReadOnlyList<Inquiry>> ListAsync(string? status, string? query, CancellationToken ct = default)
        => _cache.ListAsync(status, query, ct);

    public Task<InquiryDetail?> GetDetailAsync(string inquiryId, CancellationToken ct = default)
        => _cache.GetDetailAsync(inquiryId, ct);

    /// <summary>Sends an operator reply to the customer (suppressed if opted out), records the outbound
    /// message, advances the inquiry, optionally marks the source draft Used, and pushes live updates.
    /// Throws <see cref="InvalidOperationException"/> when the contact opted out or no gateway is configured.</summary>
    public async Task<MessageRecord?> SendOperatorReplyAsync(
        string inquiryId, string body, int? fromDraftSpItemId, string? operatorUser,
        IReadOnlyList<OutboundAttachment>? attachments = null, CancellationToken ct = default)
    {
        var inquiry = await _store.GetInquiryByIdAsync(inquiryId, ct);
        if (inquiry is null) return null;

        var contact = await _store.GetContactAsync(inquiry.CustomerPhone, ct);
        if (contact?.OptOut == true)
            throw new InvalidOperationException("Contact has opted out — outbound suppressed.");
        if (!_sms.IsConfigured)
            throw new InvalidOperationException("SMS gateway not configured.");

        var hasAttachments = attachments is { Count: > 0 };
        if (hasAttachments && !_blob.IsConfigured)
            throw new InvalidOperationException("Sending images/PDFs requires MMS, which isn't configured yet.");

        // Attachments: store the DURABLE copy in SharePoint (the permanent record Pulse renders — never lost)
        // and collect the image bytes to actually send over MMS. A PDF is rasterized to per-page JPEGs (MMS
        // rejects application/pdf); the source PDF is also kept durably. See project_outbound_mms_durability.
        var durable = new List<MessageMedia>();
        var toSend  = new List<(string Name, string ContentType, byte[] Bytes)>();
        if (hasAttachments)
        {
            var stamp = DateTimeOffset.UtcNow.ToUnixTimeMilliseconds();
            var idx   = 0;
            foreach (var att in attachments!)
            {
                var isPdf = (att.ContentType?.Contains("pdf", StringComparison.OrdinalIgnoreCase) ?? false)
                            || att.FileName.EndsWith(".pdf", StringComparison.OrdinalIgnoreCase);
                if (isPdf)
                {
                    var renderedPages = _pdf.RenderToJpegs(att.Bytes);
                    if (renderedPages.Count == 0)
                        throw new InvalidOperationException("Couldn't render that PDF into images to send.");
                    await SaveDurableAsync(inquiry.Id, $"{stamp}-{idx}-src.pdf", att.Bytes, "application/pdf", "pdf", durable, ct);
                    foreach (var pg in renderedPages)
                    {
                        var pageName = $"{stamp}-{idx}-p{pg.Number}.jpg";
                        await SaveDurableAsync(inquiry.Id, pageName, pg.Jpeg, "image/jpeg", "image", durable, ct);
                        toSend.Add((pageName, "image/jpeg", pg.Jpeg));
                    }
                }
                else if (_mmsImageTypes.Contains(att.ContentType ?? ""))
                {
                    var name = $"{stamp}-{idx}.{InquiryRules.ExtForContentType(att.ContentType!)}";
                    await SaveDurableAsync(inquiry.Id, name, att.Bytes, att.ContentType!, "image", durable, ct);
                    toSend.Add((name, att.ContentType!, att.Bytes));
                }
                else
                {
                    throw new InvalidOperationException($"Attachment type '{att.ContentType}' can't be sent over MMS (images or PDF only).");
                }
                idx++;
            }
        }

        var now      = DateTimeOffset.UtcNow.ToString("o");
        var statusCb = StatusCallbackUrl();
        string? sid  = null;

        if (toSend.Count == 0)
        {
            sid = await _sms.SendAsync(inquiry.CustomerPhone, body, statusCb, null, ct);
        }
        else
        {
            // Pack page/image media into MMS messages within the carrier budget (~1 MB total, <=8 items each);
            // the operator's text rides on the first message. Multi-page PDFs may span a few messages.
            foreach (var (batch, isFirst) in PackForMms(toSend))
            {
                var urls = new List<string>();
                foreach (var m in batch)
                {
                    try { urls.Add(await _blob.UploadAndGetSasUrlAsync(inquiry.Id, m.Name, m.Bytes, m.ContentType, ct)); }
                    catch (Exception ex) { _log.LogWarning(ex, "[Inquiry] MMS media upload failed for {Name}", m.Name); }
                }
                if (urls.Count == 0) continue;
                var s = await _sms.SendAsync(inquiry.CustomerPhone, isFirst ? body : "", statusCb, urls, ct);
                sid ??= s;
            }
        }

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
            MediaJson      = durable.Count > 0 ? JsonSerializer.Serialize(durable, _mediaJsonOpts) : null,
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

        _cache.ApplyInquiry(inquiry);
        _cache.ApplyMessage(inquiry.Id, msg);
        if (fromDraftSpItemId is int usedDraft) _cache.SetDraftStatus(inquiry.Id, usedDraft, DraftStatus.Used);
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
        return await SendOperatorReplyAsync(inquiryId, draft.Body, draftSpItemId, operatorUser, null, ct);
    }

    public async Task DismissDraftAsync(string inquiryId, int draftSpItemId, CancellationToken ct = default)
    {
        await _store.UpdateDraftStatusAsync(draftSpItemId, DraftStatus.Dismissed, ct);
        _cache.SetDraftStatus(inquiryId, draftSpItemId, DraftStatus.Dismissed);
    }

    public async Task<InquiryNote?> AddNoteAsync(string inquiryId, string author, string body, CancellationToken ct = default)
    {
        var inquiry = await _store.GetInquiryByIdAsync(inquiryId, ct);
        if (inquiry is null) return null;
        var note = new InquiryNote { InquiryId = inquiryId, Author = author, Body = body, CreatedAt = DateTimeOffset.UtcNow.ToString("o") };
        await _store.CreateNoteAsync(note, ct);
        _cache.ApplyNote(note);
        _notify.NotifyInquiry("Updated", inquiry);   // broadcast so other proxies refresh (and pick up the note)
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
        _cache.ApplyQuotation(quotation);
        if (!string.Equals(inquiry.Status, InquiryStatus.Closed, StringComparison.OrdinalIgnoreCase) &&
            !string.Equals(inquiry.Status, InquiryStatus.Quoted, StringComparison.OrdinalIgnoreCase))
        {
            inquiry.Status    = InquiryStatus.Quoted;
            inquiry.UpdatedAt = DateTimeOffset.UtcNow.ToString("o");
            await _store.UpdateInquiryAsync(inquiry, ct);
            _cache.ApplyInquiry(inquiry);
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
        _cache.ApplyInquiry(inquiry);   // evicts if status moved to Closed/Spam
        _notify.NotifyInquiry("Updated", inquiry);
        return inquiry;
    }

    /// <summary>One-time operator-identity backfill (Windows login -> Shredder username), rewriting AssignedTo /
    /// note author / quote linker. Dry-run by default; when applied it patches SP then refreshes the active-inquiry
    /// cache and pushes a live update per affected inquiry so peers + open clients pick up the rename.</summary>
    public async Task<IdentityBackfillResult> BackfillIdentityAsync(string fromName, string toName, bool apply, CancellationToken ct = default)
    {
        var result = await _store.BackfillIdentityAsync(fromName, toName, apply, ct);
        if (apply)
        {
            foreach (var id in result.AffectedInquiryIds)
            {
                try
                {
                    await _cache.RefreshOneAsync(id, ct);   // targeted SP re-read so notes/quotes pick up the rename
                    var inq = await _store.GetInquiryByIdAsync(id, ct);
                    if (inq is not null) _notify.NotifyInquiry("Updated", inq);
                }
                catch (Exception ex) { _log.LogWarning(ex, "[Inquiry] backfill cache refresh failed for {Id}", id); }
            }
        }
        return result;
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
            _cache.ApplyInquiry(inquiry);
            _notify.NotifyInquiry("Updated", inquiry);
        }
        return inquiry;
    }

    // ── Phase 7: unread state (button-only — no auto-read-on-open) ─────────────────────────────────
    /// <summary>Sets one message's read flag, recounts the inquiry's unread INBOUND total, and pushes it live.</summary>
    public async Task<Inquiry?> SetMessageReadAsync(string inquiryId, int messageSpItemId, bool read, CancellationToken ct = default)
    {
        var inquiry = await _store.GetInquiryByIdAsync(inquiryId, ct);
        if (inquiry is null) return null;
        await _messages.SetMessageReadAsync(messageSpItemId, read, ct);
        _cache.SetMessageRead(inquiryId, messageSpItemId, read);
        return await RecountUnreadAsync(inquiry, ct);
    }

    /// <summary>Marks every message in the inquiry read or unread (mark-all), then recounts + pushes live.</summary>
    public async Task<Inquiry?> MarkAllAsync(string inquiryId, bool read, CancellationToken ct = default)
    {
        var inquiry = await _store.GetInquiryByIdAsync(inquiryId, ct);
        if (inquiry is null) return null;
        await _messages.SetAllReadByInquiryAsync(inquiryId, read, ct);
        _cache.SetAllRead(inquiryId, read);
        return await RecountUnreadAsync(inquiry, ct);
    }

    /// <summary>Total unread inbound across all active inquiries — the app-level badge source (Phase 7b).</summary>
    public int UnreadTotal() => _cache.TotalUnread();

    private async Task<Inquiry> RecountUnreadAsync(Inquiry inquiry, CancellationToken ct)
    {
        var uc = _cache.UnreadCount(inquiry.Id);
        if (uc < 0)   // not cached (Closed/Spam) — count from the store
            uc = (await _messages.GetByInquiryAsync(inquiry.Id, 500, ct))
                 .Count(m => string.Equals(m.Direction, "in", StringComparison.OrdinalIgnoreCase) && !m.IsRead);
        inquiry.UnreadCount = uc;
        inquiry.UpdatedAt   = DateTimeOffset.UtcNow.ToString("o");
        await _store.UpdateInquiryAsync(inquiry, ct);
        _cache.ApplyInquiry(inquiry);
        _notify.NotifyInquiry("Updated", inquiry);
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
