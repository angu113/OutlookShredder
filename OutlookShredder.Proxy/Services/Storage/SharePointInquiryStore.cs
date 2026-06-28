using System.Text.Json;
using OutlookShredder.Proxy.Models;
using Graph = Microsoft.Graph.Models;

namespace OutlookShredder.Proxy.Services.Storage;

/// <summary>
/// SharePoint-backed <see cref="IInquiryStore"/>: the <c>Inquiries</c> + <c>MessagingContacts</c> lists. Owns
/// all the inquiry-specific Graph code so it stays out of the <see cref="SharePointService"/> connection
/// class; it reuses that connection only for the shared primitives (Graph client, site id, paginated read).
/// Columns that are $filtered/$ordered are indexed AT construction (per the SP rule — a brand-new column
/// can't be observed first, and SP 500s on an unindexed $filter even on a tiny list).
///
/// To port to Azure SQL: implement <see cref="IInquiryStore"/> against SQL and swap the DI registration —
/// nothing else in the pipeline changes.
/// </summary>
public sealed class SharePointInquiryStore : IInquiryStore
{
    private readonly SharePointService             _sp;
    private readonly ILogger<SharePointInquiryStore> _log;

    private string? _inquiriesListId;
    private string? _contactsListId;
    private string? _draftsListId;
    private string? _notesListId;
    private string? _quotationsListId;

    public SharePointInquiryStore(SharePointService sp, ILogger<SharePointInquiryStore> log)
    {
        _sp  = sp;
        _log = log;
    }

    public async Task EnsureProvisionedAsync(CancellationToken ct = default)
    {
        await GetInquiriesListIdAsync(ct);
        await GetContactsListIdAsync(ct);
        await GetDraftsListIdAsync(ct);
        await GetNotesListIdAsync(ct);
        await GetQuotationsListIdAsync(ct);
    }

    // ── Inquiries ─────────────────────────────────────────────────────────────

    private async Task<string> GetInquiriesListIdAsync(CancellationToken ct)
    {
        if (_inquiriesListId is not null) return _inquiriesListId;

        var graph  = _sp.GetGraph();
        var siteId = await _sp.GetSiteIdAsync();
        const string listName = "Inquiries";

        var lists = await graph.Sites[siteId].Lists
            .GetAsync(r => r.QueryParameters.Filter = $"displayName eq '{listName}'", ct);

        string listId;
        if (lists?.Value?.FirstOrDefault()?.Id is string eid)
        {
            listId = eid;
        }
        else
        {
            _log.LogInformation("[Inquiry] Creating Inquiries list");
            var created = await graph.Sites[siteId].Lists.PostAsync(new Graph.List
            {
                DisplayName = listName,
                ListProp    = new Graph.ListInfo { Template = "genericList" },
            }, cancellationToken: ct);
            listId = created?.Id ?? throw new Exception("Failed to create Inquiries list");
            _log.LogInformation("[Inquiry] Created Inquiries list -> {Id}", listId);
        }

        var existing = await graph.Sites[siteId].Lists[listId].Columns.GetAsync(cancellationToken: ct);
        var have = existing?.Value?.Select(c => c.Name ?? "").ToHashSet(StringComparer.OrdinalIgnoreCase) ?? [];

        (string Name, string Type)[] cols =
        [
            ("CinqId",        "text"),    // = Title; the indexed lookup key for collision checks
            ("CustomerPhone", "text"),
            ("IqStatus",      "text"),
            ("AssignedTo",    "text"),
            ("CustomerName",  "text"),
            ("ContactName",   "text"),
            ("CreatedAt",     "text"),
            ("UpdatedAt",     "text"),
            ("LastMessageAt", "text"),
            ("UnreadCount",   "number"),
            ("AwaitingReply", "boolean"),
        ];
        foreach (var (name, type) in cols)
        {
            if (have.Contains(name)) continue;
            try
            {
                var col = type switch
                {
                    "number"  => new Graph.ColumnDefinition { Name = name, Number  = new Graph.NumberColumn() },
                    "boolean" => new Graph.ColumnDefinition { Name = name, Boolean = new Graph.BooleanColumn() },
                    _         => new Graph.ColumnDefinition { Name = name, Text    = new Graph.TextColumn() },
                };
                await graph.Sites[siteId].Lists[listId].Columns.PostAsync(col, cancellationToken: ct);
                _log.LogInformation("[Inquiry] Created Inquiries column '{Name}'", name);
            }
            catch (Exception ex) { _log.LogWarning(ex, "[Inquiry] Failed to create Inquiries column '{Name}'", name); }
        }

        var allCols = await graph.Sites[siteId].Lists[listId].Columns.GetAsync(cancellationToken: ct);
        foreach (var col in allCols?.Value ?? [])
        {
            if (col.Name is not ("CinqId" or "CustomerPhone" or "LastMessageAt" or "IqStatus")) continue;
            if (col.Indexed == true || col.Id is null) continue;
            try
            {
                await graph.Sites[siteId].Lists[listId].Columns[col.Id]
                    .PatchAsync(new Graph.ColumnDefinition { Indexed = true }, cancellationToken: ct);
                _log.LogInformation("[Inquiry] Indexed Inquiries column '{Name}'", col.Name);
            }
            catch (Exception ex) { _log.LogWarning(ex, "[Inquiry] Failed to index Inquiries column '{Name}'", col.Name); }
        }

        return _inquiriesListId = listId;
    }

    private static Graph.ListItem InquiryToItem(Inquiry inq) => new()
    {
        Fields = new Graph.FieldValueSet
        {
            AdditionalData = new Dictionary<string, object?>
            {
                ["Title"]         = inq.Id,
                ["CinqId"]        = inq.Id,
                ["CustomerPhone"] = inq.CustomerPhone,
                ["IqStatus"]      = inq.Status,
                ["AssignedTo"]    = inq.AssignedTo,
                ["CustomerName"]  = inq.CustomerName,
                ["ContactName"]   = inq.ContactName,
                ["CreatedAt"]     = inq.CreatedAt,
                ["UpdatedAt"]     = inq.UpdatedAt,
                ["LastMessageAt"] = inq.LastMessageAt,
                ["UnreadCount"]   = (double)inq.UnreadCount,
                ["AwaitingReply"] = inq.AwaitingReply,
            },
        },
    };

    private static Inquiry ItemToInquiry(Graph.ListItem item)
    {
        var d = item.Fields?.AdditionalData ?? new Dictionary<string, object?>();
        string S(string k) => d.TryGetValue(k, out var v) ? v?.ToString() ?? "" : "";
        bool B(string k) => d.TryGetValue(k, out var v) && v is not null &&
            (v is JsonElement je ? je.ValueKind == JsonValueKind.True : v is bool b && b);
        return new Inquiry
        {
            SpItemId      = int.TryParse(item.Id, out var id) ? id : null,
            Id            = S("CinqId") is { Length: > 0 } cinq ? cinq : S("Title"),
            CustomerPhone = S("CustomerPhone"),
            Status        = S("IqStatus") is { Length: > 0 } st ? st : InquiryStatus.Open,
            AssignedTo    = S("AssignedTo") is { Length: > 0 } at ? at : null,
            CustomerName  = S("CustomerName") is { Length: > 0 } cn ? cn : null,
            ContactName   = S("ContactName") is { Length: > 0 } con ? con : null,
            CreatedAt     = S("CreatedAt"),
            UpdatedAt     = S("UpdatedAt"),
            LastMessageAt = S("LastMessageAt"),
            UnreadCount   = ReadInt(d, "UnreadCount") ?? 0,
            AwaitingReply = B("AwaitingReply"),
        };
    }

    private static int? ReadInt(IDictionary<string, object?> d, string key)
    {
        if (!d.TryGetValue(key, out var v) || v is null) return null;
        return v switch
        {
            JsonElement je when je.ValueKind == JsonValueKind.Number => (int)je.GetDouble(),
            double dd  => (int)dd,
            decimal dm => (int)dm,
            long l     => (int)l,
            int i      => i,
            _ => int.TryParse(v.ToString(), out var p) ? p : null,
        };
    }

    public async Task<int> CreateInquiryAsync(Inquiry inquiry, CancellationToken ct = default)
    {
        var graph  = _sp.GetGraph();
        var siteId = await _sp.GetSiteIdAsync();
        var listId = await GetInquiriesListIdAsync(ct);
        var item   = await graph.Sites[siteId].Lists[listId].Items.PostAsync(InquiryToItem(inquiry), cancellationToken: ct);
        if (item?.Id is null || !int.TryParse(item.Id, out var id))
            throw new Exception("Inquiries write returned no item ID");
        inquiry.SpItemId = id;
        return id;
    }

    public async Task UpdateInquiryAsync(Inquiry inquiry, CancellationToken ct = default)
    {
        if (inquiry.SpItemId is null) return;
        var graph  = _sp.GetGraph();
        var siteId = await _sp.GetSiteIdAsync();
        var listId = await GetInquiriesListIdAsync(ct);
        var fields = new Dictionary<string, object?>
        {
            ["IqStatus"]      = inquiry.Status,
            ["AssignedTo"]    = inquiry.AssignedTo,
            ["CustomerName"]  = inquiry.CustomerName,
            ["ContactName"]   = inquiry.ContactName,
            ["UpdatedAt"]     = inquiry.UpdatedAt,
            ["LastMessageAt"] = inquiry.LastMessageAt,
            ["UnreadCount"]   = (double)inquiry.UnreadCount,
            ["AwaitingReply"] = inquiry.AwaitingReply,
        };
        await graph.Sites[siteId].Lists[listId].Items[inquiry.SpItemId.Value.ToString()].Fields
            .PatchAsync(new Graph.FieldValueSet { AdditionalData = fields }, cancellationToken: ct);
    }

    public async Task<IReadOnlyList<Inquiry>> GetInquiriesByPhoneAsync(string phone, CancellationToken ct = default)
    {
        var listId = await GetInquiriesListIdAsync(ct);
        var items  = await _sp.ReadAllListItemsAsync(listId, expand: ["fields"],
            filter: $"fields/CustomerPhone eq '{phone.Replace("'", "''")}'",
            orderby: ["fields/LastMessageAt desc"], ct: ct);
        return items.Select(ItemToInquiry).ToList();
    }

    public async Task<IReadOnlyList<Inquiry>> GetInquiriesAsync(string? status = null, string? query = null, CancellationToken ct = default)
    {
        var listId = await GetInquiriesListIdAsync(ct);
        // Status is indexed → filter server-side when given; free-text query is a small in-memory pass.
        var filter = string.IsNullOrWhiteSpace(status) ? null : $"fields/IqStatus eq '{status.Replace("'", "''")}'";
        var items  = await _sp.ReadAllListItemsAsync(listId, expand: ["fields"],
            filter: filter, orderby: ["fields/LastMessageAt desc"], ct: ct);
        var all = items.Select(ItemToInquiry);
        if (!string.IsNullOrWhiteSpace(query))
        {
            var q = query.Trim();
            all = all.Where(i =>
                i.CustomerPhone.Contains(q, StringComparison.OrdinalIgnoreCase) ||
                i.Id.Contains(q, StringComparison.OrdinalIgnoreCase));
        }
        return all.ToList();
    }

    public async Task<Inquiry?> GetInquiryByIdAsync(string cinqId, CancellationToken ct = default)
    {
        var graph  = _sp.GetGraph();
        var siteId = await _sp.GetSiteIdAsync();
        var listId = await GetInquiriesListIdAsync(ct);
        var match  = await graph.Sites[siteId].Lists[listId].Items.GetAsync(r =>
        {
            r.QueryParameters.Expand = ["fields"];
            r.QueryParameters.Filter = $"fields/CinqId eq '{cinqId.Replace("'", "''")}'";
            r.QueryParameters.Top    = 1;
        }, ct);
        var item = match?.Value?.FirstOrDefault();
        return item is null ? null : ItemToInquiry(item);
    }

    // ── MessagingContacts ─────────────────────────────────────────────────────

    private async Task<string> GetContactsListIdAsync(CancellationToken ct)
    {
        if (_contactsListId is not null) return _contactsListId;

        var graph  = _sp.GetGraph();
        var siteId = await _sp.GetSiteIdAsync();
        const string listName = "MessagingContacts";

        var lists = await graph.Sites[siteId].Lists
            .GetAsync(r => r.QueryParameters.Filter = $"displayName eq '{listName}'", ct);

        string listId;
        if (lists?.Value?.FirstOrDefault()?.Id is string eid)
        {
            listId = eid;
        }
        else
        {
            _log.LogInformation("[Inquiry] Creating MessagingContacts list");
            var created = await graph.Sites[siteId].Lists.PostAsync(new Graph.List
            {
                DisplayName = listName,
                ListProp    = new Graph.ListInfo { Template = "genericList" },
            }, cancellationToken: ct);
            listId = created?.Id ?? throw new Exception("Failed to create MessagingContacts list");
            _log.LogInformation("[Inquiry] Created MessagingContacts list -> {Id}", listId);
        }

        var existing = await graph.Sites[siteId].Lists[listId].Columns.GetAsync(cancellationToken: ct);
        var have = existing?.Value?.Select(c => c.Name ?? "").ToHashSet(StringComparer.OrdinalIgnoreCase) ?? [];

        (string Name, string Type)[] cols =
        [
            ("Phone",             "text"),   // = Title (the contact key)
            ("DisplayName",       "text"),
            ("ConsentCapturedAt", "text"),
            ("ConsentMethod",     "text"),
            ("OptOut",            "boolean"),
            ("OptOutAt",          "text"),
        ];
        foreach (var (name, type) in cols)
        {
            if (have.Contains(name)) continue;
            try
            {
                var col = type switch
                {
                    "boolean" => new Graph.ColumnDefinition { Name = name, Boolean = new Graph.BooleanColumn() },
                    _         => new Graph.ColumnDefinition { Name = name, Text    = new Graph.TextColumn() },
                };
                await graph.Sites[siteId].Lists[listId].Columns.PostAsync(col, cancellationToken: ct);
                _log.LogInformation("[Inquiry] Created MessagingContacts column '{Name}'", name);
            }
            catch (Exception ex) { _log.LogWarning(ex, "[Inquiry] Failed to create MessagingContacts column '{Name}'", name); }
        }

        var allCols = await graph.Sites[siteId].Lists[listId].Columns.GetAsync(cancellationToken: ct);
        foreach (var col in allCols?.Value ?? [])
        {
            if (col.Name != "Phone") continue;
            if (col.Indexed == true || col.Id is null) continue;
            try
            {
                await graph.Sites[siteId].Lists[listId].Columns[col.Id]
                    .PatchAsync(new Graph.ColumnDefinition { Indexed = true }, cancellationToken: ct);
                _log.LogInformation("[Inquiry] Indexed MessagingContacts column 'Phone'");
            }
            catch (Exception ex) { _log.LogWarning(ex, "[Inquiry] Failed to index MessagingContacts column 'Phone'"); }
        }

        return _contactsListId = listId;
    }

    public async Task<MessagingContact?> GetContactAsync(string phone, CancellationToken ct = default)
    {
        var graph  = _sp.GetGraph();
        var siteId = await _sp.GetSiteIdAsync();
        var listId = await GetContactsListIdAsync(ct);
        var match  = await graph.Sites[siteId].Lists[listId].Items.GetAsync(r =>
        {
            r.QueryParameters.Expand = ["fields"];
            r.QueryParameters.Filter = $"fields/Phone eq '{phone.Replace("'", "''")}'";
            r.QueryParameters.Top    = 1;
        }, ct);
        var item = match?.Value?.FirstOrDefault();
        if (item is null) return null;

        var d = item.Fields?.AdditionalData ?? new Dictionary<string, object?>();
        string S(string k) => d.TryGetValue(k, out var v) ? v?.ToString() ?? "" : "";
        bool B(string k) => d.TryGetValue(k, out var v) && v is not null &&
            (v is JsonElement je ? je.ValueKind == JsonValueKind.True : v is bool b && b);
        return new MessagingContact
        {
            SpItemId          = int.TryParse(item.Id, out var id) ? id : null,
            Phone             = S("Phone") is { Length: > 0 } p ? p : S("Title"),
            DisplayName       = S("DisplayName") is { Length: > 0 } dn ? dn : null,
            ConsentCapturedAt = S("ConsentCapturedAt") is { Length: > 0 } cc ? cc : null,
            ConsentMethod     = S("ConsentMethod") is { Length: > 0 } cm ? cm : null,
            OptOut            = B("OptOut"),
            OptOutAt          = S("OptOutAt") is { Length: > 0 } oa ? oa : null,
        };
    }

    public async Task<int> UpsertContactAsync(MessagingContact contact, CancellationToken ct = default)
    {
        var graph  = _sp.GetGraph();
        var siteId = await _sp.GetSiteIdAsync();
        var listId = await GetContactsListIdAsync(ct);
        var fields = new Dictionary<string, object?>
        {
            ["Title"]             = contact.Phone,
            ["Phone"]             = contact.Phone,
            ["DisplayName"]       = contact.DisplayName,
            ["ConsentCapturedAt"] = contact.ConsentCapturedAt,
            ["ConsentMethod"]     = contact.ConsentMethod,
            ["OptOut"]            = contact.OptOut,
            ["OptOutAt"]          = contact.OptOutAt,
        };

        if (contact.SpItemId is int existingId)
        {
            await graph.Sites[siteId].Lists[listId].Items[existingId.ToString()].Fields
                .PatchAsync(new Graph.FieldValueSet { AdditionalData = fields }, cancellationToken: ct);
            return existingId;
        }

        var created = await graph.Sites[siteId].Lists[listId].Items
            .PostAsync(new Graph.ListItem { Fields = new Graph.FieldValueSet { AdditionalData = fields } }, cancellationToken: ct);
        if (created?.Id is null || !int.TryParse(created.Id, out var id))
            throw new Exception("MessagingContacts write returned no item ID");
        contact.SpItemId = id;
        return id;
    }

    // ── Drafts ────────────────────────────────────────────────────────────────

    private async Task<string> GetDraftsListIdAsync(CancellationToken ct)
    {
        if (_draftsListId is not null) return _draftsListId;

        var graph  = _sp.GetGraph();
        var siteId = await _sp.GetSiteIdAsync();
        const string listName = "Drafts";

        var lists = await graph.Sites[siteId].Lists
            .GetAsync(r => r.QueryParameters.Filter = $"displayName eq '{listName}'", ct);

        string listId;
        if (lists?.Value?.FirstOrDefault()?.Id is string eid)
        {
            listId = eid;
        }
        else
        {
            _log.LogInformation("[Inquiry] Creating Drafts list");
            var created = await graph.Sites[siteId].Lists.PostAsync(new Graph.List
            {
                DisplayName = listName,
                ListProp    = new Graph.ListInfo { Template = "genericList" },
            }, cancellationToken: ct);
            listId = created?.Id ?? throw new Exception("Failed to create Drafts list");
            _log.LogInformation("[Inquiry] Created Drafts list -> {Id}", listId);
        }

        var existing = await graph.Sites[siteId].Lists[listId].Columns.GetAsync(cancellationToken: ct);
        var have = existing?.Value?.Select(c => c.Name ?? "").ToHashSet(StringComparer.OrdinalIgnoreCase) ?? [];

        (string Name, string Type)[] cols =
        [
            ("InquiryId",           "text"),
            ("TriggeringMessageId", "text"),
            ("DraftSource",         "text"),
            ("TemplateId",          "text"),
            ("Body",                "note"),
            ("SuggestedIntent",     "text"),
            ("SuggestedUrgency",    "text"),
            ("NeedsQuote",          "boolean"),
            ("OptionsJson",         "note"),
            ("DraftStatus",         "text"),
            ("CreatedAt",           "text"),
        ];
        foreach (var (name, type) in cols)
        {
            if (have.Contains(name)) continue;
            try
            {
                var col = type switch
                {
                    "boolean" => new Graph.ColumnDefinition { Name = name, Boolean = new Graph.BooleanColumn() },
                    "note"    => new Graph.ColumnDefinition { Name = name, Text    = new Graph.TextColumn { AllowMultipleLines = true } },
                    _         => new Graph.ColumnDefinition { Name = name, Text    = new Graph.TextColumn() },
                };
                await graph.Sites[siteId].Lists[listId].Columns.PostAsync(col, cancellationToken: ct);
                _log.LogInformation("[Inquiry] Created Drafts column '{Name}'", name);
            }
            catch (Exception ex) { _log.LogWarning(ex, "[Inquiry] Failed to create Drafts column '{Name}'", name); }
        }

        var allCols = await graph.Sites[siteId].Lists[listId].Columns.GetAsync(cancellationToken: ct);
        foreach (var col in allCols?.Value ?? [])
        {
            if (col.Name is not ("InquiryId" or "CreatedAt")) continue;
            if (col.Indexed == true || col.Id is null) continue;
            try
            {
                await graph.Sites[siteId].Lists[listId].Columns[col.Id]
                    .PatchAsync(new Graph.ColumnDefinition { Indexed = true }, cancellationToken: ct);
                _log.LogInformation("[Inquiry] Indexed Drafts column '{Col}'", col.Name);
            }
            catch (Exception ex) { _log.LogWarning(ex, "[Inquiry] Failed to index Drafts column '{Col}'", col.Name); }
        }

        return _draftsListId = listId;
    }

    public async Task<int> CreateDraftAsync(InquiryDraft draft, CancellationToken ct = default)
    {
        var graph  = _sp.GetGraph();
        var siteId = await _sp.GetSiteIdAsync();
        var listId = await GetDraftsListIdAsync(ct);
        var item   = new Graph.ListItem
        {
            Fields = new Graph.FieldValueSet
            {
                AdditionalData = new Dictionary<string, object?>
                {
                    ["Title"]               = draft.InquiryId,
                    ["InquiryId"]           = draft.InquiryId,
                    ["TriggeringMessageId"] = draft.TriggeringMessageId,
                    ["DraftSource"]         = draft.Source,
                    ["TemplateId"]          = draft.TemplateId,
                    ["Body"]                = draft.Body,
                    ["SuggestedIntent"]     = draft.SuggestedIntent,
                    ["SuggestedUrgency"]    = draft.SuggestedUrgency,
                    ["NeedsQuote"]          = draft.NeedsQuote,
                    ["OptionsJson"]         = draft.OptionsJson,
                    ["DraftStatus"]         = draft.Status,
                    ["CreatedAt"]           = draft.CreatedAt,
                },
            },
        };
        var created = await graph.Sites[siteId].Lists[listId].Items.PostAsync(item, cancellationToken: ct);
        if (created?.Id is null || !int.TryParse(created.Id, out var id))
            throw new Exception("Drafts write returned no item ID");
        draft.SpItemId = id;
        return id;
    }

    public async Task<IReadOnlyList<InquiryDraft>> GetDraftsByInquiryAsync(string inquiryId, CancellationToken ct = default)
    {
        var listId = await GetDraftsListIdAsync(ct);
        var items  = await _sp.ReadAllListItemsAsync(listId, expand: ["fields"],
            filter: $"fields/InquiryId eq '{inquiryId.Replace("'", "''")}'",
            orderby: ["fields/CreatedAt desc"], ct: ct);

        var result = new List<InquiryDraft>();
        foreach (var item in items)
        {
            var d = item.Fields?.AdditionalData ?? new Dictionary<string, object?>();
            string S(string k) => d.TryGetValue(k, out var v) ? v?.ToString() ?? "" : "";
            bool B(string k) => d.TryGetValue(k, out var v) && v is not null &&
                (v is JsonElement je ? je.ValueKind == JsonValueKind.True : v is bool b && b);
            result.Add(new InquiryDraft
            {
                SpItemId            = int.TryParse(item.Id, out var id) ? id : null,
                InquiryId           = S("InquiryId"),
                TriggeringMessageId = S("TriggeringMessageId") is { Length: > 0 } tm ? tm : null,
                Source              = S("DraftSource") is { Length: > 0 } src ? src : DraftSource.Ai,
                TemplateId          = S("TemplateId") is { Length: > 0 } ti ? ti : null,
                Body                = S("Body"),
                SuggestedIntent     = S("SuggestedIntent") is { Length: > 0 } si ? si : null,
                SuggestedUrgency    = S("SuggestedUrgency") is { Length: > 0 } su ? su : null,
                NeedsQuote          = B("NeedsQuote"),
                OptionsJson         = S("OptionsJson") is { Length: > 0 } oj ? oj : null,
                Status              = S("DraftStatus") is { Length: > 0 } st ? st : DraftStatus.Pending,
                CreatedAt           = S("CreatedAt"),
            });
        }
        return result;
    }

    public async Task UpdateDraftStatusAsync(int spItemId, string status, CancellationToken ct = default)
    {
        var graph  = _sp.GetGraph();
        var siteId = await _sp.GetSiteIdAsync();
        var listId = await GetDraftsListIdAsync(ct);
        await graph.Sites[siteId].Lists[listId].Items[spItemId.ToString()].Fields
            .PatchAsync(new Graph.FieldValueSet { AdditionalData = new Dictionary<string, object?> { ["DraftStatus"] = status } },
                cancellationToken: ct);
    }

    // ── Inquiry child lists (Notes, Quotations) — same shape: a genericList keyed + indexed on InquiryId ──

    /// <summary>Gets/creates a simple inquiry child list with the given extra columns, indexing InquiryId.</summary>
    private async Task<string> GetChildListIdAsync(string listName, (string Name, string Type)[] cols, string[] indexCols, CancellationToken ct)
    {
        var graph  = _sp.GetGraph();
        var siteId = await _sp.GetSiteIdAsync();

        var lists = await graph.Sites[siteId].Lists
            .GetAsync(r => r.QueryParameters.Filter = $"displayName eq '{listName}'", ct);

        string listId;
        if (lists?.Value?.FirstOrDefault()?.Id is string eid)
        {
            listId = eid;
        }
        else
        {
            _log.LogInformation("[Inquiry] Creating {List} list", listName);
            var created = await graph.Sites[siteId].Lists.PostAsync(new Graph.List
            {
                DisplayName = listName,
                ListProp    = new Graph.ListInfo { Template = "genericList" },
            }, cancellationToken: ct);
            listId = created?.Id ?? throw new Exception($"Failed to create {listName} list");
            _log.LogInformation("[Inquiry] Created {List} list -> {Id}", listName, listId);
        }

        var existing = await graph.Sites[siteId].Lists[listId].Columns.GetAsync(cancellationToken: ct);
        var have = existing?.Value?.Select(c => c.Name ?? "").ToHashSet(StringComparer.OrdinalIgnoreCase) ?? [];
        foreach (var (name, type) in cols)
        {
            if (have.Contains(name)) continue;
            try
            {
                var col = type switch
                {
                    "note" => new Graph.ColumnDefinition { Name = name, Text = new Graph.TextColumn { AllowMultipleLines = true } },
                    _      => new Graph.ColumnDefinition { Name = name, Text = new Graph.TextColumn() },
                };
                await graph.Sites[siteId].Lists[listId].Columns.PostAsync(col, cancellationToken: ct);
                _log.LogInformation("[Inquiry] Created {List} column '{Name}'", listName, name);
            }
            catch (Exception ex) { _log.LogWarning(ex, "[Inquiry] Failed to create {List} column '{Name}'", listName, name); }
        }

        var allCols = await graph.Sites[siteId].Lists[listId].Columns.GetAsync(cancellationToken: ct);
        foreach (var col in allCols?.Value ?? [])
        {
            if (col.Name is null || !indexCols.Contains(col.Name)) continue;
            if (col.Indexed == true || col.Id is null) continue;
            try
            {
                await graph.Sites[siteId].Lists[listId].Columns[col.Id]
                    .PatchAsync(new Graph.ColumnDefinition { Indexed = true }, cancellationToken: ct);
                _log.LogInformation("[Inquiry] Indexed {List} column '{Col}'", listName, col.Name);
            }
            catch (Exception ex) { _log.LogWarning(ex, "[Inquiry] Failed to index {List} column '{Col}'", listName, col.Name); }
        }
        return listId;
    }

    private async Task<string> GetNotesListIdAsync(CancellationToken ct)
        // "NoteAuthor" not "Author" — Author is a RESERVED SharePoint internal field (built-in "Created By"),
        // so a text column named Author silently collides with it and writes to it fail (500).
        => _notesListId ??= await GetChildListIdAsync("InquiryNotes",
            [("InquiryId", "text"), ("NoteAuthor", "text"), ("Body", "note"), ("CreatedAt", "text")],
            ["InquiryId", "CreatedAt"], ct);

    private async Task<string> GetQuotationsListIdAsync(CancellationToken ct)
        => _quotationsListId ??= await GetChildListIdAsync("InquiryQuotations",
            [("InquiryId", "text"), ("HskNumber", "text"), ("LinkedAt", "text"), ("LinkedBy", "text")],
            ["InquiryId", "LinkedAt"], ct);

    public async Task<int> CreateNoteAsync(InquiryNote note, CancellationToken ct = default)
    {
        var graph  = _sp.GetGraph();
        var siteId = await _sp.GetSiteIdAsync();
        var listId = await GetNotesListIdAsync(ct);
        var item = new Graph.ListItem
        {
            Fields = new Graph.FieldValueSet
            {
                AdditionalData = new Dictionary<string, object?>
                {
                    ["Title"]      = note.InquiryId,
                    ["InquiryId"]  = note.InquiryId,
                    ["NoteAuthor"] = note.Author,
                    ["Body"]       = note.Body,
                    ["CreatedAt"]  = note.CreatedAt,
                },
            },
        };
        var created = await graph.Sites[siteId].Lists[listId].Items.PostAsync(item, cancellationToken: ct);
        if (created?.Id is null || !int.TryParse(created.Id, out var id)) throw new Exception("InquiryNotes write returned no item ID");
        note.SpItemId = id;
        return id;
    }

    public async Task<IReadOnlyList<InquiryNote>> GetNotesByInquiryAsync(string inquiryId, CancellationToken ct = default)
    {
        var listId = await GetNotesListIdAsync(ct);
        var items  = await _sp.ReadAllListItemsAsync(listId, expand: ["fields"],
            filter: $"fields/InquiryId eq '{inquiryId.Replace("'", "''")}'",
            orderby: ["fields/CreatedAt asc"], ct: ct);
        return items.Select(item =>
        {
            var d = item.Fields?.AdditionalData ?? new Dictionary<string, object?>();
            string S(string k) => d.TryGetValue(k, out var v) ? v?.ToString() ?? "" : "";
            return new InquiryNote
            {
                SpItemId  = int.TryParse(item.Id, out var id) ? id : null,
                InquiryId = S("InquiryId"),
                Author    = S("NoteAuthor"),
                Body      = S("Body"),
                CreatedAt = S("CreatedAt"),
            };
        }).ToList();
    }

    public async Task<int> CreateQuotationAsync(InquiryQuotation quotation, CancellationToken ct = default)
    {
        var graph  = _sp.GetGraph();
        var siteId = await _sp.GetSiteIdAsync();
        var listId = await GetQuotationsListIdAsync(ct);
        var item = new Graph.ListItem
        {
            Fields = new Graph.FieldValueSet
            {
                AdditionalData = new Dictionary<string, object?>
                {
                    ["Title"]     = quotation.HskNumber,
                    ["InquiryId"] = quotation.InquiryId,
                    ["HskNumber"] = quotation.HskNumber,
                    ["LinkedAt"]  = quotation.LinkedAt,
                    ["LinkedBy"]  = quotation.LinkedBy,
                },
            },
        };
        var created = await graph.Sites[siteId].Lists[listId].Items.PostAsync(item, cancellationToken: ct);
        if (created?.Id is null || !int.TryParse(created.Id, out var id)) throw new Exception("InquiryQuotations write returned no item ID");
        quotation.SpItemId = id;
        return id;
    }

    public async Task<IReadOnlyList<InquiryQuotation>> GetQuotationsByInquiryAsync(string inquiryId, CancellationToken ct = default)
    {
        var listId = await GetQuotationsListIdAsync(ct);
        var items  = await _sp.ReadAllListItemsAsync(listId, expand: ["fields"],
            filter: $"fields/InquiryId eq '{inquiryId.Replace("'", "''")}'",
            orderby: ["fields/LinkedAt asc"], ct: ct);
        return items.Select(item =>
        {
            var d = item.Fields?.AdditionalData ?? new Dictionary<string, object?>();
            string S(string k) => d.TryGetValue(k, out var v) ? v?.ToString() ?? "" : "";
            return new InquiryQuotation
            {
                SpItemId  = int.TryParse(item.Id, out var id) ? id : null,
                InquiryId = S("InquiryId"),
                HskNumber = S("HskNumber"),
                LinkedAt  = S("LinkedAt"),
                LinkedBy  = S("LinkedBy"),
            };
        }).ToList();
    }
}
