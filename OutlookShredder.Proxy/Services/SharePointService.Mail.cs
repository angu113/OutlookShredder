using System.Text.Json;
using Microsoft.Graph.Models;
using OutlookShredder.Proxy.Models;

namespace OutlookShredder.Proxy.Services;

/// <summary>
/// Mail-workbench persistence (wip/mail-classification.md, Phase 1b). Two lists:
///   MailItems          — immutable captured email/document (never mutated by classification).
///   MailClassifications — versioned descriptive layer (FK MailItemId, IsCurrent flag, full history).
/// Lives on SharePointService (the single SP write layer) as a partial class so it reuses the
/// Graph client, site-id resolution, EnsureListColumnsAsync, and GetStr helpers.
/// </summary>
public partial class SharePointService
{
    private const string MailItemsList = "MailItems";
    private const string MailClassList = "MailClassifications";
    private const string MailHintsList = "MailTaxonomyHints";
    private const string MailProjectsList = "MailProjects";

    // ── Provisioning ──────────────────────────────────────────────────────────────

    public async Task<object> EnsureMailListsAsync()
    {
        var siteId = await GetSiteIdAsync();

        var items = await EnsureListColumnsAsync(siteId, MailItemsList,
        [
            ("MailItemId","text"), ("SourceType","text"), ("SourceMailbox","text"),
            ("InternetMessageId","text"), ("WrapperGraphId","text"), ("ConversationId","text"),
            ("FromAddress","text"), ("FromName","text"), ("ToLine","note"), ("CcLine","note"),
            ("EmailSubject","text"), ("ReceivedAt","dateTime"),
            ("BodyText","note"), ("HasAttachments","boolean"), ("AttachmentsJson","note"),
            ("RawEmlUrl","text"), ("CapturedAt","dateTime"), ("RefsJson","note"),
            ("Completed","boolean"), ("CompletedAt","dateTime"), ("CompletedBy","text"),
            ("IsRead","boolean"), ("ReadAt","dateTime"), ("ReadBy","text"),
            ("ClaimedBy","text"), ("ClaimedAt","dateTime"),
            ("Direction","text"),   // "in" (default) | "out" — outbound = workbench-composed self-BCC copies
        ]);
        await IndexListColumnsAsync(siteId, MailItemsList, "MailItemId", "WrapperGraphId", "ReceivedAt", "Completed", "IsRead", "ConversationId");

        var cls = await EnsureListColumnsAsync(siteId, MailClassList,
        [
            ("MailItemId","text"), ("Version","number"), ("IsCurrent","boolean"),
            ("CategoryPath","text"), ("OtherLabel","text"), ("SupplierName","text"), ("Confidence","number"),
            ("KeywordTags","note"), ("PoNumber","text"), ("SoNumber","text"), ("Amount","text"),
            ("SupplierReference","text"), ("PayLink","note"),
            ("Reasoning","note"), ("AiProvider","text"), ("AiModel","text"),
            ("RawAiResponse","note"), ("ClassifiedAt","dateTime"),
        ]);
        await IndexListColumnsAsync(siteId, MailClassList, "MailItemId", "IsCurrent", "CategoryPath");

        var hints = await EnsureListColumnsAsync(siteId, MailHintsList,
        [
            ("CategoryPath", "text"), ("Hint", "note"), ("Source", "text"), ("CreatedAt", "dateTime"),
        ]);
        await IndexListColumnsAsync(siteId, MailHintsList, "CategoryPath");

        var projects = await EnsureListColumnsAsync(siteId, MailProjectsList,
        [
            ("ProjectId","text"), ("Status","text"), ("RefsJson","note"), ("ConversationIdsJson","note"),
            ("CreatedAt","dateTime"), ("CreatedBy","text"),
        ]);
        await IndexListColumnsAsync(siteId, MailProjectsList, "ProjectId", "Status");

        return new { mailItems = items, mailClassifications = cls, mailTaxonomyHints = hints, mailProjects = projects };
    }

    // ── MailProjects (Layer 2: cross-conversation grouping) ──────────────────────────

    public async Task<List<MailProjectRow>> ReadProjectsAsync(CancellationToken ct = default)
    {
        var rows = await ReadAllListItemsAsync(MailProjectsList,
            ["Title","ProjectId","Status","RefsJson","ConversationIdsJson","CreatedAt","CreatedBy"], null, ct);
        return rows.Select(f => new MailProjectRow
        {
            SpId            = GetStr(f, "__spId") ?? "",
            ProjectId       = GetStr(f, "ProjectId") ?? "",
            Name            = GetStr(f, "Title") ?? "",
            Status          = GetStr(f, "Status") ?? "active",
            Refs            = SafeList(GetStr(f, "RefsJson")),
            ConversationIds = SafeList(GetStr(f, "ConversationIdsJson")),
            CreatedAt       = GetStr(f, "CreatedAt") ?? "",
            CreatedBy       = GetStr(f, "CreatedBy") ?? "",
        }).Where(p => p.ProjectId.Length > 0).ToList();
    }

    public async Task<string> WriteProjectAsync(MailProjectRow p, CancellationToken ct = default)
    {
        var siteId = await GetSiteIdAsync();
        var listId = await ResolveListIdAsync(MailProjectsList);
        if (string.IsNullOrEmpty(p.ProjectId)) p.ProjectId = Guid.NewGuid().ToString("N");
        await GetGraph().Sites[siteId].Lists[listId].Items.PostAsync(new ListItem
        {
            Fields = new FieldValueSet { AdditionalData = new Dictionary<string, object?>
            {
                ["Title"]               = Trunc(p.Name, 250),
                ["ProjectId"]           = p.ProjectId,
                ["Status"]              = string.IsNullOrEmpty(p.Status) ? "active" : p.Status,
                ["RefsJson"]            = JsonSerializer.Serialize(p.Refs),
                ["ConversationIdsJson"] = JsonSerializer.Serialize(p.ConversationIds),
                ["CreatedAt"]           = DateTimeOffset.UtcNow.ToString("o"),
                ["CreatedBy"]           = p.CreatedBy,
            } }
        }, cancellationToken: ct);
        return p.ProjectId;
    }

    /// <summary>Patches an existing project's name/status/conversation membership by ProjectId.</summary>
    public async Task<bool> UpdateProjectAsync(string projectId, string? name, string? status, List<string>? conversationIds, CancellationToken ct = default)
    {
        var siteId = await GetSiteIdAsync();
        var listId = await ResolveListIdAsync(MailProjectsList);
        var res = await GetGraph().Sites[siteId].Lists[listId].Items.GetAsync(req =>
        {
            req.QueryParameters.Expand = ["fields($select=ProjectId)"];
            req.QueryParameters.Filter = $"fields/ProjectId eq '{Esc(projectId)}'";
            req.QueryParameters.Top    = 1;
        }, ct);
        var spId = res?.Value?.FirstOrDefault()?.Id;
        if (spId is null) return false;
        var fields = new Dictionary<string, object?>();
        if (name is not null)   fields["Title"]  = Trunc(name, 250);
        if (status is not null) fields["Status"] = status;
        if (conversationIds is not null) fields["ConversationIdsJson"] = JsonSerializer.Serialize(conversationIds);
        if (fields.Count == 0) return true;
        await GetGraph().Sites[siteId].Lists[listId].Items[spId].Fields
            .PatchAsync(new FieldValueSet { AdditionalData = fields }, cancellationToken: ct);
        return true;
    }

    private static List<string> SafeList(string? json)
    {
        if (string.IsNullOrWhiteSpace(json)) return [];
        try { return JsonSerializer.Deserialize<List<string>>(json) ?? []; } catch { return []; }
    }

    /// <summary>Reads the live taxonomy hints (custom leaves + classification guidance) from SP.</summary>
    public async Task<List<TaxonomyHintRow>> ReadTaxonomyHintsAsync(CancellationToken ct = default)
    {
        var rows = await ReadAllListItemsAsync(MailHintsList, ["CategoryPath", "Hint", "Source", "CreatedAt"], null, ct);
        return rows.Select(f => new TaxonomyHintRow
        {
            CategoryPath = GetStr(f, "CategoryPath") ?? "",
            Hint         = GetStr(f, "Hint") ?? "",
            Source       = GetStr(f, "Source") ?? "",
            CreatedAt    = GetStr(f, "CreatedAt") ?? "",
        }).Where(h => h.CategoryPath.Length > 0).ToList();
    }

    /// <summary>Appends a taxonomy hint (a confirmed leaf + optional guidance) to SP.</summary>
    public async Task WriteTaxonomyHintAsync(string categoryPath, string? hint, string source, CancellationToken ct = default)
    {
        var siteId = await GetSiteIdAsync();
        var listId = await ResolveListIdAsync(MailHintsList);
        await GetGraph().Sites[siteId].Lists[listId].Items.PostAsync(new ListItem
        {
            Fields = new FieldValueSet
            {
                AdditionalData = new Dictionary<string, object?>
                {
                    ["Title"]        = categoryPath,
                    ["CategoryPath"] = categoryPath,
                    ["Hint"]         = Trunc(hint, 8000),
                    ["Source"]       = source,
                    ["CreatedAt"]    = DateTimeOffset.UtcNow.ToString("o"),
                }
            }
        }, cancellationToken: ct);
        _log.LogInformation("[MailWB] taxonomy hint added: {Path} ({Source})", categoryPath, source);
    }

    /// <summary>Deletes taxonomy-hint rows whose CategoryPath equals <paramref name="path"/> or sits under it
    /// (path + "/"). Returns the number removed. Used to retire a confirmed custom leaf.</summary>
    public async Task<int> DeleteTaxonomyHintsAsync(string path, CancellationToken ct = default)
    {
        var siteId = await GetSiteIdAsync();
        var listId = await ResolveListIdAsync(MailHintsList);
        var rows = await ReadAllListItemsAsync(MailHintsList, ["CategoryPath"], null, ct);
        var prefix = path.TrimEnd('/') + "/";
        var removed = 0;
        foreach (var f in rows)
        {
            var cp = GetStr(f, "CategoryPath") ?? "";
            if (!(string.Equals(cp, path, StringComparison.OrdinalIgnoreCase) || cp.StartsWith(prefix, StringComparison.OrdinalIgnoreCase))) continue;
            var spId = GetStr(f, "__spId");
            if (spId is null) continue;
            try { await GetGraph().Sites[siteId].Lists[listId].Items[spId].DeleteAsync(cancellationToken: ct); removed++; }
            catch (Exception ex) { _log.LogWarning(ex, "[MailWB] hint delete failed {Cp}", cp); }
        }
        return removed;
    }

    private async Task IndexListColumnsAsync(string siteId, string listName, params string[] colNames)
    {
        var listId = await ResolveListIdAsync(listName);
        var cols   = await GetGraph().Sites[siteId].Lists[listId].Columns.GetAsync();
        foreach (var n in colNames)
        {
            var c = cols?.Value?.FirstOrDefault(x => string.Equals(x.Name, n, StringComparison.OrdinalIgnoreCase));
            if (c?.Id is null || c.Indexed == true) continue;
            try
            {
                await GetGraph().Sites[siteId].Lists[listId].Columns[c.Id].PatchAsync(new ColumnDefinition { Indexed = true });
                _log.LogInformation("[SP] Indexed {List} column '{Col}'", listName, n);
            }
            catch (Exception ex) { _log.LogWarning(ex, "[SP] could not index {List}.{Col}", listName, n); }
        }
    }

    // ── MailItems write/read ────────────────────────────────────────────────────────

    /// <summary>
    /// Writes a new MailItem and returns its stable MailItemId. Does NOT dedup — the caller
    /// (MailWorkbenchService) dedups in memory via <see cref="GetExistingWrapperIdsAsync"/>. The SP
    /// `fields/WrapperGraphId eq` filter false-matches on these long, near-identical Graph ids (all
    /// messages in a folder share a ~140-char prefix), so it can't be used for dedup.
    /// </summary>
    public async Task<string> WriteMailItemAsync(MailItemInput input, CancellationToken ct = default)
    {
        var siteId = await GetSiteIdAsync();
        var listId = await ResolveListIdAsync(MailItemsList);

        var mailItemId = Guid.NewGuid().ToString("N");
        await GetGraph().Sites[siteId].Lists[listId].Items.PostAsync(new ListItem
        {
            Fields = new FieldValueSet
            {
                AdditionalData = new Dictionary<string, object?>
                {
                    ["Title"]             = Trunc(input.Subject, 250),
                    ["MailItemId"]        = mailItemId,
                    ["SourceType"]        = input.SourceType,
                    ["SourceMailbox"]     = input.SourceMailbox,
                    ["InternetMessageId"] = input.InternetMessageId,
                    ["WrapperGraphId"]    = input.WrapperGraphId,
                    ["ConversationId"]    = input.ConversationId,
                    ["RefsJson"]          = input.RefsJson,
                    ["FromAddress"]       = input.FromAddress,
                    ["FromName"]          = input.FromName,
                    ["ToLine"]            = Trunc(input.ToLine, 8000),
                    ["CcLine"]            = Trunc(input.CcLine, 8000),
                    ["EmailSubject"]      = Trunc(input.Subject, 250),
                    ["ReceivedAt"]        = input.ReceivedAtIso,
                    ["BodyText"]          = Trunc(input.BodyText, 100000),
                    ["HasAttachments"]    = input.HasAttachments,
                    ["AttachmentsJson"]   = input.AttachmentsJson,
                    ["CapturedAt"]        = DateTimeOffset.UtcNow.ToString("o"),
                    ["Completed"]         = false,
                    ["Direction"]         = string.IsNullOrWhiteSpace(input.Direction) ? "in" : input.Direction,
                }
            }
        }, cancellationToken: ct);

        _log.LogInformation("[MailWB] Captured MailItem {Id} from {From} subj='{Subj}'",
            mailItemId, input.FromAddress, Trunc(input.Subject, 60));
        return mailItemId;
    }

    /// <summary>All captured WrapperGraphIds (full stored values) for reliable in-memory dedup.</summary>
    public async Task<HashSet<string>> GetExistingWrapperIdsAsync(CancellationToken ct = default)
    {
        var rows = await ReadAllListItemsAsync(MailItemsList, ["WrapperGraphId"], null, ct);
        return rows.Select(f => GetStr(f, "WrapperGraphId") ?? "")
                   .Where(s => s.Length > 0)
                   .ToHashSet(StringComparer.Ordinal);
    }

    /// <summary>All captured embedded Internet Message-IDs — the primary dedup key (discards re-sends).</summary>
    public async Task<HashSet<string>> GetExistingMessageIdsAsync(CancellationToken ct = default)
    {
        var rows = await ReadAllListItemsAsync(MailItemsList, ["InternetMessageId"], null, ct);
        return rows.Select(f => GetStr(f, "InternetMessageId") ?? "")
                   .Where(s => s.Length > 0)
                   .ToHashSet(StringComparer.Ordinal);
    }

    public async Task<List<MailItemRow>> ReadMailItemsAsync(CancellationToken ct = default)
    {
        var rows = await ReadAllListItemsAsync(MailItemsList,
            ["MailItemId","SourceType","SourceMailbox","WrapperGraphId","ConversationId","RefsJson","FromAddress","FromName",
             "EmailSubject","ReceivedAt","HasAttachments","AttachmentsJson","Completed","CompletedAt","CompletedBy",
             "IsRead","ReadAt","ReadBy","ClaimedBy","ClaimedAt","Direction"], null, ct);
        return rows.Select(MapItemRow).ToList();
    }

    private static MailItemRow MapItemRow(Dictionary<string, object?> f) => new()
    {
        SpId           = GetStr(f, "__spId") ?? "",
        MailItemId     = GetStr(f, "MailItemId") ?? "",
        WrapperGraphId = GetStr(f, "WrapperGraphId") ?? "",
        ConversationId = GetStr(f, "ConversationId") ?? "",
        RefsJson       = GetStr(f, "RefsJson") ?? "",
        SourceType     = GetStr(f, "SourceType") ?? "email",
        SourceMailbox  = GetStr(f, "SourceMailbox") ?? "",
        FromAddress    = GetStr(f, "FromAddress") ?? "",
        FromName       = GetStr(f, "FromName") ?? "",
        Subject        = GetStr(f, "EmailSubject") ?? "",
        ReceivedAt     = GetStr(f, "ReceivedAt") ?? "",
        HasAttachments = GetBool(f, "HasAttachments"),
        AttachmentsJson= GetStr(f, "AttachmentsJson") ?? "",
        Completed      = GetBool(f, "Completed"),
        CompletedAt    = GetStr(f, "CompletedAt"),
        CompletedBy    = GetStr(f, "CompletedBy"),
        IsRead         = GetBool(f, "IsRead"),
        ReadAt         = GetStr(f, "ReadAt"),
        ReadBy         = GetStr(f, "ReadBy"),
        ClaimedBy      = GetStr(f, "ClaimedBy"),
        ClaimedAt      = GetStr(f, "ClaimedAt"),
        Direction      = GetStr(f, "Direction") ?? "in",
    };

    public async Task<bool> SetMailCompletedAsync(string mailItemId, bool completed, string? by, CancellationToken ct = default)
    {
        var (siteId, listId, spId) = await ResolveMailItemSpIdAsync(mailItemId, ct);
        if (spId is null) return false;
        await GetGraph().Sites[siteId].Lists[listId].Items[spId].Fields.PatchAsync(new FieldValueSet
        {
            AdditionalData = new Dictionary<string, object?>
            {
                ["Completed"]   = completed,
                ["CompletedAt"] = completed ? DateTimeOffset.UtcNow.ToString("o") : null,
                ["CompletedBy"] = completed ? by : null,
            }
        }, cancellationToken: ct);
        return true;
    }

    /// <summary>Sets the global read flag (read-by-anyone) on a MailItem.</summary>
    public async Task<bool> SetMailReadAsync(string mailItemId, bool read, string? by, CancellationToken ct = default)
    {
        var (siteId, listId, spId) = await ResolveMailItemSpIdAsync(mailItemId, ct);
        if (spId is null) return false;
        await GetGraph().Sites[siteId].Lists[listId].Items[spId].Fields.PatchAsync(new FieldValueSet
        {
            AdditionalData = new Dictionary<string, object?>
            {
                ["IsRead"] = read,
                ["ReadAt"] = read ? DateTimeOffset.UtcNow.ToString("o") : null,
                ["ReadBy"] = read ? by : null,
            }
        }, cancellationToken: ct);
        return true;
    }

    /// <summary>Patches a MailItem's header-derived fields (ReceivedAt + ConversationId + RefsJson) from the repair pass.</summary>
    public async Task<bool> UpdateMailDerivedAsync(string mailItemId, string receivedIso, string conversationId, string refsJson, CancellationToken ct = default)
    {
        var (siteId, listId, spId) = await ResolveMailItemSpIdAsync(mailItemId, ct);
        if (spId is null) return false;
        await GetGraph().Sites[siteId].Lists[listId].Items[spId].Fields.PatchAsync(
            new FieldValueSet { AdditionalData = new Dictionary<string, object?>
            {
                ["ReceivedAt"]     = receivedIso,
                ["ConversationId"] = conversationId,
                ["RefsJson"]       = refsJson,
            } }, cancellationToken: ct);
        return true;
    }

    /// <summary>All items' (MailItemId, RawEmlUrl, ReceivedAt) — for the header/refs repair pass.</summary>
    public async Task<List<(string MailItemId, string RawEmlUrl, string ReceivedAt)>> ReadMailEmlPathsAsync(CancellationToken ct = default)
    {
        var rows = await ReadAllListItemsAsync(MailItemsList, ["MailItemId", "RawEmlUrl", "ReceivedAt"], null, ct);
        return rows.Select(f => (GetStr(f, "MailItemId") ?? "", GetStr(f, "RawEmlUrl") ?? "", GetStr(f, "ReceivedAt") ?? ""))
                   .Where(t => t.Item1.Length > 0).ToList();
    }

    /// <summary>Sets/clears the claim (owner) on a MailItem. Empty claimedBy releases the claim.</summary>
    public async Task<bool> SetMailClaimAsync(string mailItemId, string? claimedBy, string? claimedAtIso, CancellationToken ct = default)
    {
        var (siteId, listId, spId) = await ResolveMailItemSpIdAsync(mailItemId, ct);
        if (spId is null) return false;
        var claimed = !string.IsNullOrWhiteSpace(claimedBy);
        await GetGraph().Sites[siteId].Lists[listId].Items[spId].Fields.PatchAsync(new FieldValueSet
        {
            AdditionalData = new Dictionary<string, object?>
            {
                ["ClaimedBy"] = claimed ? claimedBy : null,
                ["ClaimedAt"] = claimed ? (claimedAtIso ?? DateTimeOffset.UtcNow.ToString("o")) : null,
            }
        }, cancellationToken: ct);
        return true;
    }

    /// <summary>Reads one item with the body + .eml pointer + recipients (for the full viewer).</summary>
    public async Task<MailItemDetailRow?> ReadMailItemDetailAsync(string mailItemId, CancellationToken ct = default)
    {
        var rows = await ReadAllListItemsAsync(MailItemsList,
            ["MailItemId","SourceType","SourceMailbox","WrapperGraphId","ConversationId","FromAddress","FromName","ToLine","CcLine",
             "EmailSubject","ReceivedAt","BodyText","HasAttachments","AttachmentsJson","RawEmlUrl",
             "Completed","CompletedAt","CompletedBy","IsRead","ReadAt","ReadBy","ClaimedBy","ClaimedAt","Direction"],
            $"fields/MailItemId eq '{Esc(mailItemId)}'", ct);
        var f = rows.FirstOrDefault();
        if (f is null) return null;
        var baseRow = MapItemRow(f);
        return new MailItemDetailRow
        {
            SpId = baseRow.SpId, MailItemId = baseRow.MailItemId, WrapperGraphId = baseRow.WrapperGraphId,
            ConversationId = baseRow.ConversationId,
            SourceType = baseRow.SourceType, SourceMailbox = baseRow.SourceMailbox,
            FromAddress = baseRow.FromAddress, FromName = baseRow.FromName, Subject = baseRow.Subject,
            ReceivedAt = baseRow.ReceivedAt, HasAttachments = baseRow.HasAttachments, AttachmentsJson = baseRow.AttachmentsJson,
            Completed = baseRow.Completed, CompletedAt = baseRow.CompletedAt, CompletedBy = baseRow.CompletedBy,
            IsRead = baseRow.IsRead, ReadAt = baseRow.ReadAt, ReadBy = baseRow.ReadBy,
            ClaimedBy = baseRow.ClaimedBy, ClaimedAt = baseRow.ClaimedAt, Direction = baseRow.Direction,
            ToLine = GetStr(f, "ToLine") ?? "", CcLine = GetStr(f, "CcLine") ?? "",
            BodyText = GetStr(f, "BodyText") ?? "", RawEmlUrl = GetStr(f, "RawEmlUrl"),
        };
    }

    private async Task<(string SiteId, string ListId, string? SpId)> ResolveMailItemSpIdAsync(string mailItemId, CancellationToken ct)
    {
        var siteId = await GetSiteIdAsync();
        var listId = await ResolveListIdAsync(MailItemsList);
        var res = await GetGraph().Sites[siteId].Lists[listId].Items.GetAsync(req =>
        {
            req.QueryParameters.Expand = ["fields($select=MailItemId)"];
            req.QueryParameters.Top    = 1;
            req.QueryParameters.Filter = $"fields/MailItemId eq '{Esc(mailItemId)}'";
        }, ct);
        return (siteId, listId, res?.Value?.FirstOrDefault()?.Id);
    }

    /// <summary>Patches a MailItem's attachment manifest (now carrying file paths) + the raw-.eml pointer.</summary>
    public async Task UpdateMailItemFilesAsync(string mailItemId, string attachmentsJson, string? rawEmlUrl, CancellationToken ct = default)
    {
        var (siteId, listId, spId) = await ResolveMailItemSpIdAsync(mailItemId, ct);
        if (spId is null) return;
        var fields = new Dictionary<string, object?> { ["AttachmentsJson"] = Trunc(attachmentsJson, 100000) };
        if (!string.IsNullOrEmpty(rawEmlUrl)) fields["RawEmlUrl"] = rawEmlUrl;
        await GetGraph().Sites[siteId].Lists[listId].Items[spId].Fields
            .PatchAsync(new FieldValueSet { AdditionalData = fields }, cancellationToken: ct);
    }

    /// <summary>Deletes a MailItem list row by its SP item id.</summary>
    public async Task DeleteMailItemAsync(string spId, CancellationToken ct = default)
    {
        var siteId = await GetSiteIdAsync();
        var listId = await ResolveListIdAsync(MailItemsList);
        await GetGraph().Sites[siteId].Lists[listId].Items[spId].DeleteAsync(cancellationToken: ct);
    }

    /// <summary>Deletes all MailClassification rows for a given MailItemId.</summary>
    public async Task DeleteClassificationsForItemAsync(string mailItemId, CancellationToken ct = default)
    {
        var siteId = await GetSiteIdAsync();
        var listId = await ResolveListIdAsync(MailClassList);
        foreach (var r in await ReadClassificationsForItemAsync(mailItemId, ct))
            if (r.SpId.Length > 0)
                await GetGraph().Sites[siteId].Lists[listId].Items[r.SpId].DeleteAsync(cancellationToken: ct);
    }

    /// <summary>Reads one captured item's classify-relevant text (for re-classification).</summary>
    public async Task<MailClassifyInput?> GetClassifyInputAsync(string mailItemId, CancellationToken ct = default)
    {
        var rows = await ReadAllListItemsAsync(MailItemsList,
            ["MailItemId","FromAddress","FromName","ToLine","EmailSubject","BodyText","AttachmentsJson"],
            $"fields/MailItemId eq '{Esc(mailItemId)}'", ct);
        var f = rows.FirstOrDefault();
        if (f is null) return null;

        var names = new List<string>();
        try
        {
            var arr = JsonSerializer.Deserialize<List<MailAttManifest>>(GetStr(f, "AttachmentsJson") ?? "[]");
            if (arr is not null) names = arr.Select(a => a.Name).Where(n => !string.IsNullOrEmpty(n)).ToList();
        }
        catch { /* manifest optional */ }

        return new MailClassifyInput
        {
            Subject         = GetStr(f, "EmailSubject") ?? "",
            FromAddress     = GetStr(f, "FromAddress") ?? "",
            FromName        = GetStr(f, "FromName") ?? "",
            ToLine          = GetStr(f, "ToLine") ?? "",
            BodyText        = GetStr(f, "BodyText") ?? "",
            AttachmentNames = names,
        };
    }

    // ── MailClassifications write/read ───────────────────────────────────────────────

    /// <summary>Writes a new classification version for an item; supersedes the prior current one.</summary>
    public async Task<int> WriteClassificationAsync(string mailItemId, MailClassificationResult r, CancellationToken ct = default)
    {
        var siteId = await GetSiteIdAsync();
        var listId = await ResolveListIdAsync(MailClassList);

        var prior = await ReadClassificationsForItemAsync(mailItemId, ct);
        var nextVersion = prior.Count == 0 ? 1 : prior.Max(p => p.Version) + 1;

        // Flip any current rows to not-current.
        foreach (var p in prior.Where(p => p.IsCurrent && p.SpId.Length > 0))
        {
            await GetGraph().Sites[siteId].Lists[listId].Items[p.SpId].Fields.PatchAsync(
                new FieldValueSet { AdditionalData = new Dictionary<string, object?> { ["IsCurrent"] = false } }, cancellationToken: ct);
        }

        await GetGraph().Sites[siteId].Lists[listId].Items.PostAsync(new ListItem
        {
            Fields = new FieldValueSet
            {
                AdditionalData = new Dictionary<string, object?>
                {
                    ["Title"]        = mailItemId,
                    ["MailItemId"]   = mailItemId,
                    ["Version"]      = nextVersion,
                    ["IsCurrent"]    = true,
                    ["CategoryPath"] = r.Category,
                    ["OtherLabel"]   = r.OtherLabel,
                    ["SupplierName"] = r.SupplierName,
                    ["Confidence"]   = r.Confidence,
                    ["KeywordTags"]  = string.Join(", ", r.Keywords),
                    ["PoNumber"]     = r.PoNumber,
                    ["SoNumber"]     = r.SoNumber,
                    ["Amount"]       = r.Amount,
                    ["SupplierReference"] = r.SupplierReference,
                    ["PayLink"]      = r.PayLink,
                    ["Reasoning"]    = Trunc(r.Reasoning, 8000),
                    ["AiProvider"]   = r.AiProvider,
                    ["AiModel"]      = r.AiModel,
                    ["RawAiResponse"]= Trunc(r.RawResponse, 100000),
                    ["ClassifiedAt"] = DateTimeOffset.UtcNow.ToString("o"),
                }
            }
        }, cancellationToken: ct);

        _log.LogInformation("[MailWB] Classified {Id} v{Ver} -> {Cat}", mailItemId, nextVersion, r.Category);
        return nextVersion;
    }

    public async Task<List<MailClassRow>> ReadClassificationsForItemAsync(string mailItemId, CancellationToken ct = default)
    {
        var rows = await ReadAllListItemsAsync(MailClassList,
            ["MailItemId","Version","IsCurrent","CategoryPath","OtherLabel","SupplierName","Confidence","KeywordTags",
             "PoNumber","SoNumber","Amount","SupplierReference","PayLink","Reasoning","AiProvider","AiModel","ClassifiedAt"],
            $"fields/MailItemId eq '{Esc(mailItemId)}'", ct);
        return rows.Select(MapClassRow).OrderByDescending(c => c.Version).ToList();
    }

    /// <summary>All current classifications (one per item). Used for the tree + item lists.</summary>
    public async Task<List<MailClassRow>> ReadCurrentClassificationsAsync(CancellationToken ct = default)
    {
        // Read ALL classification rows (no server-side "IsCurrent eq true" filter — that boolean
        // filter lags badly on SP's eventual-consistency index right after a write, so the tree
        // would show stale/zero counts immediately after capture/reclassify). Compute the current
        // row per item as the highest Version in memory — always correct and lag-free.
        var rows = await ReadAllListItemsAsync(MailClassList,
            ["MailItemId","Version","IsCurrent","CategoryPath","OtherLabel","SupplierName","Confidence","KeywordTags",
             "PoNumber","SoNumber","Amount","SupplierReference","PayLink","AiProvider","AiModel","ClassifiedAt"],
            null, ct);
        return rows.Select(MapClassRow)
            .GroupBy(c => c.MailItemId, StringComparer.Ordinal)
            .Select(g => g.OrderByDescending(c => c.Version).First())
            .ToList();
    }

    private static MailClassRow MapClassRow(Dictionary<string, object?> f) => new()
    {
        SpId         = GetStr(f, "__spId") ?? "",
        MailItemId   = GetStr(f, "MailItemId") ?? "",
        Version      = (int)GetDouble(f, "Version"),
        IsCurrent    = GetBool(f, "IsCurrent"),
        CategoryPath = GetStr(f, "CategoryPath") ?? "Other",
        OtherLabel   = GetStr(f, "OtherLabel"),
        SupplierName = GetStr(f, "SupplierName"),
        Confidence   = GetDouble(f, "Confidence"),
        KeywordTags  = GetStr(f, "KeywordTags") ?? "",
        PoNumber     = GetStr(f, "PoNumber"),
        SoNumber     = GetStr(f, "SoNumber"),
        Amount       = GetStr(f, "Amount"),
        SupplierReference = GetStr(f, "SupplierReference"),
        PayLink      = GetStr(f, "PayLink"),
        AiProvider   = GetStr(f, "AiProvider") ?? "",
        AiModel      = GetStr(f, "AiModel") ?? "",
        ClassifiedAt = GetStr(f, "ClassifiedAt") ?? "",
    };

    // ── Shared helpers ────────────────────────────────────────────────────────────

    private async Task<List<Dictionary<string, object?>>> ReadAllListItemsAsync(
        string listName, string[] selectFields, string? filter, CancellationToken ct)
    {
        var siteId = await GetSiteIdAsync();
        var listId = await ResolveListIdAsync(listName);
        var rows   = new List<Dictionary<string, object?>>();

        var page = await GetGraph().Sites[siteId].Lists[listId].Items.GetAsync(req =>
        {
            req.QueryParameters.Expand = [$"fields($select={string.Join(",", selectFields)})"];
            req.QueryParameters.Top    = 200;
            if (filter is not null) req.QueryParameters.Filter = filter;
        }, ct);

        while (page?.Value is not null)
        {
            foreach (var it in page.Value)
            {
                var f = new Dictionary<string, object?>();
                var ad = it.Fields?.AdditionalData;
                if (ad is not null) foreach (var kv in ad) f[kv.Key] = kv.Value;
                f["__spId"] = it.Id;
                rows.Add(f);
            }
            if (string.IsNullOrEmpty(page.OdataNextLink)) break;
            page = await GetGraph().Sites[siteId].Lists[listId].Items.WithUrl(page.OdataNextLink).GetAsync(cancellationToken: ct);
        }
        return rows;
    }

    private static string Esc(string s) => (s ?? "").Replace("'", "''");
    private static string? Trunc(string? s, int max) =>
        string.IsNullOrEmpty(s) ? s : (s.Length <= max ? s : s[..max]);

    private static bool GetBool(IDictionary<string, object?> d, string key)
    {
        if (!d.TryGetValue(key, out var v) || v is null) return false;
        if (v is bool b) return b;
        if (v is JsonElement je) return je.ValueKind == JsonValueKind.True || (je.ValueKind == JsonValueKind.String && bool.TryParse(je.GetString(), out var pb) && pb);
        return bool.TryParse(v.ToString(), out var p) && p;
    }

    private static double GetDouble(IDictionary<string, object?> d, string key)
    {
        if (!d.TryGetValue(key, out var v) || v is null) return 0;
        if (v is double dv) return dv;
        if (v is JsonElement je)
        {
            if (je.ValueKind == JsonValueKind.Number) return je.GetDouble();
            if (je.ValueKind == JsonValueKind.String && double.TryParse(je.GetString(), out var sd)) return sd;
            return 0;
        }
        return double.TryParse(v.ToString(), out var p) ? p : 0;
    }
}

// ── DTOs ────────────────────────────────────────────────────────────────────────

/// <summary>A cross-conversation project: a named cluster of conversations linked by shared business refs.</summary>
public sealed class MailProjectRow
{
    public string SpId            { get; set; } = "";
    public string ProjectId       { get; set; } = "";
    public string Name            { get; set; } = "";
    public string Status          { get; set; } = "active";   // active | archived
    public List<string> Refs            { get; set; } = [];
    public List<string> ConversationIds { get; set; } = [];
    public string CreatedAt       { get; set; } = "";
    public string CreatedBy       { get; set; } = "";
}

/// <summary>A live taxonomy hint row (custom leaf + optional classification guidance).</summary>
public sealed class TaxonomyHintRow
{
    public string CategoryPath { get; set; } = "";
    public string Hint         { get; set; } = "";
    public string Source       { get; set; } = "";
    public string CreatedAt    { get; set; } = "";
}

/// <summary>Attachment manifest entry stored as AttachmentsJson on a MailItem.</summary>
public sealed class MailAttManifest
{
    public string Name        { get; set; } = "";
    public string ContentType { get; set; } = "";
    public long   Size        { get; set; }
    public string? WebUrl     { get; set; }   // set in Phase 1b doc-library storage
}

public sealed class MailItemInput
{
    public string SourceType        { get; set; } = "email";
    public string SourceMailbox     { get; set; } = "";
    public string InternetMessageId { get; set; } = "";
    public string WrapperGraphId    { get; set; } = "";
    public string ConversationId    { get; set; } = "";
    public string RefsJson          { get; set; } = "[]";
    public string FromAddress       { get; set; } = "";
    public string FromName          { get; set; } = "";
    public string ToLine            { get; set; } = "";
    public string CcLine            { get; set; } = "";
    public string Subject           { get; set; } = "";
    public string ReceivedAtIso     { get; set; } = "";
    public string BodyText          { get; set; } = "";
    public bool   HasAttachments    { get; set; }
    public string AttachmentsJson   { get; set; } = "[]";
    public string Direction         { get; set; } = "in";   // "in" | "out"
}

public class MailItemRow
{
    public string  SpId            { get; set; } = "";
    public string  MailItemId      { get; set; } = "";
    public string  WrapperGraphId  { get; set; } = "";
    public string  ConversationId  { get; set; } = "";
    public string  RefsJson        { get; set; } = "";
    public string  SourceType      { get; set; } = "email";
    public string  SourceMailbox   { get; set; } = "";
    public string  FromAddress     { get; set; } = "";
    public string  FromName        { get; set; } = "";
    public string  Subject         { get; set; } = "";
    public string  ReceivedAt      { get; set; } = "";
    public bool    HasAttachments  { get; set; }
    public string  AttachmentsJson { get; set; } = "";
    public bool    Completed       { get; set; }
    public string? CompletedAt     { get; set; }
    public string? CompletedBy     { get; set; }
    public bool    IsRead          { get; set; }
    public string? ReadAt          { get; set; }
    public string? ReadBy          { get; set; }
    public string? ClaimedBy       { get; set; }
    public string? ClaimedAt       { get; set; }
    public string  Direction       { get; set; } = "in";   // "in" | "out" (workbench-sent)
}

/// <summary>A MailItem row plus the heavy/extra fields the full viewer needs.</summary>
public sealed class MailItemDetailRow : MailItemRow
{
    public string  ToLine    { get; set; } = "";
    public string  CcLine    { get; set; } = "";
    public string  BodyText  { get; set; } = "";
    public string? RawEmlUrl { get; set; }
}

public sealed class MailClassRow
{
    public string  SpId         { get; set; } = "";
    public string  MailItemId   { get; set; } = "";
    public int     Version      { get; set; }
    public bool    IsCurrent    { get; set; }
    public string  CategoryPath { get; set; } = "Other";
    public string? OtherLabel   { get; set; }
    public string? SupplierName { get; set; }
    public double  Confidence   { get; set; }
    public string  KeywordTags  { get; set; } = "";
    public string? PoNumber     { get; set; }
    public string? SoNumber     { get; set; }
    public string? Amount       { get; set; }
    public string? SupplierReference { get; set; }
    public string? PayLink      { get; set; }
    public string  AiProvider   { get; set; } = "";
    public string  AiModel      { get; set; } = "";
    public string  ClassifiedAt { get; set; } = "";
}
