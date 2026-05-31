using System.Text.Json;

namespace OutlookShredder.Proxy.Services;

/// <summary>
/// Layer 2: cross-conversation "projects". A project is a named cluster of conversations that belong
/// to one piece of business (e.g. the press-brake purchase: rigging + container delivery + vendor
/// threads, linked by a shared booking/container ref). Projects are persisted in the MailProjects SP
/// list; the suggestion engine proposes them by finding a business reference shared across two or more
/// conversations that aren't already grouped — the operator confirms/names them (hybrid model).
/// </summary>
public sealed class MailProjectService
{
    private readonly SharePointService _sp;
    private readonly MailCacheService _cache;
    private readonly ILogger<MailProjectService> _log;
    private readonly SemaphoreSlim _gate = new(1, 1);
    private static readonly TimeSpan Ttl = TimeSpan.FromMinutes(2);

    private volatile List<MailProjectRow> _projects = new();
    private DateTimeOffset _loadedAt = DateTimeOffset.MinValue;

    public MailProjectService(SharePointService sp, MailCacheService cache, ILogger<MailProjectService> log)
    { _sp = sp; _cache = cache; _log = log; }

    private async Task EnsureFreshAsync(CancellationToken ct)
    {
        if (DateTimeOffset.UtcNow - _loadedAt < Ttl) return;
        await _gate.WaitAsync(ct);
        try
        {
            if (DateTimeOffset.UtcNow - _loadedAt < Ttl) return;
            _projects = await _sp.ReadProjectsAsync(ct);
            _loadedAt = DateTimeOffset.UtcNow;
        }
        catch (Exception ex) { _log.LogWarning(ex, "[MailProjects] load failed — using last-good"); _loadedAt = DateTimeOffset.UtcNow; }
        finally { _gate.Release(); }
    }

    public async Task<List<MailProjectRow>> GetProjectsAsync(bool activeOnly = true, CancellationToken ct = default)
    {
        await EnsureFreshAsync(ct);
        return activeOnly
            ? _projects.Where(p => !string.Equals(p.Status, "archived", StringComparison.OrdinalIgnoreCase)).ToList()
            : _projects.ToList();
    }

    public async Task<MailProjectRow?> GetProjectAsync(string projectId, CancellationToken ct = default)
    {
        await EnsureFreshAsync(ct);
        return _projects.FirstOrDefault(p => string.Equals(p.ProjectId, projectId, StringComparison.OrdinalIgnoreCase));
    }

    public async Task<MailProjectRow> CreateAsync(string name, List<string> conversationIds, List<string> refs, string? by, CancellationToken ct = default)
    {
        var row = new MailProjectRow
        {
            Name = string.IsNullOrWhiteSpace(name) ? "Untitled project" : name.Trim(),
            Status = "active",
            ConversationIds = conversationIds.Distinct(StringComparer.OrdinalIgnoreCase).ToList(),
            Refs = refs.Distinct(StringComparer.OrdinalIgnoreCase).ToList(),
            CreatedBy = by ?? "",
        };
        row.ProjectId = await _sp.WriteProjectAsync(row, ct);
        _loadedAt = DateTimeOffset.MinValue;   // force reload
        _log.LogInformation("[MailProjects] created '{Name}' ({Convs} convs)", row.Name, row.ConversationIds.Count);
        return row;
    }

    public async Task<bool> ArchiveAsync(string projectId, CancellationToken ct = default)
    {
        var ok = await _sp.UpdateProjectAsync(projectId, name: null, status: "archived", conversationIds: null, ct);
        _loadedAt = DateTimeOffset.MinValue;
        return ok;
    }

    public async Task<bool> UpdateAsync(string projectId, string? name, List<string>? conversationIds, CancellationToken ct = default)
    {
        var ok = await _sp.UpdateProjectAsync(projectId, name, status: null, conversationIds, ct);
        _loadedAt = DateTimeOffset.MinValue;
        return ok;
    }

    /// <summary>The set of conversation keys belonging to a project (for the item filter).</summary>
    public async Task<HashSet<string>> ConversationIdsForAsync(string projectId, CancellationToken ct = default)
    {
        var p = await GetProjectAsync(projectId, ct);
        return p is null ? new(StringComparer.OrdinalIgnoreCase) : p.ConversationIds.ToHashSet(StringComparer.OrdinalIgnoreCase);
    }

    /// <summary>All conversation keys claimed by ANY active project — these items are shown only under
    /// Projects, not in the main taxonomy tree, until removed from the project.</summary>
    public async Task<HashSet<string>> AllProjectConversationIdsAsync(CancellationToken ct = default)
    {
        var ps = await GetProjectsAsync(activeOnly: true, ct);
        return ps.SelectMany(p => p.ConversationIds).ToHashSet(StringComparer.OrdinalIgnoreCase);
    }

    /// <summary>
    /// Proposes projects: a business ref shared by ≥2 conversations that aren't already grouped. Refs
    /// linking the same conversation set are merged into one suggestion. Refs spanning too many
    /// conversations (likely generic) are skipped.
    /// </summary>
    public async Task<List<ProjectSuggestion>> SuggestAsync(CancellationToken ct = default)
    {
        var items    = _cache.GetItems();
        var projects = await GetProjectsAsync(activeOnly: true, ct);
        var inProject = projects.SelectMany(p => p.ConversationIds).ToHashSet(StringComparer.OrdinalIgnoreCase);

        var refToConvs   = new Dictionary<string, HashSet<string>>(StringComparer.OrdinalIgnoreCase);
        var convSubject  = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        var convReceived = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);

        foreach (var it in items)
        {
            var conv = string.IsNullOrEmpty(it.ConversationId) ? "id:" + it.MailItemId : it.ConversationId;
            if (string.CompareOrdinal(it.ReceivedAt, convReceived.GetValueOrDefault(conv, "")) > 0)
            { convReceived[conv] = it.ReceivedAt; convSubject[conv] = it.Subject; }
            foreach (var r in ParseRefs(it.RefsJson))
            {
                if (!refToConvs.TryGetValue(r, out var set)) refToConvs[r] = set = new(StringComparer.OrdinalIgnoreCase);
                set.Add(conv);
            }
        }

        // Merge refs that produce the same fresh conversation-set into one suggestion.
        var byConvSet = new Dictionary<string, (HashSet<string> Convs, List<string> Refs)>(StringComparer.Ordinal);
        foreach (var (r, convs) in refToConvs)
        {
            if (convs.Count < 2 || convs.Count > 12) continue;
            var fresh = convs.Where(c => !inProject.Contains(c)).ToHashSet(StringComparer.OrdinalIgnoreCase);
            if (fresh.Count < 2) continue;
            var key = string.Join("|", fresh.OrderBy(x => x, StringComparer.OrdinalIgnoreCase));
            if (!byConvSet.TryGetValue(key, out var agg)) byConvSet[key] = agg = (fresh, new List<string>());
            agg.Refs.Add(r);
        }

        return byConvSet.Values.Select(v => new ProjectSuggestion
        {
            Refs            = v.Refs.OrderBy(x => x, StringComparer.OrdinalIgnoreCase).ToList(),
            ConversationIds = v.Convs.ToList(),
            Count           = v.Convs.Count,
            Samples         = v.Convs.Select(c => convSubject.GetValueOrDefault(c, "")).Where(s => s.Length > 0).Distinct().Take(6).ToList(),
        }).OrderByDescending(s => s.Count).ToList();
    }

    private static List<string> ParseRefs(string? json)
    {
        if (string.IsNullOrWhiteSpace(json)) return [];
        try { return JsonSerializer.Deserialize<List<string>>(json) ?? []; } catch { return []; }
    }
}

/// <summary>A proposed project: conversations linked by one or more shared business references.</summary>
public sealed class ProjectSuggestion
{
    public List<string> Refs            { get; set; } = [];
    public List<string> ConversationIds { get; set; } = [];
    public int          Count           { get; set; }
    public List<string> Samples         { get; set; } = [];
}
