using System.Text;
using OutlookShredder.Proxy.Models;

namespace OutlookShredder.Proxy.Services;

/// <summary>
/// Effective taxonomy = the static <see cref="MailTaxonomy"/> base + live "hints" stored in SP
/// (custom leaves + operator-confirmed classification guidance). Cached with a short TTL so a
/// confirmed leaf/hint takes effect on the NEXT classification with no code deploy. Used by the
/// classifier (prompt + tool enum + coerce), the workbench (tree), and the controller.
/// </summary>
public sealed class MailTaxonomyService
{
    private readonly SharePointService _sp;
    private readonly ILogger<MailTaxonomyService> _log;
    private readonly SemaphoreSlim _gate = new(1, 1);
    private static readonly TimeSpan Ttl = TimeSpan.FromMinutes(3);

    private volatile List<MailTaxonomy.Leaf> _leaves = MailTaxonomy.Leaves.ToList();
    private volatile List<TaxonomyHintRow> _hints = new();
    private volatile HashSet<string> _validPaths = MailTaxonomy.ValidPaths.ToHashSet(StringComparer.OrdinalIgnoreCase);
    private DateTimeOffset _loadedAt = DateTimeOffset.MinValue;

    public MailTaxonomyService(SharePointService sp, ILogger<MailTaxonomyService> log) { _sp = sp; _log = log; }

    private async Task EnsureFreshAsync(CancellationToken ct)
    {
        if (DateTimeOffset.UtcNow - _loadedAt < Ttl) return;
        await _gate.WaitAsync(ct);
        try
        {
            if (DateTimeOffset.UtcNow - _loadedAt < Ttl) return;
            var hints = await _sp.ReadTaxonomyHintsAsync(ct);

            // Custom leaves = hint CategoryPaths not already in the static taxonomy.
            var custom = hints.Select(h => h.CategoryPath)
                .Where(p => !MailTaxonomy.ValidPaths.Contains(p))
                .Distinct(StringComparer.OrdinalIgnoreCase)
                .Select(p =>
                {
                    var slash = p.IndexOf('/');
                    var top   = slash < 0 ? p : p[..slash];
                    var sub   = slash < 0 ? "" : p[(slash + 1)..];
                    var desc  = hints.FirstOrDefault(h => string.Equals(h.CategoryPath, p, StringComparison.OrdinalIgnoreCase) && h.Hint.Length > 0)?.Hint
                                ?? "Operator-confirmed custom category.";
                    return new MailTaxonomy.Leaf(top, sub, desc);
                }).ToList();

            _hints      = hints;
            _leaves     = MailTaxonomy.Leaves.Concat(custom).ToList();
            _validPaths = _leaves.Select(l => l.Path).ToHashSet(StringComparer.OrdinalIgnoreCase);
            _loadedAt   = DateTimeOffset.UtcNow;
            if (custom.Count > 0 || hints.Count > 0)
                _log.LogInformation("[MailTaxonomy] effective taxonomy: {Custom} custom leaf(s), {Hints} hint(s)", custom.Count, hints.Count);
        }
        catch (Exception ex) { _log.LogWarning(ex, "[MailTaxonomy] hint load failed — using last-good taxonomy"); _loadedAt = DateTimeOffset.UtcNow; }
        finally { _gate.Release(); }
    }

    public async Task<List<MailTaxonomy.Leaf>> GetLeavesAsync(CancellationToken ct = default) { await EnsureFreshAsync(ct); return _leaves; }

    public async Task<List<TaxonomyHintRow>> GetHintsAsync(CancellationToken ct = default) { await EnsureFreshAsync(ct); return _hints; }

    /// <summary>Snaps a category to a known effective leaf (sync; uses the last-loaded set).</summary>
    public string Coerce(string? category)
    {
        if (string.IsNullOrWhiteSpace(category)) return "Other";
        var c = category.Trim();
        return _validPaths.Contains(c) ? c : "Other";
    }

    /// <summary>Renders the taxonomy block (base + custom leaves) + a learned-hints section for the prompt.</summary>
    public async Task<string> RenderForPromptAsync(CancellationToken ct = default)
    {
        await EnsureFreshAsync(ct);
        var leaves = _leaves;
        var sb = new StringBuilder();
        foreach (var top in leaves.Select(l => l.Top).Distinct())
        {
            sb.Append("- ").Append(top).Append('\n');
            foreach (var leaf in leaves.Where(l => l.Top == top))
            {
                if (string.IsNullOrEmpty(leaf.Sub)) sb.Append("    (").Append(top).Append("): ").Append(leaf.Description).Append('\n');
                else                                sb.Append("    \"").Append(leaf.Path).Append("\": ").Append(leaf.Description).Append('\n');
            }
        }
        var guidance = _hints.Where(h => !string.IsNullOrWhiteSpace(h.Hint)).ToList();
        if (guidance.Count > 0)
        {
            sb.Append("\nLearned hints (operator-confirmed — apply these):\n");
            foreach (var h in guidance) sb.Append("    ").Append(h.CategoryPath).Append(": ").Append(h.Hint).Append('\n');
        }
        return sb.ToString();
    }

    /// <summary>Records a confirmed leaf + optional guidance to SP and forces a refresh.</summary>
    public async Task AddLeafHintAsync(string categoryPath, string? hint, string source, CancellationToken ct = default)
    {
        await _sp.WriteTaxonomyHintAsync(categoryPath, hint, source, ct);
        _loadedAt = DateTimeOffset.MinValue;   // next use reloads
    }
}
