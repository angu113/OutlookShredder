using System.Text.Json;
using OutlookShredder.Proxy.Models;

namespace OutlookShredder.Proxy.Services;

/// <summary>One suggested PO&lt;-&gt;SO link based on a shared MSPC, for a card that has no direct
/// HSK-SO link. Surfaced on both the PO card (as "Link to: {SoNumber}?") and the slip card
/// ("Link to: {PoNumber}?"); the user confirms (full link) or rejects (remembered per pair).</summary>
public record LinkSuggestion(string PoNumber, string PoSpItemId, string SoNumber, int SharedMspcs);

/// <summary>
/// Fallback linkage for the Transfer board: when a sales-order (slip) card has no PO whose captured
/// SalesOrders include its HSK-SO, look for active POs that share an MSPC with the SO (and vice-versa)
/// and suggest the link. "Any shared MSPC" qualifies. Rejected pairs (PurchaseOrders.RejectedLinks) and
/// already-linked pairs (SalesOrders) are excluded. Suggestions are cached briefly so the board can
/// fetch them cheaply.
/// </summary>
public class TransferLinkSuggestionService
{
    private readonly SharePointService   _sp;
    private readonly WorkflowCardService _wf;
    private readonly IConfiguration      _config;
    private readonly ILogger<TransferLinkSuggestionService> _log;

    private readonly SemaphoreSlim _gate = new(1, 1);
    private IReadOnlyList<LinkSuggestion> _cache = [];
    private DateTimeOffset _cacheAt = DateTimeOffset.MinValue;
    private static readonly TimeSpan Ttl = TimeSpan.FromSeconds(60);

    public TransferLinkSuggestionService(SharePointService sp, WorkflowCardService wf,
        IConfiguration config, ILogger<TransferLinkSuggestionService> log)
    {
        _sp = sp; _wf = wf; _config = config; _log = log;
    }

    public void Invalidate() => _cacheAt = DateTimeOffset.MinValue;

    public async Task<IReadOnlyList<LinkSuggestion>> GetSuggestionsAsync(bool force = false)
    {
        if (!force && DateTimeOffset.UtcNow - _cacheAt < Ttl) return _cache;
        await _gate.WaitAsync();
        try
        {
            if (!force && DateTimeOffset.UtcNow - _cacheAt < Ttl) return _cache;
            _cache   = await ComputeAsync();
            _cacheAt = DateTimeOffset.UtcNow;
            return _cache;
        }
        catch (Exception ex) { _log.LogWarning(ex, "[LINK-SUGGEST] compute failed"); return _cache; }
        finally { _gate.Release(); }
    }

    private async Task<List<LinkSuggestion>> ComputeAsync()
    {
        var pos = await _sp.ReadPurchaseOrdersAsync();
        int lookback = _config.GetValue("WaitingBoard:LookbackDays", 45);
        var cutoff = DateTimeOffset.UtcNow.AddDays(-lookback);
        bool Active(PurchaseOrderRecord p) =>
            string.IsNullOrWhiteSpace(p.MaterialReceivedAt) &&
            ((DateTimeOffset.TryParse(p.ReceivedAt, out var b) && b >= cutoff)
             || string.Equals(p.ConfirmStatus, "Confirmed", StringComparison.OrdinalIgnoreCase)
             || !string.IsNullOrWhiteSpace(p.ExpectedDate) || !string.IsNullOrWhiteSpace(p.BoardDate));

        var poViews = pos
            .Where(p => !string.IsNullOrWhiteSpace(p.PoNumber) && Active(p))
            .Select(p => new PoView(p.PoNumber!, p.SpItemId, MspcsFromLineItems(p.LineItems),
                                    SplitSet(p.SalesOrders), SplitSet(p.RejectedLinks)))
            .Where(v => v.Mspcs.Count > 0)
            .ToList();
        if (poViews.Count == 0) return [];

        // SO universe = the HSK-SO of the slip cards currently on the board.
        var slips = await _wf.GetAllAsync();
        var activeSOs = slips
            .Select(c => (c.DocumentNumber ?? "").Trim())
            .Where(d => d.StartsWith("HSK-SO", StringComparison.OrdinalIgnoreCase))
            .ToHashSet(StringComparer.OrdinalIgnoreCase);
        if (activeSOs.Count == 0) return [];

        // SO -> MSPCs, read from each SO's picking-slip ERP document line items.
        var erp = await _sp.ReadErpDocumentsAsync(top: 20000, includeArchived: true);
        var soMspcs = new Dictionary<string, HashSet<string>>(StringComparer.OrdinalIgnoreCase);
        foreach (var d in erp)
        {
            if (!string.Equals(d.DocumentType, "PickingSlip", StringComparison.OrdinalIgnoreCase)) continue;
            var num = (d.DocumentNumber ?? "").Trim();
            if (!activeSOs.Contains(num)) continue;
            var m = MspcsFromLineItems(d.LineItemsJson);
            if (m.Count == 0) continue;
            if (soMspcs.TryGetValue(num, out var ex)) ex.UnionWith(m); else soMspcs[num] = m;
        }

        var suggestions = Compute(poViews, soMspcs);
        _log.LogInformation("[LINK-SUGGEST] {N} suggestion(s) from {Pos} active PO(s) x {Sos} slip SO(s)",
            suggestions.Count, poViews.Count, soMspcs.Count);
        return suggestions;
    }

    // ── pure ───────────────────────────────────────────────────────────────────

    public record PoView(string PoNumber, string PoSpItemId, HashSet<string> Mspcs,
                         HashSet<string> LinkedSOs, HashSet<string> RejectedSOs);

    /// <summary>Distinct MSPC codes (those containing '/') from a line-items JSON array — handles both
    /// the PO row's camelCase <c>mspc</c> and the ERP doc's <c>Code</c> field.</summary>
    public static HashSet<string> MspcsFromLineItems(string? json)
    {
        var set = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        if (string.IsNullOrWhiteSpace(json)) return set;
        try
        {
            using var doc = JsonDocument.Parse(json);
            if (doc.RootElement.ValueKind != JsonValueKind.Array) return set;
            foreach (var el in doc.RootElement.EnumerateArray())
            {
                if (el.ValueKind != JsonValueKind.Object) continue;
                foreach (var name in (ReadOnlySpan<string>)["mspc", "Mspc", "code", "Code"])
                {
                    if (el.TryGetProperty(name, out var v) && v.ValueKind == JsonValueKind.String)
                    {
                        var s = v.GetString();
                        if (!string.IsNullOrWhiteSpace(s) && s.Contains('/'))
                            set.Add(s.Trim().ToUpperInvariant());
                    }
                }
            }
        }
        catch { /* malformed JSON → no MSPCs */ }
        return set;
    }

    /// <summary>CSV (comma/semicolon) → case-insensitive set.</summary>
    public static HashSet<string> SplitSet(string? csv) =>
        (csv ?? "").Split([',', ';'], StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries)
                   .ToHashSet(StringComparer.OrdinalIgnoreCase);

    /// <summary>Every (PO, SO) pair that shares ≥1 MSPC and isn't already linked or rejected.</summary>
    public static List<LinkSuggestion> Compute(
        IEnumerable<PoView> pos, IReadOnlyDictionary<string, HashSet<string>> soMspcs)
    {
        var result = new List<LinkSuggestion>();
        foreach (var po in pos)
        {
            if (po.Mspcs.Count == 0) continue;
            foreach (var (so, mspcs) in soMspcs)
            {
                if (po.LinkedSOs.Contains(so) || po.RejectedSOs.Contains(so)) continue;
                int shared = po.Mspcs.Count(mspcs.Contains);
                if (shared > 0) result.Add(new LinkSuggestion(po.PoNumber, po.PoSpItemId, so, shared));
            }
        }
        return result;
    }
}
