using System.Text.Json;
using System.Text.RegularExpressions;
using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.Hosting;
using Microsoft.Extensions.Logging;
using OutlookShredder.Proxy.Models;

namespace OutlookShredder.Proxy.Services;

/// <summary>
/// Accelerator (Fulfillment loop - wip/fulfillment-loop.md): matches classified supplier
/// confirmation / receipt emails to open PurchaseOrders by PO number (HSK-PO...) and records the
/// confirmation (via = email | payment) so the PO monitor flips to Confirmed without a manual tap.
///
/// Manual confirm stays the baseline; this only ACCELERATES on a high-confidence signal - an exact
/// OUR-PO# match inside a confirmation-category email - and every auto-confirm is reversible
/// (unconfirm) with a provenance note. Historical/stale POs are skipped (the backlog stays stale).
/// </summary>
public sealed class PoConfirmationMatcherService : IHostedService, IDisposable
{
    private static readonly string[] ConfirmCategories =
        { "Supplier/Order Confirmations", "Supplier/Receipts" };

    private static readonly Regex HskPoRx =
        new(@"\bHSK-PO\d+\b", RegexOptions.IgnoreCase | RegexOptions.Compiled);

    private readonly SharePointService _sp;
    private readonly MailCacheService  _mail;
    private readonly IConfiguration    _config;
    private readonly ILogger<PoConfirmationMatcherService> _log;
    private Timer? _timer;

    public PoConfirmationMatcherService(SharePointService sp, MailCacheService mail,
        IConfiguration config, ILogger<PoConfirmationMatcherService> log)
    {
        _sp = sp; _mail = mail; _config = config; _log = log;
    }

    private bool   Enabled         => _config.GetValue("PoMatcher:Enabled", true);
    private bool   Auto            => _config.GetValue("PoMatcher:Auto", true);
    private int    IntervalMinutes => Math.Max(1, _config.GetValue("PoMatcher:IntervalMinutes", 10));
    private double ConfidenceFloor => _config.GetValue("PoMatcher:ConfidenceFloor", 0.0);
    private int    WindowDays      => Math.Max(1, _config.GetValue("PoMonitor:MonitorWindowBusinessDays", 3));

    public Task StartAsync(CancellationToken ct)
    {
        if (!Enabled) { _log.LogInformation("[PoMatcher] disabled"); return Task.CompletedTask; }
        _timer = new Timer(async _ =>
        {
            try { await RunOnceAsync(CancellationToken.None); }
            catch (Exception ex) { _log.LogWarning(ex, "[PoMatcher] cycle failed"); }
        }, null, TimeSpan.FromSeconds(45), TimeSpan.FromMinutes(IntervalMinutes));   // delay so caches warm
        return Task.CompletedTask;
    }

    public Task StopAsync(CancellationToken ct) { _timer?.Dispose(); return Task.CompletedTask; }
    public void Dispose() => _timer?.Dispose();

    public record MatchResult(int Scanned, int Matched, int Confirmed, bool Auto, List<string> Details);

    /// <summary>One matching pass (also used by the on-demand endpoint). Confirms Pending,
    /// non-stale POs whose HSK-PO number appears in a confirmation-category email.</summary>
    public async Task<MatchResult> RunOnceAsync(CancellationToken ct)
    {
        var now    = DateTimeOffset.UtcNow;
        var window = WindowDays;

        // Pending, non-stale POs keyed by upper PO number (a number can have >1 row).
        var pending = new Dictionary<string, List<PurchaseOrderRecord>>(StringComparer.OrdinalIgnoreCase);
        foreach (var po in await _sp.ReadPurchaseOrdersAsync())
        {
            if (!string.Equals(po.ConfirmStatus, "Pending", StringComparison.OrdinalIgnoreCase)) continue;
            if (string.IsNullOrWhiteSpace(po.PoNumber)) continue;
            if (DateTimeOffset.TryParse(po.ReceivedAt, out var booked)
                && PoConfirmationMonitor.IsStale(booked, now, window)) continue;   // leave backlog stale
            var key = po.PoNumber!.Trim().ToUpperInvariant();
            if (!pending.TryGetValue(key, out var list)) pending[key] = list = new();
            list.Add(po);
        }

        int scanned = 0, matched = 0, confirmed = 0;
        var details = new List<string>();

        foreach (var c in _mail.GetCurrents())
        {
            if (!ConfirmCategories.Contains(c.CategoryPath, StringComparer.OrdinalIgnoreCase)) continue;
            if (c.Confidence < ConfidenceFloor) continue;
            if (!_mail.TryGetItem(c.MailItemId, out var item)) continue;
            if (string.Equals(item.Direction, "out", StringComparison.OrdinalIgnoreCase)) continue;   // our own sent mail
            scanned++;

            var via  = c.CategoryPath.Contains("Receipt", StringComparison.OrdinalIgnoreCase) ? "payment" : "email";
            var subj = item.Subject.Length > 80 ? item.Subject[..80] : item.Subject;

            foreach (var poNum in CandidatePoNumbers(c, item))
            {
                if (!pending.TryGetValue(poNum, out var list) || list.Count == 0) continue;
                matched++;
                var note = $"Auto-matched from {c.CategoryPath} email: \"{subj}\" [{c.MailItemId}]";
                foreach (var po in list)
                {
                    if (Auto)
                    {
                        await _sp.UpdatePurchaseOrderConfirmAsync(po.SpItemId, confirmed: true, via: via,
                                                                  expectedDate: null, note: note);
                        confirmed++;
                        details.Add($"{po.PoNumber} <- {via} (\"{subj}\")");
                    }
                    else details.Add($"SUGGEST {po.PoNumber} <- {via} (\"{subj}\")");
                }
                list.Clear();   // don't re-confirm this PO again this pass
            }
        }

        _log.LogInformation("[PoMatcher] pass: scanned={Scanned} matched={Matched} confirmed={Confirmed} (auto={Auto})",
            scanned, matched, confirmed, Auto);
        return new MatchResult(scanned, matched, confirmed, Auto, details);
    }

    /// <summary>Collect candidate HSK-PO numbers from the classification's PoNumber field and the
    /// item's RefsJson (deterministic regex over subject+body at capture).</summary>
    private static IEnumerable<string> CandidatePoNumbers(MailClassRow c, MailItemRow item)
    {
        var set = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

        var pn = c.PoNumber?.Trim();
        if (!string.IsNullOrEmpty(pn))
        {
            foreach (Match m in HskPoRx.Matches(pn)) set.Add(m.Value.ToUpperInvariant());
            if (pn.StartsWith("HSK-PO", StringComparison.OrdinalIgnoreCase)) set.Add(pn.ToUpperInvariant());
        }

        if (!string.IsNullOrWhiteSpace(item.RefsJson))
        {
            try
            {
                foreach (var r in JsonSerializer.Deserialize<List<string>>(item.RefsJson) ?? new())
                    if (!string.IsNullOrWhiteSpace(r) && HskPoRx.IsMatch(r))
                        set.Add(r.Trim().ToUpperInvariant());
            }
            catch { /* ignore malformed refs */ }
        }
        return set;
    }
}
