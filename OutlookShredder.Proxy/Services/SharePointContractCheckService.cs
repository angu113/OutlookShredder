using Microsoft.Extensions.Hosting;
using OutlookShredder.Proxy.Models;

namespace OutlookShredder.Proxy.Services;

/// <summary>
/// SharePoint data-contract self-check (the "observe the real boundary" layer). Unit tests use plain C# objects,
/// so they can't catch a Graph/Kiota deserialization mismatch — a SharePoint Boolean comes back as a native bool
/// (not a JsonElement), and reading it wrong silently yields false. This service EXERCISES the real round-trip at
/// startup (and on demand via <c>/api/diag/sp-contract</c>): it writes a throwaway Messages row with a known
/// typed value, reads it back THROUGH THE PRODUCTION READER, asserts the value survived, then deletes the row.
/// A mismatch is logged at Error — loud, at startup, not in production. Extend with Number/DateTime probes as new
/// typed fields are added.
/// </summary>
public sealed class SharePointContractCheckService : IHostedService
{
    private const string Sentinel = "__SP_CONTRACT_CHECK__";

    private readonly SharePointService _sp;
    private readonly ILogger<SharePointContractCheckService> _log;
    private Timer? _timer;

    public ContractCheckResult? Last { get; private set; }

    public SharePointContractCheckService(SharePointService sp, ILogger<SharePointContractCheckService> log)
    {
        _sp = sp; _log = log;
    }

    public Task StartAsync(CancellationToken ct)
    {
        // Run once ~30s after start (lists provisioned + caches warm by then), off the startup path.
        _timer = new Timer(async _ =>
        {
            try { await RunAsync(); }
            catch (Exception ex) { _log.LogWarning(ex, "[SpContract] self-check threw"); }
        }, null, TimeSpan.FromSeconds(30), Timeout.InfiniteTimeSpan);
        return Task.CompletedTask;
    }

    public Task StopAsync(CancellationToken ct) { _timer?.Dispose(); return Task.CompletedTask; }

    public async Task<ContractCheckResult> RunAsync(CancellationToken ct = default)
    {
        var probes = new List<FieldProbe>
        {
            // Boolean round-trip on Messages.IsRead, BOTH directions — testing 'true' is what catches an
            // always-false reader (the exact Kiota Boolean bug that broke per-message read state).
            await ProbeBoolAsync(true, ct),
            await ProbeBoolAsync(false, ct),
        };

        var result = new ContractCheckResult(probes.All(p => p.Ok), probes);
        foreach (var p in probes)
        {
            if (p.Ok) _log.LogInformation("[SpContract] OK   {Field} ({Type}) — {Detail}", p.Field, p.Type, p.Detail);
            else      _log.LogError       ("[SpContract] FAIL {Field} ({Type}) — {Detail}", p.Field, p.Type, p.Detail);
        }
        _log.LogInformation("[SpContract] self-check {Status} ({Pass}/{Total} fields round-tripped)",
            result.AllOk ? "PASSED" : "FAILED", probes.Count(p => p.Ok), probes.Count);
        Last = result;
        return result;
    }

    private async Task<FieldProbe> ProbeBoolAsync(bool value, CancellationToken ct)
    {
        int? id = null;
        try
        {
            var probe = new MessageRecord
            {
                InquiryId = Sentinel, ConversationId = Sentinel, Channel = "diag", Direction = "in",
                Body = "sp-contract-check", IsRead = value, TimestampUtc = DateTimeOffset.UtcNow.ToString("o"),
            };
            await _sp.WriteMessageAsync(probe, ct);
            id = probe.SpItemId;

            // Read back through the production reader, tolerating brief SP propagation lag.
            MessageRecord? back = null;
            for (int attempt = 0; attempt < 4 && back is null; attempt++)
            {
                if (attempt > 0) await Task.Delay(600, ct);
                back = (await _sp.ReadMessagesByInquiryAsync(Sentinel, 50, ct)).FirstOrDefault(m => m.SpItemId == id);
            }

            var ok = back is not null && back.IsRead == value;
            return new FieldProbe("Messages.IsRead", "Boolean", ok,
                ok ? $"{value} survived the SP round-trip"
                   : $"wrote {value}, read back {(back is null ? "<not found>" : back.IsRead.ToString())} — Boolean type mismatch (Kiota returns native bool, not JsonElement)");
        }
        catch (Exception ex) { return new FieldProbe("Messages.IsRead", "Boolean", false, ex.Message); }
        finally { if (id is int i) { try { await _sp.DeleteMessageItemAsync(i, ct); } catch { /* best-effort cleanup */ } } }
    }
}

public sealed record FieldProbe(string Field, string Type, bool Ok, string Detail);
public sealed record ContractCheckResult(bool AllOk, IReadOnlyList<FieldProbe> Probes);
