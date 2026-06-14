using OutlookShredder.Proxy.Models;

namespace OutlookShredder.Proxy.Services;

/// <summary>
/// ShadowCat Payment Reconciliation orchestration (v1, on-demand): parse the OB "Payment In" export
/// and the Heartland transaction export, match the card payments, compute per-pay-type subtotals, and
/// cache the last run. Stateless per run; resolution-status persistence (SharePoint) is layered on
/// separately. Config: <c>ShadowRecon</c> (tolerances + optional Heartland column aliases).
/// </summary>
public class PaymentReconciliationService(IConfiguration config, ILogger<PaymentReconciliationService> log)
{
    private volatile ReconRunResult? _last;

    public string         Status    => _last is null ? "none" : "success";
    public DateTime?      LastRunAt => _last?.RunAt;
    public ReconRunResult? GetLastResult() => _last;

    /// <summary>
    /// Parse + match the two exports for a SINGLE business day. This is a daily rec: the OB export
    /// defines the day (its latest payment date, unless <paramref name="targetDate"/> is given), and
    /// both sides are filtered to that day — so a Heartland "batch download" that spans many days only
    /// reconciles the relevant day rather than flagging every other day as missing. Caches the run.
    /// </summary>
    public ReconRunResult Run(string obCsv, string heartlandCsv, string? obName = null, string? hlName = null, DateTime? targetDate = null)
    {
        var obAll = ObPaymentsCsvParser.Parse(obCsv);
        var hlAll = HeartlandCsvParser.Parse(heartlandCsv, LoadColumnMap());

        // The business day being reconciled: caller-specified, else the OB export's latest date.
        var day = targetDate?.Date
            ?? (obAll.Count > 0 ? obAll.Max(p => p.Date)
            :  (hlAll.Count > 0 ? hlAll.Max(p => p.Date) : DateTime.UtcNow.Date));

        var obDay   = obAll.Where(p => p.Date == day).ToList();
        var obCard  = obDay.Where(p => ObPaymentsCsvParser.IsCardMethod(p.PayType)).ToList();
        var obOther = obDay.Where(p => !ObPaymentsCsvParser.IsCardMethod(p.PayType)).ToList();
        var hlDay   = hlAll.Where(p => p.Date == day).ToList();

        var opts = LoadOptions();
        var result = PaymentMatcher.Match(obCard, hlDay, opts);
        // Explain card-side gaps via OB's cash/check/ACH entries, and list every payment.
        PaymentMatcher.ApplyCrossTypeInsight(result, obOther, opts);

        // Subtotals across ALL OB payment methods for the day, for the detail view.
        result.Subtotals = obDay
            .GroupBy(p => string.IsNullOrWhiteSpace(p.PayType) ? "(unspecified)" : p.PayType!)
            .Select(g => new PayTypeSubtotal { PayType = g.Key, Count = g.Count(), Total = g.Sum(p => p.Amount) })
            .OrderByDescending(s => s.Total)
            .ToList();

        result.RunAt           = DateTime.UtcNow;
        result.TargetDate      = day;
        result.ObCount         = obCard.Count;
        result.ProcessorCount  = hlDay.Count;
        result.ObSource        = obName;
        result.ProcessorSource = hlName;
        foreach (var row in result.Rows)
            if (row.Status is not (ReconRowStatus.Matched or ReconRowStatus.Informational))
                row.Resolution = "open";

        _last = result;
        log.LogInformation(
            "[Recon] Run {Day:yyyy-MM-dd} OB={Ob}(card {Card}) HL={Hl} -> matched={M} possible={P} missingInOb={Mi} missingInProc={Mp} mismatch={Am} ambiguous={Ab}",
            day, obDay.Count, obCard.Count, hlDay.Count, result.Counts.Matched, result.Counts.PossibleMatch,
            result.Counts.MissingInOb, result.Counts.MissingInProcessor, result.Counts.AmountMismatch, result.Counts.Ambiguous);
        return result;
    }

    private ReconMatchOptions LoadOptions() => new()
    {
        AmountTolerance         = config.GetValue("ShadowRecon:AmountTolerance", 0.01m),
        DateToleranceDays       = config.GetValue("ShadowRecon:DateToleranceDays", 1),
        PossibleAmountTolerance = config.GetValue("ShadowRecon:PossibleAmountTolerance", 1.00m),
        PossiblePercent         = config.GetValue("ShadowRecon:PossiblePercent", 0.05m),
    };

    private HeartlandColumnMap LoadColumnMap()
    {
        var map = new HeartlandColumnMap();
        config.GetSection("ShadowRecon:Heartland:Columns").Bind(map); // overrides defaults only if configured
        return map;
    }
}
