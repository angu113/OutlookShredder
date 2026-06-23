using System.Collections.Concurrent;
using System.Text.Json;
using OutlookShredder.Proxy.Models;

namespace OutlookShredder.Proxy.Services;

/// <summary>
/// Offline, read-only classification eval harness (wip/sidecar/project-classification-eval-harness.md).
/// Iterates the human-labeled MailGoldenLabels corpus, calls the real in-process classifier per item,
/// compares predicted vs gold, and produces per-leaf P/R/F1 + confusion matrix + calibration curve.
///
/// NEVER calls WriteClassificationAsync, the bus, or KickMatcherForCategory.
/// Zero workflow side effects — safe to run at any time.
/// </summary>
public sealed class MailEvalService
{
    private readonly SharePointService   _sp;
    private readonly MailCacheService    _cache;
    private readonly MailClassifierService _classifier;
    private readonly ILogger<MailEvalService> _log;

    // One concurrent run allowed (a full run can be hundreds of AI calls).
    private readonly SemaphoreSlim _runGate = new(1, 1);
    private readonly EvalProgress  _progress = new();

    // Results of the most recent completed run.
    private volatile EvalRunResults? _lastResults;

    private static readonly JsonSerializerOptions _jsonOpts = new()
    {
        PropertyNamingPolicy = JsonNamingPolicy.CamelCase,
        DefaultIgnoreCondition = System.Text.Json.Serialization.JsonIgnoreCondition.WhenWritingNull,
    };

    public MailEvalService(SharePointService sp, MailCacheService cache,
        MailClassifierService classifier, ILogger<MailEvalService> log)
    {
        _sp         = sp;
        _cache      = cache;
        _classifier = classifier;
        _log        = log;
    }

    // ── Public surface ──────────────────────────────────────────────────────────────

    public EvalRunStart StartRun(EvalRunRequest req)
    {
        if (!_progress.TryBegin())
            return new EvalRunStart { AlreadyRunning = true };

        _ = Task.Run(() => RunAsync(req, CancellationToken.None));
        return new EvalRunStart { AlreadyRunning = false };
    }

    public EvalSnapshot GetSnapshot() => _progress.SnapshotNow();

    public EvalRunResults? GetResults() => _lastResults;

    public EvalReport? GetReport() => _lastResults is null ? null : _lastResults.Report;

    // ── Seed golden set from current AI classifications ─────────────────────────────

    /// <summary>
    /// Bootstrap MailGoldenLabels from the current in-memory classifications.
    /// overwrite=true  — patch every row.
    /// overwrite=false — skip rows that already exist.
    /// resumeOnly=true — patch only rows that exist but are missing Subject (picks up a partial run).
    /// </summary>
    public async Task<SeedGoldenResult> SeedGoldenFromCurrentsAsync(bool overwrite, bool resumeOnly, CancellationToken ct)
    {
        var currents    = _cache.GetCurrents();
        var itemLookup  = _cache.GetItems().ToDictionary(i => i.MailItemId, StringComparer.Ordinal);

        // Read existing rows once so we can decide what to skip/target.
        var existingRows = (resumeOnly || !overwrite)
            ? (await _sp.ReadGoldenLabelsAsync(ct)).ToDictionary(g => g.MailItemId, StringComparer.Ordinal)
            : new Dictionary<string, MailGoldenLabelRow>(StringComparer.Ordinal);

        int written = 0, skipped = 0;
        foreach (var c in currents)
        {
            bool hasRow    = existingRows.ContainsKey(c.MailItemId);
            bool hasSubject = hasRow && existingRows[c.MailItemId].Subject.Length > 0;

            if (resumeOnly && hasSubject) { skipped++; continue; }
            if (!overwrite && !resumeOnly && hasRow) { skipped++; continue; }

            var item = itemLookup.GetValueOrDefault(c.MailItemId);
            await _sp.UpsertGoldenLabelAsync(new MailGoldenLabelRow
            {
                MailItemId     = c.MailItemId,
                GoldenCategory = c.CategoryPath,
                Subject        = item?.Subject ?? "",
                FromAddress    = item?.FromAddress ?? "",
                LabeledBy      = "bootstrap",
                LabeledAt      = DateTimeOffset.UtcNow.ToString("o"),
                Notes          = "Bootstrapped from AI classification — human review required.",
            }, ct);
            written++;
            if (written % 25 == 0)
                _log.LogInformation("[MailEval] seed {Written}/{Total}", written, currents.Count);
        }
        _log.LogInformation("[MailEval] seed done: {Written} written, {Skipped} skipped", written, skipped);
        return new SeedGoldenResult { Written = written, Skipped = skipped };
    }

    // ── Runner ──────────────────────────────────────────────────────────────────────

    private async Task RunAsync(EvalRunRequest req, CancellationToken ct)
    {
        var startedAt = DateTimeOffset.UtcNow;
        try
        {
            _log.LogInformation("[MailEval] run started (sampleSize={S}, recordResponses={R})",
                req.SampleSize, req.RecordResponses);

            var labels = await _sp.ReadGoldenLabelsAsync(ct);
            if (labels.Count == 0)
            {
                _log.LogWarning("[MailEval] no golden labels found — run aborted. Call POST /api/mail-eval/seed-golden first.");
                _progress.End();
                return;
            }

            // Pin the leaf set at run start so the confusion matrix has stable axes.
            var leafFilter = req.Leaves is { Count: > 0 } l
                ? l.ToHashSet(StringComparer.OrdinalIgnoreCase)
                : null;

            var corpus = labels
                .Where(g => g.GoldenCategory.Length > 0
                    && (leafFilter is null || leafFilter.Contains(g.GoldenCategory)))
                .ToList();

            if (req.SampleSize.HasValue && req.SampleSize.Value < corpus.Count)
                corpus = Stratified(corpus, req.SampleSize.Value);

            _progress.SetTotal(corpus.Count);
            _log.LogInformation("[MailEval] corpus: {N} items", corpus.Count);

            // Build a lookup: MailItemId → MailItemRow (for classifier input).
            var itemLookup = _cache.GetItems().ToDictionary(i => i.MailItemId, StringComparer.Ordinal);

            var sem     = new SemaphoreSlim(4, 4);   // max 4 concurrent classifier calls
            var results = new ConcurrentBag<EvalResultItem>();

            var tasks = corpus.Select((label, i) => Task.Run(async () =>
            {
                await sem.WaitAsync(ct);
                try
                {
                    var item = itemLookup.GetValueOrDefault(label.MailItemId);
                    var input = item is null ? MinimalInput(label) : ItemToInput(item);
                    MailClassificationResult? pred = null;
                    string? rawResponse = null;
                    string  provider    = "";
                    try
                    {
                        pred = await _classifier.ClassifyAsync(input, ct);
                        provider = pred?.AiProvider ?? "";
                        if (req.RecordResponses) rawResponse = pred?.RawResponse;
                    }
                    catch (Exception ex)
                    {
                        _log.LogWarning(ex, "[MailEval] classify failed for {Id}", label.MailItemId);
                    }
                    var r = new EvalResultItem
                    {
                        MailItemId  = label.MailItemId,
                        Subject     = item?.Subject ?? "",
                        Gold        = label.GoldenCategory,
                        Predicted   = pred?.Category ?? "",
                        Confidence  = pred?.Confidence ?? 0,
                        Provider    = provider,
                        Match       = pred is not null && string.Equals(pred.Category, label.GoldenCategory, StringComparison.OrdinalIgnoreCase),
                        RawResponse = rawResponse,
                    };
                    results.Add(r);
                    _progress.Record(r.Match);
                    if ((results.Count % 10) == 0 || results.Count == corpus.Count)
                        _log.LogInformation("[MailEval] {N}/{T} processed", results.Count, corpus.Count);
                }
                finally { sem.Release(); }
            }, ct)).ToList();

            await Task.WhenAll(tasks);

            var allResults = results.ToList();
            var report     = ComputeMetrics(allResults, startedAt, DateTimeOffset.UtcNow);
            _lastResults   = new EvalRunResults { Items = allResults, Report = report };

            PersistResults(allResults, report);
            _log.LogInformation("[MailEval] run DONE: {N} items, accuracy {Acc:P1}", allResults.Count, report.OverallAccuracy);
        }
        catch (Exception ex) { _log.LogError(ex, "[MailEval] run failed"); }
        finally { _progress.End(); _runGate.Release(); }
    }

    // ── Metrics (pure — unit-testable) ──────────────────────────────────────────────

    public static EvalReport ComputeMetrics(IReadOnlyList<EvalResultItem> results,
        DateTimeOffset startedAt, DateTimeOffset finishedAt)
    {
        var leaves = results.Select(r => r.Gold).Concat(results.Select(r => r.Predicted))
            .Where(s => s.Length > 0).Distinct().OrderBy(s => s).ToList();

        // Per-leaf TP/FP/FN
        var perLeaf = new Dictionary<string, (int Tp, int Fp, int Fn, int Support)>(StringComparer.OrdinalIgnoreCase);
        foreach (var leaf in leaves) perLeaf[leaf] = (0, 0, 0, 0);

        foreach (var r in results)
        {
            if (r.Gold.Length == 0 || r.Predicted.Length == 0) continue;
            bool correct = string.Equals(r.Gold, r.Predicted, StringComparison.OrdinalIgnoreCase);
            var (tp, fp, fn, sup) = perLeaf.GetValueOrDefault(r.Gold);
            perLeaf[r.Gold] = correct
                ? (tp + 1, fp, fn, sup + 1)
                : (tp,     fp, fn + 1, sup + 1);
            if (!correct)
            {
                var (pTp, pFp, pFn, pSup) = perLeaf.GetValueOrDefault(r.Predicted);
                perLeaf[r.Predicted] = (pTp, pFp + 1, pFn, pSup);
            }
        }

        var leafMetrics = perLeaf.Select(kv =>
        {
            var (tp, fp, fn, sup) = kv.Value;
            double precision = (tp + fp) == 0 ? 0 : (double)tp / (tp + fp);
            double recall    = (tp + fn) == 0 ? 0 : (double)tp / (tp + fn);
            double f1        = (precision + recall) == 0 ? 0 : 2 * precision * recall / (precision + recall);
            return new LeafMetric { Leaf = kv.Key, Precision = precision, Recall = recall, F1 = f1, Support = sup };
        }).OrderByDescending(m => m.Support).ToList();

        // Confusion matrix: gold -> predicted -> count
        var confusion = new Dictionary<string, Dictionary<string, int>>(StringComparer.OrdinalIgnoreCase);
        foreach (var r in results)
        {
            if (r.Gold.Length == 0 || r.Predicted.Length == 0) continue;
            if (!confusion.TryGetValue(r.Gold, out var row))
                confusion[r.Gold] = row = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);
            row[r.Predicted] = row.GetValueOrDefault(r.Predicted) + 1;
        }

        // Confidence calibration: 10 decile buckets
        var calibration = Enumerable.Range(0, 10).Select(i => new CalibrationBucket
        {
            LowConfidence  = i * 0.1,
            HighConfidence = (i + 1) * 0.1,
        }).ToList();
        foreach (var r in results.Where(r => r.Predicted.Length > 0))
        {
            int bucket = Math.Min(9, (int)(r.Confidence * 10));
            calibration[bucket].Count++;
            if (r.Match) calibration[bucket].Correct++;
        }
        foreach (var b in calibration)
            b.Accuracy = b.Count == 0 ? null : (double)b.Correct / b.Count;

        // Provider stats
        var byProvider = results.Where(r => r.Provider.Length > 0)
            .GroupBy(r => r.Provider)
            .ToDictionary(g => g.Key, g => new ProviderStat
            {
                Provider = g.Key,
                Count    = g.Count(),
                Correct  = g.Count(r => r.Match),
            });

        int total        = results.Count;
        int correctTotal = results.Count(r => r.Match);

        return new EvalReport
        {
            StartedAt       = startedAt,
            FinishedAt      = finishedAt,
            TotalItems      = total,
            CorrectItems    = correctTotal,
            OverallAccuracy = total == 0 ? 0 : (double)correctTotal / total,
            ByLeaf          = leafMetrics,
            Confusion       = confusion,
            Calibration     = calibration,
            ByProvider      = byProvider.Values.OrderByDescending(p => p.Count).ToList(),
        };
    }

    // ── Helpers ─────────────────────────────────────────────────────────────────────

    private static MailClassifyInput ItemToInput(MailItemRow item) => new()
    {
        Subject         = item.Subject,
        FromAddress     = item.FromAddress,
        FromName        = item.FromName,
        BodyText        = "",   // body not in L1 cache (heavy field); classifier still works on subject+from
        AttachmentNames = [],
    };

    private static MailClassifyInput MinimalInput(MailGoldenLabelRow label) => new()
    {
        Subject     = label.MailItemId,
        FromAddress = "",
        BodyText    = "",
    };

    private static List<MailGoldenLabelRow> Stratified(List<MailGoldenLabelRow> corpus, int n)
    {
        var byLeaf = corpus.GroupBy(r => r.GoldenCategory).ToList();
        int perLeaf = Math.Max(1, n / byLeaf.Count);
        var result = byLeaf.SelectMany(g => g.Take(perLeaf)).Take(n).ToList();
        // Top up to n from remainder if needed.
        if (result.Count < n)
        {
            var taken = result.Select(r => r.MailItemId).ToHashSet(StringComparer.Ordinal);
            result.AddRange(corpus.Where(r => !taken.Contains(r.MailItemId)).Take(n - result.Count));
        }
        return result;
    }

    private void PersistResults(List<EvalResultItem> items, EvalReport report)
    {
        try
        {
            var dir = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
                "ShredderData", "Eval");
            Directory.CreateDirectory(dir);
            var ts   = report.StartedAt.ToString("yyyyMMdd-HHmmss");
            var path = Path.Combine(dir, $"{ts}-results.jsonl");
            using var sw = new StreamWriter(path, append: false, encoding: System.Text.Encoding.UTF8);
            foreach (var r in items)
                sw.WriteLine(JsonSerializer.Serialize(r, _jsonOpts));
            _log.LogInformation("[MailEval] results persisted to {Path}", path);
        }
        catch (Exception ex) { _log.LogWarning(ex, "[MailEval] could not persist results"); }
    }

    // ── Progress ────────────────────────────────────────────────────────────────────

    private sealed class EvalProgress
    {
        private int _running;
        private int _total;
        private int _processed;
        private int _correct;

        public bool TryBegin() => Interlocked.CompareExchange(ref _running, 1, 0) == 0;
        public void SetTotal(int t) => Interlocked.Exchange(ref _total, t);
        public void Record(bool match)
        {
            Interlocked.Increment(ref _processed);
            if (match) Interlocked.Increment(ref _correct);
        }
        public void End() => Interlocked.Exchange(ref _running, 0);

        public EvalSnapshot SnapshotNow()
        {
            int total = _total, processed = _processed, correct = _correct;
            return new EvalSnapshot
            {
                Running  = _running == 1,
                Total    = total,
                Processed = processed,
                Correct  = correct,
                Accuracy = processed == 0 ? null : (double)correct / processed,
            };
        }
    }
}

// ── DTOs ─────────────────────────────────────────────────────────────────────────

public sealed class EvalRunRequest
{
    public int?         SampleSize      { get; set; }
    public List<string> Leaves          { get; set; } = [];
    public bool         RecordResponses { get; set; } = true;
}

public sealed class EvalRunStart
{
    public bool AlreadyRunning { get; set; }
}

public sealed class EvalSnapshot
{
    public bool    Running   { get; set; }
    public int     Total     { get; set; }
    public int     Processed { get; set; }
    public int     Correct   { get; set; }
    public double? Accuracy  { get; set; }
}

public sealed class EvalResultItem
{
    public string  MailItemId  { get; set; } = "";
    public string  Subject     { get; set; } = "";
    public string  Gold        { get; set; } = "";
    public string  Predicted   { get; set; } = "";
    public double  Confidence  { get; set; }
    public string  Provider    { get; set; } = "";
    public bool    Match       { get; set; }
    public string? RawResponse { get; set; }
}

public sealed class EvalRunResults
{
    public List<EvalResultItem> Items  { get; set; } = [];
    public EvalReport           Report { get; set; } = new();
}

public sealed class EvalReport
{
    public DateTimeOffset StartedAt       { get; set; }
    public DateTimeOffset FinishedAt      { get; set; }
    public int            TotalItems      { get; set; }
    public int            CorrectItems    { get; set; }
    public double         OverallAccuracy { get; set; }
    public List<LeafMetric>            ByLeaf      { get; set; } = [];
    public Dictionary<string, Dictionary<string, int>> Confusion { get; set; } = [];
    public List<CalibrationBucket>     Calibration { get; set; } = [];
    public List<ProviderStat>          ByProvider  { get; set; } = [];
}

public sealed class LeafMetric
{
    public string Leaf      { get; set; } = "";
    public double Precision { get; set; }
    public double Recall    { get; set; }
    public double F1        { get; set; }
    public int    Support   { get; set; }
}

public sealed class CalibrationBucket
{
    public double  LowConfidence  { get; set; }
    public double  HighConfidence { get; set; }
    public int     Count          { get; set; }
    public int     Correct        { get; set; }
    public double? Accuracy       { get; set; }
}

public sealed class ProviderStat
{
    public string Provider { get; set; } = "";
    public int    Count    { get; set; }
    public int    Correct  { get; set; }
}

public sealed class SeedGoldenResult
{
    public int Written { get; set; }
    public int Skipped { get; set; }
}
