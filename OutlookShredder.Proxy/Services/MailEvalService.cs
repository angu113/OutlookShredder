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
    private readonly MailWorkbenchService _workbench;
    private readonly MailRuleService     _rules;
    private readonly ILogger<MailEvalService> _log;

    // One concurrent run allowed (a full run can be hundreds of AI calls) — gated by _progress.TryBegin().
    private readonly EvalProgress  _progress = new();
    // Separate gate for the (cheap, AI-free) rule-impact run so it can run independently of an eval.
    private readonly EvalProgress  _riProgress = new();

    // Results of the most recent completed run.
    private volatile EvalRunResults? _lastResults;
    private volatile RuleImpactReport? _lastImpact;

    private static readonly JsonSerializerOptions _jsonOpts = new()
    {
        PropertyNamingPolicy = JsonNamingPolicy.CamelCase,
        DefaultIgnoreCondition = System.Text.Json.Serialization.JsonIgnoreCondition.WhenWritingNull,
    };

    public MailEvalService(SharePointService sp, MailCacheService cache,
        MailClassifierService classifier, MailWorkbenchService workbench,
        MailRuleService rules, ILogger<MailEvalService> log)
    {
        _sp         = sp;
        _cache      = cache;
        _classifier = classifier;
        _workbench  = workbench;
        _rules      = rules;
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

    // ── Golden-set inspection (read-only, no AI) ────────────────────────────────────

    /// <summary>
    /// Read-only snapshot of the MailGoldenLabels corpus: counts by labeler (bootstrap vs human)
    /// and by category, plus the rows. No AI, no SP writes — answers "how complete is the labeling?"
    /// before a (token-spending) run. A row is human-labeled when LabeledBy is set to anything other
    /// than "bootstrap"; bootstrap rows are the AI's own guess and still need correction.
    /// </summary>
    public async Task<GoldenStatus> GetGoldenStatusAsync(CancellationToken ct)
    {
        var rows = await _sp.ReadGoldenLabelsAsync(ct);
        bool IsBootstrap(MailGoldenLabelRow r) =>
            r.LabeledBy.Length == 0 || string.Equals(r.LabeledBy, "bootstrap", StringComparison.OrdinalIgnoreCase);

        return new GoldenStatus
        {
            Total         = rows.Count,
            Bootstrap     = rows.Count(IsBootstrap),
            HumanLabeled  = rows.Count(r => !IsBootstrap(r)),
            BlankCategory = rows.Count(r => r.GoldenCategory.Length == 0),
            ByLabeledBy   = rows.GroupBy(r => r.LabeledBy.Length == 0 ? "(blank)" : r.LabeledBy)
                                .ToDictionary(g => g.Key, g => g.Count()),
            ByCategory    = rows.Where(r => r.GoldenCategory.Length > 0)
                                .GroupBy(r => r.GoldenCategory)
                                .OrderByDescending(g => g.Count())
                                .ToDictionary(g => g.Key, g => g.Count()),
            Rows          = rows.Select(r => new GoldenRowLite
            {
                MailItemId     = r.MailItemId,
                Subject        = r.Subject,
                GoldenCategory = r.GoldenCategory,
                LabeledBy      = r.LabeledBy,
            }).ToList(),
        };
    }

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

            // Evaluate ONLY human-corrected rows by default. Bootstrap rows are the AI's own past
            // guess, so scoring against them measures the model agreeing with itself (~100%, useless).
            // IncludeBootstrap=true is a plumbing smoke-test escape hatch only.
            static bool IsHuman(MailGoldenLabelRow g) =>
                g.LabeledBy.Length > 0 && !string.Equals(g.LabeledBy, "bootstrap", StringComparison.OrdinalIgnoreCase);

            var corpus = labels
                .Where(g => g.GoldenCategory.Length > 0
                    && (req.IncludeBootstrap || IsHuman(g))
                    && (leafFilter is null || leafFilter.Contains(g.GoldenCategory)))
                .ToList();

            if (corpus.Count == 0)
            {
                _log.LogWarning("[MailEval] {Total} golden rows but 0 are human-labeled — nothing to score. " +
                    "Correct some labels in the eval UI first (a run over bootstrap labels is self-agreement).", labels.Count);
                _progress.End();
                return;
            }

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
                    var input = await BuildInputAsync(label, itemLookup, ct);
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
                        Subject     = input.Subject ?? "",
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
        finally { _progress.End(); }
    }

    // ── Rule impact: deterministic re-run of the CURRENT ruleset over existing items ─────
    // Answers "if I add/edit a rule, what flips?" with ZERO AI cost — it builds the SAME signals
    // production builds on capture and runs the pure rule engine. Reports old→new category per matched
    // item + the resulting distribution. Optionally APPLIES (writes a versioned classification,
    // AiProvider="rule") for changed items — no matcher kick (analysis context, not a capture).
    // This is the toolset reuse the user asked for: corrections become rules, the engine re-runs.

    public EvalRunStart StartRuleImpact(RuleImpactRequest req)
    {
        if (!_riProgress.TryBegin()) return new EvalRunStart { AlreadyRunning = true };
        _ = Task.Run(() => RuleImpactAsync(req, CancellationToken.None));
        return new EvalRunStart { AlreadyRunning = false };
    }

    public EvalSnapshot      GetImpactSnapshot() => _riProgress.SnapshotNow();
    public RuleImpactReport? GetImpactReport()   => _lastImpact;

    private async Task RuleImpactAsync(RuleImpactRequest req, CancellationToken ct)
    {
        try
        {
            var currents = _cache.GetCurrents().ToDictionary(c => c.MailItemId, c => c.CategoryPath ?? "", StringComparer.Ordinal);
            var scopeCat = string.IsNullOrWhiteSpace(req.Category) ? null : req.Category.Trim();
            var items = _cache.GetItems()
                .Where(i => scopeCat is null
                    || string.Equals(currents.GetValueOrDefault(i.MailItemId, ""), scopeCat, StringComparison.OrdinalIgnoreCase))
                .ToList();
            _riProgress.SetTotal(items.Count);
            _log.LogInformation("[MailEval] rule-impact start: {N} items (scope={Scope}, dryRun={Dry})",
                items.Count, scopeCat ?? "ALL", req.DryRun);

            var changes = new ConcurrentBag<RuleImpactItem>();
            var newDist = new ConcurrentDictionary<string, int>(StringComparer.OrdinalIgnoreCase);
            int matched = 0, changed = 0, applied = 0;
            var sem = new SemaphoreSlim(4, 4);

            var tasks = items.Select(i => Task.Run(async () =>
            {
                await sem.WaitAsync(ct);
                try
                {
                    var cur = currents.GetValueOrDefault(i.MailItemId, "");
                    MailRule? rule = null;
                    try
                    {
                        var input = await _sp.GetClassifyInputAsync(i.MailItemId, ct);
                        if (input is not null)
                            rule = await _rules.FirstMatchAsync(
                                _workbench.BuildRuleSignals(input.FromAddress, input.Subject, input.BodyText, input.AttachmentNames), ct);
                    }
                    catch (Exception ex) { _log.LogDebug(ex, "[MailEval] impact eval failed for {Id}", i.MailItemId); }

                    // No rule match → would fall through to the AI; in a deterministic preview the item is
                    // shown unchanged (we do NOT re-run the AI here — that's the eval runner's job).
                    var newCat = rule?.CategoryPath ?? cur;
                    newDist.AddOrUpdate(newCat, 1, (_, n) => n + 1);
                    bool flip = rule is not null && !string.Equals(newCat, cur, StringComparison.OrdinalIgnoreCase);
                    if (rule is not null) Interlocked.Increment(ref matched);
                    if (flip)
                    {
                        Interlocked.Increment(ref changed);
                        changes.Add(new RuleImpactItem
                        {
                            MailItemId = i.MailItemId, Subject = i.Subject, FromAddress = i.FromAddress,
                            OldCategory = cur, NewCategory = newCat, Rule = rule!.Name,
                        });
                        if (!req.DryRun)
                        {
                            try
                            {
                                await _sp.WriteClassificationAsync(i.MailItemId, new MailClassificationResult
                                {
                                    Category = newCat, Confidence = 1.0, AiProvider = "rule", AiModel = rule.Name,
                                    Reasoning = $"Rule '{rule.Name}' applied via eval reclassify",
                                }, ct);
                                Interlocked.Increment(ref applied);
                            }
                            catch (Exception ex) { _log.LogWarning(ex, "[MailEval] apply write failed for {Id}", i.MailItemId); }
                        }
                    }
                    _riProgress.Record(flip);
                }
                finally { sem.Release(); }
            }, ct)).ToList();
            await Task.WhenAll(tasks);

            _lastImpact = new RuleImpactReport
            {
                Scope           = scopeCat ?? "ALL",
                DryRun          = req.DryRun,
                TotalScanned    = items.Count,
                RuleMatched     = matched,
                Changed         = changed,
                Applied         = applied,
                Changes         = changes.OrderBy(c => c.OldCategory, StringComparer.OrdinalIgnoreCase)
                                         .ThenBy(c => c.NewCategory, StringComparer.OrdinalIgnoreCase).ToList(),
                NewDistribution = newDist.OrderByDescending(kv => kv.Value).ToDictionary(kv => kv.Key, kv => kv.Value),
            };
            _log.LogInformation("[MailEval] rule-impact DONE: {M} matched, {C} changed, {A} applied", matched, changed, applied);
        }
        catch (Exception ex) { _log.LogError(ex, "[MailEval] rule-impact failed"); }
        finally { _riProgress.End(); }
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

    /// <summary>
    /// Build the classifier input with the SAME signal production gets — subject, from, to, BODY, and
    /// attachment names — via GetClassifyInputAsync (the exact path the production reclassify uses, body
    /// read from the MailItems list — no .eml parse). Falls back to the L1 cache (subject/from only)
    /// then a minimal input. Without the body the eval would score a weaker classifier than the one that
    /// actually runs on capture, making the baseline misleading.
    /// </summary>
    private async Task<MailClassifyInput> BuildInputAsync(MailGoldenLabelRow label,
        Dictionary<string, MailItemRow> itemLookup, CancellationToken ct)
    {
        try
        {
            var input = await _sp.GetClassifyInputAsync(label.MailItemId, ct);
            if (input is not null) return input;
        }
        catch (Exception ex) { _log.LogDebug(ex, "[MailEval] classify-input fetch failed for {Id}; using cache fallback", label.MailItemId); }

        var item = itemLookup.GetValueOrDefault(label.MailItemId);
        return item is null ? MinimalInput(label) : ItemToInput(item);
    }

    /// <summary>
    /// Apply a single human correction to the golden set: set GoldenCategory + LabeledBy (the row stops
    /// being "bootstrap" and becomes eligible for scoring). Subject/From are passed through from the
    /// caller (the UI already has them) so we don't re-read the whole list per save.
    /// </summary>
    public async Task PatchGoldenAsync(string mailItemId, string goldenCategory,
        string? subject, string? fromAddress, string? labeledBy, CancellationToken ct)
    {
        await _sp.UpsertGoldenLabelAsync(new MailGoldenLabelRow
        {
            MailItemId     = mailItemId,
            GoldenCategory = goldenCategory,
            Subject        = subject ?? "",
            FromAddress    = fromAddress ?? "",
            LabeledBy      = string.IsNullOrWhiteSpace(labeledBy) ? "human" : labeledBy,
            LabeledAt      = DateTimeOffset.UtcNow.ToString("o"),
            Notes          = "Human-corrected via eval UI.",
        }, ct);
        _log.LogInformation("[MailEval] golden patched: {Id} -> {Cat} by {By}", mailItemId, goldenCategory,
            string.IsNullOrWhiteSpace(labeledBy) ? "human" : labeledBy);
    }

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
    public int?         SampleSize       { get; set; }
    public List<string> Leaves           { get; set; } = [];
    public bool         RecordResponses  { get; set; } = true;
    /// <summary>Smoke-test escape hatch only: score bootstrap (AI-guess) rows too. Default false.</summary>
    public bool         IncludeBootstrap { get; set; }
}

public sealed class GoldenPatchRequest
{
    public string  MailItemId     { get; set; } = "";
    public string  GoldenCategory { get; set; } = "";
    public string? Subject        { get; set; }
    public string? FromAddress    { get; set; }
    public string? LabeledBy      { get; set; }
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

public sealed class GoldenStatus
{
    public int Total         { get; set; }
    public int Bootstrap     { get; set; }   // LabeledBy blank/"bootstrap" — AI guess, needs human review
    public int HumanLabeled  { get; set; }   // LabeledBy set to a real labeler
    public int BlankCategory { get; set; }   // GoldenCategory empty (unusable until labeled)
    public Dictionary<string, int> ByLabeledBy { get; set; } = [];
    public Dictionary<string, int> ByCategory  { get; set; } = [];
    public List<GoldenRowLite>     Rows        { get; set; } = [];
}

public sealed class GoldenRowLite
{
    public string MailItemId     { get; set; } = "";
    public string Subject        { get; set; } = "";
    public string GoldenCategory { get; set; } = "";
    public string LabeledBy      { get; set; } = "";
}

public sealed class RuleImpactRequest
{
    /// <summary>Restrict to items currently in this category; null/empty = the whole corpus.</summary>
    public string? Category { get; set; }
    /// <summary>Preview only (default). False = write the rule result for changed items (no matcher kick).</summary>
    public bool    DryRun   { get; set; } = true;
}

public sealed class RuleImpactItem
{
    public string MailItemId  { get; set; } = "";
    public string Subject     { get; set; } = "";
    public string FromAddress { get; set; } = "";
    public string OldCategory { get; set; } = "";
    public string NewCategory { get; set; } = "";
    public string Rule        { get; set; } = "";
}

public sealed class RuleImpactReport
{
    public string Scope        { get; set; } = "";
    public bool   DryRun       { get; set; }
    public int    TotalScanned { get; set; }
    public int    RuleMatched  { get; set; }   // items a rule fired on
    public int    Changed      { get; set; }   // items whose category would flip
    public int    Applied      { get; set; }   // items actually written (apply runs only)
    public List<RuleImpactItem>     Changes         { get; set; } = [];
    public Dictionary<string, int>  NewDistribution { get; set; } = [];
}
