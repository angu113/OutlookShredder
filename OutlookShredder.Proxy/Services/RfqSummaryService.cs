using System.Text;
using System.Text.Json;

namespace OutlookShredder.Proxy.Services;

/// <summary>
/// Produces a short (≤3 bullet) AI summary of an RFQ's quote landscape from a pre-built text input
/// (the Shredder client assembles requested items + each supplier's coverage / prices / regrets).
/// Claude with a coverage-aware system prompt; returns an empty list on any failure so the client
/// can fall back to its deterministic summary.
/// </summary>
public class RfqSummaryService
{
    private readonly IHttpClientFactory          _http;
    private readonly IConfiguration              _config;
    private readonly ILogger<RfqSummaryService>  _log;

    private const string SystemPrompt =
        "You summarize a metal-supplier RFQ quote comparison for a busy purchasing rep. " +
        "Return AT MOST 3 bullet points, one short line each, no preamble and no closing remarks. " +
        "Be specific and concrete — supplier names and dollar figures, no filler. " +
        "Keep it TERSE: aim for under ~16 words per bullet, cut every word that doesn't earn its place " +
        "(\"Vigorous writing is concise\"). A natural voice is fine, but brevity beats chattiness — " +
        "never pad to sound conversational.\n" +
        "CRITICAL — account for COVERAGE: a supplier with a lower TOTAL that regretted/skipped items is " +
        "NOT cheaper; only compare COMPLETE quotes apples-to-apples. Call out who quoted everything, the " +
        "cheapest complete option, any single-sourced item (only one supplier quoted it = no price " +
        "leverage), notable regrets/gaps, and anything the rep should act on. " +
        "Output ONLY the bullets, each starting with '- '.";

    public RfqSummaryService(IHttpClientFactory http, IConfiguration config, ILogger<RfqSummaryService> log)
    {
        _http   = http;
        _config = config;
        _log    = log;
    }

    public async Task<List<string>> SummarizeAsync(string input, CancellationToken ct)
    {
        var apiKey = _config["Anthropic:ApiKey"];
        if (string.IsNullOrWhiteSpace(apiKey) || string.IsNullOrWhiteSpace(input))
            return [];

        var model = _config["Claude:Model"] ?? "claude-sonnet-4-6";
        var body  = JsonSerializer.Serialize(new
        {
            model,
            max_tokens = 400,
            system     = SystemPrompt,
            messages   = new[] { new { role = "user", content = input } },
        });

        try
        {
            using var http = _http.CreateClient();
            http.Timeout = TimeSpan.FromSeconds(30);
            http.DefaultRequestHeaders.Add("x-api-key", apiKey);
            http.DefaultRequestHeaders.Add("anthropic-version", "2023-06-01");

            var resp = await http.PostAsync(
                "https://api.anthropic.com/v1/messages",
                new StringContent(body, Encoding.UTF8, "application/json"), ct);

            if (!resp.IsSuccessStatusCode)
            {
                _log.LogWarning("[RfqSummary] Claude returned {Status}", resp.StatusCode);
                return [];
            }

            var json = await resp.Content.ReadAsStringAsync(ct);
            using var doc = JsonDocument.Parse(json);

            var text = new StringBuilder();
            foreach (var block in doc.RootElement.GetProperty("content").EnumerateArray())
                if (block.TryGetProperty("type", out var t) && t.GetString() == "text" &&
                    block.TryGetProperty("text", out var txt))
                    text.Append(txt.GetString());

            return text.ToString()
                .Split('\n', StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries)
                .Select(l => l.TrimStart('-', '*', '•', ' ').Trim())
                .Where(l => l.Length > 0)
                .Take(3)
                .ToList();
        }
        catch (Exception ex)
        {
            _log.LogWarning(ex, "[RfqSummary] summarize failed");
            return [];
        }
    }
}
