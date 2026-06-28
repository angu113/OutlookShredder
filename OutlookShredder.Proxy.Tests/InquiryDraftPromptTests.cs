using System.Text.Json;
using OutlookShredder.Proxy.Models;
using OutlookShredder.Proxy.Services;
using Xunit;

namespace OutlookShredder.Proxy.Tests;

// Locks the pure prompt-assembly + result-coercion of InquiryDraftService (the AI call itself is not
// unit-tested — these guard the deterministic glue around it).
public class InquiryDraftPromptTests
{
    private static MessageRecord In(string body)  => new() { Direction = "in",  Body = body };
    private static MessageRecord Out(string body) => new() { Direction = "out", Body = body };

    // ── Transcript ────────────────────────────────────────────────────────────
    [Fact]
    public void Transcript_labels_customer_and_shop()
    {
        var t = InquiryDraftPrompt.BuildTranscript([In("got 2x4 angle?"), Out("yes we do"), In("how much")]);
        Assert.Equal("Customer: got 2x4 angle?\nShop: yes we do\nCustomer: how much", NormalizeNewlines(t));
    }

    [Fact]
    public void Transcript_keeps_only_the_last_max()
    {
        var msgs = Enumerable.Range(0, 20).Select(i => In($"m{i}")).ToList();
        var t = InquiryDraftPrompt.BuildTranscript(msgs, max: 3);
        var lines = t.Split('\n', StringSplitOptions.RemoveEmptyEntries);
        Assert.Equal(3, lines.Length);
        Assert.Contains("m19", t);
        Assert.DoesNotContain("m16", t);
    }

    [Fact]
    public void Transcript_skips_blank_bodies_and_flattens_newlines()
    {
        var t = InquiryDraftPrompt.BuildTranscript([In("  "), In("line1\nline2")]);
        Assert.Equal("Customer: line1 line2", NormalizeNewlines(t));
    }

    // ── User text ─────────────────────────────────────────────────────────────
    [Fact]
    public void UserText_includes_latest_message_and_optional_context()
    {
        var input = new InquiryDraftInput("do you have 1in plate?", "Customer: hi\r\nShop: hello",
            ["HSK-SO1036432"], "VIP customer");
        var text = InquiryDraftPrompt.BuildUserText(input);
        Assert.Contains("Conversation so far:", text);
        Assert.Contains("HSK-SO1036432", text);
        Assert.Contains("VIP customer", text);
        Assert.Contains("Latest customer message:", text);
        Assert.EndsWith("do you have 1in plate?", text);
    }

    [Fact]
    public void UserText_omits_context_sections_when_empty()
    {
        var text = InquiryDraftPrompt.BuildUserText(new InquiryDraftInput("hi", "", [], null));
        Assert.DoesNotContain("Conversation so far:", text);
        Assert.DoesNotContain("Linked order/quote refs:", text);
        Assert.DoesNotContain("Operator notes:", text);
        Assert.Contains("Latest customer message:", text);
    }

    // ── Coercion ──────────────────────────────────────────────────────────────
    [Theory]
    [InlineData("Quote", "Quote")]
    [InlineData("quote", "Quote")]
    [InlineData("OrderStatus", "OrderStatus")]
    [InlineData("nonsense", "Other")]
    [InlineData(null, "Other")]
    public void Coerces_intent(string? raw, string expected)
        => Assert.Equal(expected, InquiryDraftPrompt.CoerceIntent(raw));

    [Theory]
    [InlineData("High", "High")]
    [InlineData("low", "Low")]
    [InlineData("whenever", "Normal")]
    [InlineData(null, "Normal")]
    public void Coerces_urgency(string? raw, string expected)
        => Assert.Equal(expected, InquiryDraftPrompt.CoerceUrgency(raw));

    // ── Result mapping ────────────────────────────────────────────────────────
    [Fact]
    public void MapResult_parses_and_coerces_a_tool_payload()
    {
        var json = """{"reply":"  We can cut that — I'll confirm pricing.  ","intent":"quote","urgency":"HIGH","needsQuote":true}""";
        var r = InquiryDraftPrompt.MapResult(json, "claude-haiku-4-5-20251001");
        Assert.NotNull(r);
        Assert.Equal("We can cut that — I'll confirm pricing.", r!.Reply);   // trimmed
        Assert.Equal("Quote", r.Intent);
        Assert.Equal("High", r.Urgency);
        Assert.True(r.NeedsQuote);
        Assert.Equal("claude-haiku-4-5-20251001", r.AiModel);
    }

    [Theory]
    [InlineData("""{"reply":"","intent":"Question","urgency":"Low","needsQuote":false}""")]   // empty reply
    [InlineData("""{"intent":"Question","urgency":"Low","needsQuote":false}""")]              // missing reply
    public void MapResult_returns_null_without_a_reply(string json)
        => Assert.Null(InquiryDraftPrompt.MapResult(json, "m"));

    private static string NormalizeNewlines(string s) => s.Replace("\r\n", "\n");
}
