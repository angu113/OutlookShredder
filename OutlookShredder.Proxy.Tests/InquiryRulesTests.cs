using OutlookShredder.Proxy.Models;
using OutlookShredder.Proxy.Services;
using Xunit;

namespace OutlookShredder.Proxy.Tests;

// Locks the pure rules of the SMS customer-inquiry pipeline: carrier keyword classification, HSK#
// validation, CINQ id generation (format + collision retry), threading decisions, and phone normalisation.
public class InquiryRulesTests
{
    // ── Keyword classification ────────────────────────────────────────────────
    [Theory]
    [InlineData("STOP")]
    [InlineData("stop")]
    [InlineData(" Stop ")]
    [InlineData("UNSUBSCRIBE")]
    [InlineData("cancel")]
    [InlineData("END")]
    [InlineData("QUIT")]
    [InlineData("STOPALL")]
    public void Classifies_opt_out_keywords(string body)
        => Assert.Equal(InquiryRules.Keyword.OptOut, InquiryRules.ClassifyKeyword(body));

    [Theory]
    [InlineData("START")]
    [InlineData("yes")]
    [InlineData("UNSTOP")]
    public void Classifies_opt_in_keywords(string body)
        => Assert.Equal(InquiryRules.Keyword.OptIn, InquiryRules.ClassifyKeyword(body));

    [Theory]
    [InlineData("HELP")]
    [InlineData("info")]
    public void Classifies_help_keywords(string body)
        => Assert.Equal(InquiryRules.Keyword.Help, InquiryRules.ClassifyKeyword(body));

    [Theory]
    [InlineData("Do you have 2x4 angle?")]
    [InlineData("please stop sending so fast")]   // keyword embedded in a sentence is NOT a keyword
    [InlineData("yes please send a quote")]
    [InlineData("")]
    [InlineData(null)]
    public void Treats_normal_messages_as_none(string? body)
        => Assert.Equal(InquiryRules.Keyword.None, InquiryRules.ClassifyKeyword(body));

    // ── HSK# validation ───────────────────────────────────────────────────────
    [Theory]
    [InlineData("SO1036432")]
    [InlineData("HSK-SO1036432")]
    [InlineData("PO42")]
    [InlineData("Q7")]
    [InlineData("hsk-so100")]
    public void Accepts_valid_hsk(string s) => Assert.True(InquiryRules.IsValidHsk(s));

    [Theory]
    [InlineData("1036432")]      // no SO/PO/Q
    [InlineData("SO")]           // no digits
    [InlineData("HSK-XO100")]    // wrong letter
    [InlineData("SO 100")]       // space
    [InlineData("")]
    [InlineData(null)]
    public void Rejects_invalid_hsk(string? s) => Assert.False(InquiryRules.IsValidHsk(s));

    // ── CINQ id generation ────────────────────────────────────────────────────
    [Fact]
    public void RandomCinqId_has_correct_shape()
    {
        for (int i = 0; i < 200; i++)
        {
            var id = InquiryRules.RandomCinqId();
            Assert.StartsWith("CINQ-", id);
            var suffix = id["CINQ-".Length..];
            Assert.Equal(5, suffix.Length);
            Assert.True(CrockfordBase32IsValid(suffix), $"invalid Crockford suffix: {suffix}");
        }
    }

    [Fact]
    public void RandomCinqId_is_practically_unique()
    {
        var seen = new HashSet<string>();
        for (int i = 0; i < 2000; i++) seen.Add(InquiryRules.RandomCinqId());
        // 25 bits of entropy over 2000 draws — collisions are vanishingly unlikely.
        Assert.True(seen.Count >= 1995, $"too many collisions: {2000 - seen.Count}");
    }

    [Fact]
    public void NewCinqId_retries_past_collisions()
    {
        var taken = new HashSet<string>();
        var first = InquiryRules.NewCinqId(_ => false);   // nothing taken → returns immediately
        taken.Add(first);
        // Force the first N candidates to look taken, then accept — the generator must keep retrying.
        int calls = 0;
        var id = InquiryRules.NewCinqId(_ => ++calls <= 3);
        Assert.True(calls >= 4);
        Assert.StartsWith("CINQ-", id);
    }

    [Fact]
    public void NewCinqId_throws_when_space_exhausted()
        => Assert.Throws<InvalidOperationException>(() => InquiryRules.NewCinqId(_ => true, maxAttempts: 5));

    // ── Threading decisions ───────────────────────────────────────────────────
    [Fact]
    public void No_prior_inquiry_creates_new()
        => Assert.Equal(InquiryRules.ThreadAction.CreateNew, InquiryRules.DecideThread(null));

    [Theory]
    [InlineData(InquiryStatus.Open)]
    [InlineData(InquiryStatus.Quoted)]
    [InlineData(InquiryStatus.Spam)]
    public void Live_inquiry_appends(string status)
        => Assert.Equal(InquiryRules.ThreadAction.Append,
            InquiryRules.DecideThread(new Inquiry { Status = status }));

    [Fact]
    public void Closed_inquiry_reopens()
        => Assert.Equal(InquiryRules.ThreadAction.Reopen,
            InquiryRules.DecideThread(new Inquiry { Status = InquiryStatus.Closed }));

    // ── Phone normalisation ───────────────────────────────────────────────────
    [Theory]
    [InlineData("+1 (555) 222-3333", "+15552223333")]
    [InlineData("5552223333", "+5552223333")]
    [InlineData("+15552223333", "+15552223333")]
    public void Normalizes_phone_to_e164(string input, string expected)
        => Assert.Equal(expected, InquiryRules.NormalizeE164(input));

    private static bool CrockfordBase32IsValid(string s)
    {
        const string alphabet = "0123456789ABCDEFGHJKMNPQRSTVWXYZ";
        return s.Length > 0 && s.All(c => alphabet.IndexOf(char.ToUpperInvariant(c)) >= 0);
    }
}
