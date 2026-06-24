using OutlookShredder.Proxy.Services;
using Xunit;

namespace OutlookShredder.Proxy.Tests;

// Pure-function tests for MailTaxonomyService.NormalizeLeafPath — the validation behind the
// "add category" surface (POST /api/mail-eval/leaves). No SP reads, fully deterministic.
public class MailTaxonomyLeafTests
{
    [Theory]
    [InlineData("Other/Voicemail",      "Other/Voicemail")]
    [InlineData("  Other/Voicemail  ",  "Other/Voicemail")]
    [InlineData("/Other/Voicemail/",    "Other/Voicemail")]
    [InlineData(" / Other/Voicemail / ","Other/Voicemail")]
    [InlineData("Other",                "Other")]
    public void NormalizeLeafPath_accepts_and_trims(string raw, string expected)
    {
        var (ok, path, error) = MailTaxonomyService.NormalizeLeafPath(raw);
        Assert.True(ok);
        Assert.Null(error);
        Assert.Equal(expected, path);
    }

    [Theory]
    [InlineData(null)]
    [InlineData("")]
    [InlineData("   ")]
    [InlineData("/")]
    [InlineData(" // ")]
    public void NormalizeLeafPath_rejects_empty(string? raw)
    {
        var (ok, path, error) = MailTaxonomyService.NormalizeLeafPath(raw);
        Assert.False(ok);
        Assert.Equal("", path);
        Assert.NotNull(error);
    }

    [Theory]
    [InlineData("@sender:intuit.com")]
    [InlineData("@SENDER:enmark.com")]
    public void NormalizeLeafPath_rejects_reserved_sender_prefix(string raw)
    {
        var (ok, _, error) = MailTaxonomyService.NormalizeLeafPath(raw);
        Assert.False(ok);
        Assert.NotNull(error);
    }
}
