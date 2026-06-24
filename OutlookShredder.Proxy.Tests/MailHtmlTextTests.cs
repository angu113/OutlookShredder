using OutlookShredder.Proxy.Services;
using Xunit;

namespace OutlookShredder.Proxy.Tests;

// Pure-function tests for HtmlText.ToPlainText — the capture-time HTML→plain-text conversion.
// The headline case is the QuickBooks/Intuit estimate email (e.g. Peerless Coatings) whose
// <style> CSS used to leak into the stored body as "CSS soup".
public class MailHtmlTextTests
{
    [Fact]
    public void Strips_style_block_contents_keeps_readable_text()
    {
        var html =
            "<html><head><style>table{width:100%}\n@media screen and (max-width:680px){.x{font-size:9px}}</style></head>" +
            "<body><p>Your estimate is ready!</p><p>Total $22,420.00</p></body></html>";

        var text = HtmlText.ToPlainText(html);

        Assert.DoesNotContain("width:100%", text);
        Assert.DoesNotContain("@media", text);
        Assert.DoesNotContain("font-size", text);
        Assert.Contains("Your estimate is ready!", text);
        Assert.Contains("Total $22,420.00", text);
    }

    [Fact]
    public void Strips_script_blocks()
    {
        var text = HtmlText.ToPlainText("<div>Hi<script>var x = 1; alert('boo');</script> there</div>");
        Assert.DoesNotContain("alert", text);
        Assert.DoesNotContain("var x", text);
        Assert.Contains("Hi", text);
        Assert.Contains("there", text);
    }

    [Fact]
    public void Strips_html_comments()
    {
        var text = HtmlText.ToPlainText("<!-- hidden tracking pixel notes -->Visible");
        Assert.DoesNotContain("hidden", text);
        Assert.Contains("Visible", text);
    }

    [Fact]
    public void Decodes_entities_and_breaks_block_tags()
    {
        var text = HtmlText.ToPlainText("<p>A&amp;B</p><p>C</p>");
        Assert.Contains("A&B", text);
        Assert.Contains("\n", text);   // </p> becomes a line break between paragraphs
    }

    [Theory]
    [InlineData(null)]
    [InlineData("")]
    public void Null_or_empty_is_empty(string? input)
        => Assert.Equal("", HtmlText.ToPlainText(input));

    [Fact]
    public void Plain_text_is_trimmed_and_passed_through()
        => Assert.Equal("hello world", HtmlText.ToPlainText("  hello world  "));
}
