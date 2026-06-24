using System.Text.RegularExpressions;

namespace OutlookShredder.Proxy.Services;

/// <summary>
/// HTML → readable plain-text conversion for captured email bodies. The defining behaviour over a
/// naive tag-strip is that it removes the CONTENTS of &lt;style&gt;/&lt;script&gt; blocks and HTML
/// comments BEFORE stripping tags, so CSS/JS never leaks into the stored body. Without this, templated
/// HTML mail (e.g. the QuickBooks/Intuit estimate notifications suppliers like Peerless send) collapses
/// to CSS soup in the body and pollutes the AI-extraction input. Safe on already-plain text.
/// </summary>
public static class HtmlText
{
    // Strip whole style/script blocks + comments (contents included) first. Singleline so '.' spans newlines.
    private static readonly Regex StyleScriptComment = new(
        @"<style\b[^>]*>.*?</style>|<script\b[^>]*>.*?</script>|<!--.*?-->",
        RegexOptions.Compiled | RegexOptions.IgnoreCase | RegexOptions.Singleline);

    // Block-level tags that should become a line break before generic tags are flattened to spaces.
    private static readonly Regex LineBreak = new(
        @"<br\s*/?>|</p>|</div>|</tr>|</li>|</h[1-6]>",
        RegexOptions.Compiled | RegexOptions.IgnoreCase);

    private static readonly Regex Tag       = new(@"<[^>]+>",   RegexOptions.Compiled);
    private static readonly Regex HorizWs   = new(@"[ \t]{2,}", RegexOptions.Compiled);
    private static readonly Regex ExcessNl  = new(@"\n{3,}",    RegexOptions.Compiled);

    /// <summary>Convert an HTML body to readable plain text (style/script/comment-safe).</summary>
    public static string ToPlainText(string? html)
    {
        if (string.IsNullOrEmpty(html)) return "";
        var s = StyleScriptComment.Replace(html, " ");
        s = LineBreak.Replace(s, "\n");
        s = Tag.Replace(s, " ");
        s = System.Net.WebUtility.HtmlDecode(s);
        s = HorizWs.Replace(s, " ");
        s = ExcessNl.Replace(s, "\n\n");
        return s.Trim();
    }
}
