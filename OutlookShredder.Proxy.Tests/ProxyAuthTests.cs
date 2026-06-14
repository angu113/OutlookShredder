using System.Text;
using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.Logging.Abstractions;
using OutlookShredder.Proxy.Services;
using Xunit;

namespace OutlookShredder.Proxy.Tests;

// WS2 local-REST-auth contract. The canonical-string vector here is mirrored verbatim by the client
// (Shredder.Tests/ProxyAuthHandlerTests) — if either side drifts, one of the two suites fails. Also
// pins VerifyCore's outcomes (ok / skew / replay / bad-sig / missing) and the query canonicalization.
public class ProxyAuthTests
{
    // ── Shared canonical vector (keep identical to the client test) ─────────────
    private const string VMethod = "DELETE";
    private const string VPath   = "/api/sr/AW0001";
    private const string VQuery  = "?b=2&a=1";          // canonicalizes to "a=1&b=2"
    private const string VTs     = "1700000000000";
    private const string VNonce  = "bm9uY2U=";          // base64("nonce")
    private const string ExpectedCanonical =
        "DELETE\n/api/sr/AW0001\na=1&b=2\n1700000000000\nbm9uY2U=";

    private static ProxyAuthService Make(int skewSeconds = 30)
    {
        var config = new ConfigurationBuilder().AddInMemoryCollection(new Dictionary<string, string?>
        {
            ["Proxy:AuthEnforce"]       = "true",
            ["Proxy:AuthSkewSeconds"]   = skewSeconds.ToString(),
            ["Proxy:RateLimitPerSecond"] = "300",
        }).Build();
        return new ProxyAuthService(config, NullLogger<ProxyAuthService>.Instance);
    }

    [Fact]
    public void Canonical_matches_vector()
        => Assert.Equal(ExpectedCanonical, ProxyAuthService.BuildCanonical(VMethod, VPath, VQuery, VTs, VNonce));

    [Theory]
    [InlineData(null, "")]
    [InlineData("", "")]
    [InlineData("?", "")]
    [InlineData("?a=1", "a=1")]
    [InlineData("?b=2&a=1", "a=1&b=2")]          // sorted ordinally
    [InlineData("a=1&b=2", "a=1&b=2")]           // leading '?' optional
    [InlineData("?z=9&z=1", "z=1&z=9")]          // duplicate keys sort by full segment
    public void CanonicalQuery_normalizes(string? raw, string expected)
        => Assert.Equal(expected, ProxyAuthService.CanonicalQuery(raw));

    private static string Sign(byte[] token, string method, string path, string? query, string ts, string nonce)
        => Convert.ToBase64String(ProxyAuthService.ComputeSig(token,
            ProxyAuthService.BuildCanonical(method, path, query, ts, nonce)));

    private static string NowTs() => DateTimeOffset.UtcNow.ToUnixTimeMilliseconds().ToString();

    [Fact]
    public void Verify_valid_signature_is_ok()
    {
        var svc = Make();
        var token = Encoding.UTF8.GetBytes("0123456789abcdef0123456789abcdef");
        svc.SetTokenForTests(token);
        var ts = NowTs();
        var sig = Sign(token, VMethod, VPath, VQuery, ts, VNonce);
        Assert.Equal(ProxyAuthService.AuthResult.Ok, svc.VerifyCore(VMethod, VPath, VQuery, ts, VNonce, sig));
    }

    [Fact]
    public void Verify_replayed_nonce_is_rejected()
    {
        var svc = Make();
        var token = Encoding.UTF8.GetBytes("0123456789abcdef0123456789abcdef");
        svc.SetTokenForTests(token);
        var ts = NowTs();
        var sig = Sign(token, VMethod, VPath, VQuery, ts, VNonce);
        Assert.Equal(ProxyAuthService.AuthResult.Ok,     svc.VerifyCore(VMethod, VPath, VQuery, ts, VNonce, sig));
        Assert.Equal(ProxyAuthService.AuthResult.Replay, svc.VerifyCore(VMethod, VPath, VQuery, ts, VNonce, sig));
    }

    [Fact]
    public void Verify_stale_timestamp_is_skew()
    {
        var svc = Make(skewSeconds: 30);
        var token = Encoding.UTF8.GetBytes("0123456789abcdef0123456789abcdef");
        svc.SetTokenForTests(token);
        var oldTs = (DateTimeOffset.UtcNow.ToUnixTimeMilliseconds() - 60_000).ToString();   // 60s old
        var sig = Sign(token, VMethod, VPath, VQuery, oldTs, VNonce);
        Assert.Equal(ProxyAuthService.AuthResult.Skew, svc.VerifyCore(VMethod, VPath, VQuery, oldTs, VNonce, sig));
    }

    [Fact]
    public void Verify_wrong_key_is_bad_sig()
    {
        var svc = Make();
        svc.SetTokenForTests(Encoding.UTF8.GetBytes("0123456789abcdef0123456789abcdef"));
        var ts = NowTs();
        var sigFromOtherKey = Sign(Encoding.UTF8.GetBytes("XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX"),
            VMethod, VPath, VQuery, ts, VNonce);
        Assert.Equal(ProxyAuthService.AuthResult.BadSig, svc.VerifyCore(VMethod, VPath, VQuery, ts, VNonce, sigFromOtherKey));
    }

    [Fact]
    public void Verify_tampered_path_is_bad_sig()
    {
        var svc = Make();
        var token = Encoding.UTF8.GetBytes("0123456789abcdef0123456789abcdef");
        svc.SetTokenForTests(token);
        var ts = NowTs();
        var sig = Sign(token, VMethod, VPath, VQuery, ts, VNonce);
        // Same signature, different path -> mismatch.
        Assert.Equal(ProxyAuthService.AuthResult.BadSig,
            svc.VerifyCore(VMethod, "/api/sr/ZZ9999", VQuery, ts, VNonce, sig));
    }

    [Fact]
    public void Verify_missing_headers_is_rejected()
    {
        var svc = Make();
        svc.SetTokenForTests(Encoding.UTF8.GetBytes("0123456789abcdef0123456789abcdef"));
        Assert.Equal(ProxyAuthService.AuthResult.MissingHeaders,
            svc.VerifyCore(VMethod, VPath, VQuery, ts: "", nonce: "", sig: ""));
    }

    [Fact]
    public void Verify_no_token_is_missing_token()
    {
        var svc = Make();   // SetTokenForTests not called
        Assert.Equal(ProxyAuthService.AuthResult.MissingToken,
            svc.VerifyCore(VMethod, VPath, VQuery, NowTs(), VNonce, "x"));
    }
}
