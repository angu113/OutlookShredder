using System.Security.Cryptography;
using System.Text;
using OutlookShredder.Proxy.Services;
using Xunit;

namespace OutlookShredder.Proxy.Tests;

// Locks the SignalWire/Twilio HMAC-SHA1 webhook-signature scheme: Base64(HMACSHA1(token,
// url + each param appended key+value, sorted by key)). Constant-time compare; order-independent input.
public class SignalWireSignatureValidatorTests
{
    private const string Url   = "https://sms.example.com/api/sms/inbound";
    private const string Token = "test-auth-token";

    private static readonly (string K, string V)[] Params =
    {
        ("To", "+15550001111"), ("From", "+15552223333"), ("Body", "hello there"), ("MessageSid", "SM123"),
    };

    private static string Sign((string K, string V)[] prms, string token)
    {
        var sb = new StringBuilder(Url);
        foreach (var (k, v) in prms.OrderBy(p => p.K, StringComparer.Ordinal)) sb.Append(k).Append(v);
        using var h = new HMACSHA1(Encoding.UTF8.GetBytes(token));
        return Convert.ToBase64String(h.ComputeHash(Encoding.UTF8.GetBytes(sb.ToString())));
    }

    private static IEnumerable<KeyValuePair<string, string>> Kvs((string K, string V)[] prms)
        => prms.Select(p => new KeyValuePair<string, string>(p.K, p.V));

    [Fact]
    public void Accepts_a_correctly_signed_request()
        => Assert.True(SignalWireSignatureValidator.IsValid(Url, Kvs(Params), Sign(Params, Token), Token));

    [Fact]
    public void Rejects_a_tampered_param()
    {
        var sig = Sign(Params, Token);
        var tampered = Params.Select(p => p.K == "Body" ? (p.K, "different") : p).ToArray();
        Assert.False(SignalWireSignatureValidator.IsValid(Url, Kvs(tampered), sig, Token));
    }

    [Fact]
    public void Rejects_a_wrong_token()
        => Assert.False(SignalWireSignatureValidator.IsValid(Url, Kvs(Params), Sign(Params, Token), "other-token"));

    [Fact]
    public void Rejects_missing_signature_or_token()
    {
        Assert.False(SignalWireSignatureValidator.IsValid(Url, Kvs(Params), null, Token));
        Assert.False(SignalWireSignatureValidator.IsValid(Url, Kvs(Params), Sign(Params, Token), null));
    }
}
