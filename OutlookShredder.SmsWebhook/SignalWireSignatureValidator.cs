using System.Security.Cryptography;
using System.Text;

namespace OutlookShredder.SmsWebhook;

/// <summary>
/// Validates the HMAC-SHA1 signature SignalWire (Twilio-compatible) places on inbound webhooks.
/// Signature = Base64(HMACSHA1(authToken, fullUrl + each POST param appended as key+value, sorted by key)).
/// Header is <c>X-SignalWire-Signature</c> (or the legacy <c>X-Twilio-Signature</c>). Constant-time compare.
///
/// NOTE: kept byte-for-byte in sync with the proxy's copy
/// (OutlookShredder.Proxy/Services/SignalWireSignatureValidator.cs). Edit both together.
/// </summary>
public static class SignalWireSignatureValidator
{
    public static bool IsValid(string url, IEnumerable<KeyValuePair<string, string>> formParams,
        string? signatureHeader, string? authToken)
    {
        if (string.IsNullOrEmpty(signatureHeader) || string.IsNullOrEmpty(authToken)) return false;

        var sb = new StringBuilder(url);
        foreach (var kv in formParams.OrderBy(p => p.Key, StringComparer.Ordinal))
            sb.Append(kv.Key).Append(kv.Value);

        using var hmac = new HMACSHA1(Encoding.UTF8.GetBytes(authToken));
        var expected = Convert.ToBase64String(hmac.ComputeHash(Encoding.UTF8.GetBytes(sb.ToString())));

        return CryptographicOperations.FixedTimeEquals(
            Encoding.UTF8.GetBytes(expected), Encoding.UTF8.GetBytes(signatureHeader));
    }
}
