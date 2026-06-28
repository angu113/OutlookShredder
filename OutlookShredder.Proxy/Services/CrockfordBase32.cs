namespace OutlookShredder.Proxy.Services;

/// <summary>
/// Crockford Base32 encoding (https://www.crockford.com/base32.html).
/// Alphabet: 0-9, A-Z excluding I, L, O, U — 32 symbols total (no check symbol).
/// Ported from the Shredder WPF client (which can't be referenced from the proxy assembly) for
/// inquiry-id generation (CINQ-XXXXX). Kept byte-for-byte compatible with the client copy.
/// </summary>
internal static class CrockfordBase32
{
    internal const string Alphabet = "0123456789ABCDEFGHJKMNPQRSTVWXYZ";

    public static string Encode(long value, int digits = 4)
    {
        var buf = new char[digits];
        for (int i = digits - 1; i >= 0; i--)
        {
            buf[i] = Alphabet[(int)(value & 31)];
            value >>= 5;
        }
        return new string(buf);
    }

    public static long Decode(string s)
    {
        long value = 0;
        foreach (char c in s.ToUpperInvariant())
        {
            int idx = Alphabet.IndexOf(c);
            if (idx < 0) throw new ArgumentException($"Invalid Crockford Base32 character: '{c}'");
            value = (value << 5) | (uint)idx;
        }
        return value;
    }

    public static bool IsValid(string s)
        => s.Length > 0 && s.All(c => Alphabet.IndexOf(char.ToUpperInvariant(c)) >= 0);
}
