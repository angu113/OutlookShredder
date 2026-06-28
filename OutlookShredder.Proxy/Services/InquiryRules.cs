using System.Security.Cryptography;
using System.Text.RegularExpressions;
using OutlookShredder.Proxy.Models;

namespace OutlookShredder.Proxy.Services;

/// <summary>
/// Pure, side-effect-free rules for the SMS customer-inquiry pipeline: carrier keyword classification
/// (STOP/HELP/START families), HSK# validation, CINQ id generation, threading decisions, and phone
/// normalisation. Kept free of SharePoint / Service Bus so each rule is unit-testable in isolation
/// (the ingest orchestration lives in <see cref="InquiryService"/>).
/// </summary>
public static partial class InquiryRules
{
    public enum Keyword { None, OptOut, OptIn, Help }

    // Carrier compliance keywords — matched only when the message is EXACTLY the keyword (trimmed,
    // case-insensitive), the same way SignalWire/Twilio recognise them. A keyword embedded in a longer
    // sentence ("please stop sending so fast") is a normal message, not an opt-out.
    private static readonly HashSet<string> OptOutWords =
        new(StringComparer.OrdinalIgnoreCase) { "STOP", "STOPALL", "UNSUBSCRIBE", "CANCEL", "END", "QUIT" };
    private static readonly HashSet<string> OptInWords =
        new(StringComparer.OrdinalIgnoreCase) { "START", "YES", "UNSTOP" };
    private static readonly HashSet<string> HelpWords =
        new(StringComparer.OrdinalIgnoreCase) { "HELP", "INFO" };

    public static Keyword ClassifyKeyword(string? body)
    {
        var t = body?.Trim() ?? "";
        if (t.Length == 0) return Keyword.None;
        if (OptOutWords.Contains(t)) return Keyword.OptOut;
        if (OptInWords.Contains(t))  return Keyword.OptIn;
        if (HelpWords.Contains(t))   return Keyword.Help;
        return Keyword.None;
    }

    // HSK# = a customer order/quote reference. Format-only validation (we don't verify it exists in the ERP):
    // an optional "HSK-" prefix, then SO|PO|Q, then digits — e.g. "SO1036432", "HSK-SO1036432", "Q42".
    [GeneratedRegex(@"^(HSK-)?(SO|PO|Q)\d+$", RegexOptions.IgnoreCase)]
    private static partial Regex HskRegex();

    public static bool IsValidHsk(string? s)
        => !string.IsNullOrWhiteSpace(s) && HskRegex().IsMatch(s.Trim());

    /// <summary>Normalises a phone to an E.164-ish key ("+" + digits only) — the same shape as
    /// <see cref="MessagingService.SmsConvId"/> uses, so a contact / inquiry key lines up with the
    /// message ConversationId for the same number.</summary>
    public static string NormalizeE164(string? phone)
        => "+" + new string((phone ?? "").Where(char.IsDigit).ToArray());

    /// <summary>A fresh CINQ inquiry id: "CINQ-" + 5 random Crockford Base32 chars (25 bits of entropy,
    /// ~33.5M space, no check symbol). Collision handling is the caller's job (see <see cref="NewCinqId"/>).</summary>
    public static string RandomCinqId()
    {
        // 25 bits = 5 Crockford chars. Draw a uint and mask to 25 bits.
        Span<byte> buf = stackalloc byte[4];
        RandomNumberGenerator.Fill(buf);
        long value = BitConverter.ToUInt32(buf) & ((1L << 25) - 1);
        return "CINQ-" + CrockfordBase32.Encode(value, 5);
    }

    /// <summary>Generates a CINQ id not already present per <paramref name="exists"/>, retrying on
    /// collision up to <paramref name="maxAttempts"/>. Throws if the (astronomically unlikely) retry
    /// budget is exhausted rather than returning a duplicate id.</summary>
    public static string NewCinqId(Func<string, bool> exists, int maxAttempts = 20)
    {
        for (int i = 0; i < maxAttempts; i++)
        {
            var id = RandomCinqId();
            if (!exists(id)) return id;
        }
        throw new InvalidOperationException("CINQ id generation exhausted its retry budget");
    }

    public enum ThreadAction { CreateNew, Append, Reopen }

    /// <summary>Decides where an inbound message threads. SMS is serial with no deterministic
    /// conversation boundary, so we always append to the customer's latest inquiry — reopening it if it
    /// was Closed. No prior inquiry → start a new one. A Spam thread stays Spam (append, no resurrection);
    /// an operator "split to new thread" action is the explicit way to break a thread.</summary>
    public static ThreadAction DecideThread(Inquiry? latest)
    {
        if (latest is null) return ThreadAction.CreateNew;
        if (string.Equals(latest.Status, InquiryStatus.Closed, StringComparison.OrdinalIgnoreCase))
            return ThreadAction.Reopen;
        return ThreadAction.Append;
    }
}
