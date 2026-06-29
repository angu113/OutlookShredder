using System.IO;
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

    // ── Inbound media (MMS now; email attachments later — channel-agnostic) ──────────────────────────
    public enum MediaKind { Ignore, Caption, Image, Pdf, Cad, File }

    private static readonly HashSet<string> CadExts = new(StringComparer.OrdinalIgnoreCase)
    { "dxf", "dwg", "dwf", "dgn", "step", "stp", "iges", "igs", "stl", "sat", "x_t", "x_b", "3dm", "ipt", "sldprt", "catpart", "prt" };

    /// <summary>Classifies one inbound media part by content-type (+ optional name). An MMS delivers the
    /// message text as a <c>text/plain</c> part (promoted to the body) and a SMIL layout as
    /// <c>application/smil</c>/<c>text/html</c> (dropped). Binary parts route to inline preview (image/pdf),
    /// a "Save to CAD" download (drawing files), or a generic "Download" (everything else).</summary>
    public static MediaKind ClassifyMedia(string? contentType, string? name)
    {
        var ct  = (contentType ?? "").ToLowerInvariant();
        var ext = (Path.GetExtension(name ?? "")).TrimStart('.').ToLowerInvariant();
        if (ct.Contains("smil") || ct == "text/html")            return MediaKind.Ignore;
        if (ct.StartsWith("text/plain"))                          return MediaKind.Caption;
        if (ct.StartsWith("image/") && !ct.Contains("dxf") && !ct.Contains("dwg"))
        {
            // image/vnd.dxf / image/vnd.dwg masquerade as images but are CAD drawings.
            return CadExts.Contains(ext) ? MediaKind.Cad : MediaKind.Image;
        }
        if (ct.Contains("pdf") || ext == "pdf")                   return MediaKind.Pdf;
        if (IsCad(ct, ext))                                       return MediaKind.Cad;
        return MediaKind.File;
    }

    private static bool IsCad(string ct, string ext) =>
        CadExts.Contains(ext) || ct.Contains("dxf") || ct.Contains("dwg") || ct.Contains("autocad")
        || ct.Contains("acad") || ct.Contains("step") || ct.Contains("iges") || ct.Contains("x-step");

    /// <summary>A storage file extension for a media part — prefers any extension on the name, else maps the
    /// content type (MMS parts carry no filename, so the content type is usually all we have).</summary>
    public static string ExtForContentType(string? contentType, string? name = null)
    {
        var ext = (Path.GetExtension(name ?? "")).TrimStart('.').ToLowerInvariant();
        if (ext.Length > 0) return ext;
        var ct = (contentType ?? "").ToLowerInvariant();
        return ct switch
        {
            "image/jpeg" or "image/jpg" => "jpg",
            "image/png"  => "png",
            "image/gif"  => "gif",
            "image/webp" => "webp",
            "image/heic" => "heic",
            "application/pdf" => "pdf",
            _ when ct.Contains("dxf")      => "dxf",
            _ when ct.Contains("dwg")      => "dwg",
            _ when ct.StartsWith("image/") => ct[6..],
            _ => "bin",
        };
    }

    /// <summary>The content type to serve a stored media file back with, derived from its extension.</summary>
    public static string MimeForName(string name)
    {
        var ext = (Path.GetExtension(name)).TrimStart('.').ToLowerInvariant();
        return ext switch
        {
            "jpg" or "jpeg" => "image/jpeg",
            "png"  => "image/png",
            "gif"  => "image/gif",
            "webp" => "image/webp",
            "heic" => "image/heic",
            "pdf"  => "application/pdf",
            "dxf"  => "image/vnd.dxf",
            "dwg"  => "image/vnd.dwg",
            _      => "application/octet-stream",
        };
    }

    /// <summary>Guards a media file name used as a drive path segment — opaque keys we mint
    /// (<c>{sid}-{i}.{ext}</c>) only; rejects traversal / separators.</summary>
    public static bool IsSafeMediaName(string? name) =>
        !string.IsNullOrWhiteSpace(name) && name.IndexOfAny(['/', '\\']) < 0 && !name.Contains("..")
        && name.Length <= 200;

    /// <summary>Light shape inference for catalog matching (Phase 6): a square tube/box or an equal angle given
    /// with a SINGLE dimension implies a square/equal cross-section ("2 box" → also 2x2, "2 angle" → also 2x2),
    /// which a literal token parse misses. Appends the inferred square dims so the matcher finds the right
    /// family (only the thickness is then missing). Best-effort; returns the text unchanged when it doesn't apply.</summary>
    public static string ExpandTerseDims(string? text)
    {
        var t = (text ?? "").Trim();
        if (t.Length == 0) return t;
        var lower = t.ToLowerInvariant();
        var squareLike = lower.Contains("box") || lower.Contains("square tube") || lower.Contains("sq tube")
                      || lower.Contains("square tubing") || lower.Contains("angle");
        if (!squareLike) return t;
        if (Regex.IsMatch(t, @"\d\s*[xX]\s*\d")) return t;                                  // already has two dims
        var m = Regex.Match(t, @"\b\d+(?:[-\s]?\d+/\d+|\.\d+|/\d+)?\b");                     // first numeric dimension
        if (!m.Success) return t;
        var dim = m.Value;
        return $"{t} {dim}x{dim}";                                                          // give the tokenizer the square cross-section
    }

    public enum ThreadAction { CreateNew, Append, Reopen }

    /// <summary>Decides where an inbound message threads, given the customer's inquiries (most-recent first).
    /// ONE THREAD PER CUSTOMER: a returning customer always continues their existing conversation, never spawns
    /// a duplicate. We append to the most-recent **active** (non-Closed) inquiry if one exists (Spam counts as
    /// active, so new spam stays in its thread, out of the inbox). If ALL the customer's threads are **Closed**,
    /// we **reopen** the most-recent one (the caller flips it back to Open) rather than minting a new CINQ — this
    /// reverses the old "Closed is terminal → start fresh" behavior. Only a customer with NO prior inquiry gets a
    /// new one. (Multiple concurrent open threads for one customer is a future "split" concern.)</summary>
    public static (ThreadAction Action, Inquiry? Target) DecideThread(IReadOnlyList<Inquiry> existing)
    {
        if (existing.Count == 0) return (ThreadAction.CreateNew, null);
        var active = existing.FirstOrDefault(i =>
            !string.Equals(i.Status, InquiryStatus.Closed, StringComparison.OrdinalIgnoreCase));
        if (active is not null) return (ThreadAction.Append, active);
        return (ThreadAction.Reopen, existing[0]);   // all closed → continue (reopen) the most-recent thread
    }
}
