using System.Collections.Concurrent;
using System.Security.AccessControl;
using System.Security.Cryptography;
using System.Security.Principal;
using System.Text;

namespace OutlookShredder.Proxy.Services;

/// <summary>
/// WS2 — local REST auth. Locks the loopback proxy API so only our WPF client (same Windows user)
/// can call sensitive endpoints. The proxy mints a random 32-byte token at startup, DPAPI-protects
/// it, and writes it to an ACL'd file only the current user can read; the client signs every request
/// with an HMAC over (method, path, query, timestamp, nonce). Endpoints used by non-WPF callers
/// (Chrome extension, Office add-in) that can't read the token file are exempt.
///
/// Threat model in scope: a *different* local user / non-WPF local process / browser hitting the
/// proxy. Out of scope (trust EDR/OS): same-user secret extraction, local admin/SYSTEM, code injection.
///
/// DENY-BY-DEFAULT: every /api/* requires a valid signature EXCEPT the exempt allowlist.
/// WARN-ONLY first release: when <see cref="Enforce"/> is false the middleware logs would-rejects
/// but serves them; flip Proxy:AuthEnforce=true to hard-enforce (401) once logs are clean.
///
/// The canonical-string + HMAC helpers here are mirrored byte-for-byte by the client's
/// ProxyAuthHandler — a shared test vector (AuthCanonicalTests) pins them together.
/// </summary>
public sealed class ProxyAuthService
{
    public const string TsHeader    = "X-Shredder-Ts";
    public const string NonceHeader = "X-Shredder-Nonce";
    public const string SigHeader   = "X-Shredder-Sig";
    public const string PidHeader   = "X-Caller-PID";

    private readonly ILogger<ProxyAuthService> _log;
    private readonly long _skewMs;
    private readonly int  _rateLimitPerSec;

    public bool Enforce { get; }

    // Set once at startup by EnsureTokenWritten(); held in memory for process life.
    private byte[]? _token;

    // Seen-nonce store: nonce -> expiry (epoch ms). Bounded to ~one skew window of traffic by
    // lazy eviction — any nonce older than the skew window would already fail the timestamp check.
    private readonly ConcurrentDictionary<string, long> _nonces = new();
    private int _nonceOps;

    // Fixed-window per-second request counter (non-exempt requests only).
    private readonly object _rateLock = new();
    private long _rateWindowSec;
    private int  _rateWindowCount;

    // Endpoints reachable by non-WPF callers (Chrome extension / Office add-in / scripts) that cannot
    // read the DPAPI token file. method "*" = any verb. StartsWithSegments matches the prefix + children.
    private static readonly (string Method, string Prefix)[] Exempt =
    {
        ("GET",  "/api/health"),         // Home dashboard + health checks (read-only)
        ("GET",  "/api/ready"),          // startup readiness snapshot (polled before the token race resolves)
        ("GET",  "/api/events"),         // SSE stream (add-in dashboard)
        ("GET",  "/api/items"),          // add-in dashboard data (+ /by-rfq/...)
        ("GET",  "/api/qc"),             // QC read (Python diagnostic script)
        ("POST", "/api/extract"),        // Office.js add-in extraction
        ("POST", "/api/addin/heartbeat"),// add-in heartbeat
        ("*",    "/api/steve"),          // Chrome extension RPA relay
        ("*",    "/api/phone-search"),   // Chrome extension RPA relay
        ("*",    "/api/sms"),            // SignalWire inbound/status webhooks — gated by SW signature instead
    };

    // Dev-only mail-eval labeling/report UI is served by the proxy and called same-origin from a
    // browser, which can't read the DPAPI token to sign requests. Exempt its prefix so the page works.
    // Gated by config (default on) and harmless beyond that: the proxy binds loopback only, and the
    // surface is eval tooling (read golden/items, patch a golden label, run an offline eval).
    private static volatile bool _mailEvalDevUi = true;

    public ProxyAuthService(IConfiguration config, ILogger<ProxyAuthService> log)
    {
        _log = log;
        Enforce          = config.GetValue("Proxy:AuthEnforce", false);
        _skewMs          = config.GetValue("Proxy:AuthSkewSeconds", 30) * 1000L;
        _rateLimitPerSec = config.GetValue("Proxy:RateLimitPerSecond", 300);
        _mailEvalDevUi   = config.GetValue("Proxy:MailEvalDevUi", true);
    }

    // ── Token lifecycle ───────────────────────────────────────────────────────

    /// <summary>
    /// Generate a fresh token, DPAPI-protect it (CurrentUser), and write it to an ACL'd file that only
    /// the current user can access. Rotates every launch. MUST be called synchronously before Kestrel
    /// starts listening so the file exists before the first connection can be accepted.
    /// </summary>
    public void EnsureTokenWritten()
    {
        try
        {
            _token = RandomNumberGenerator.GetBytes(32);
            var dir = Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
                "ShredderData", "Proxy");
            Directory.CreateDirectory(dir);
            var path = Path.Combine(dir, "auth.bin");
            var tmp  = path + ".tmp";

            var protectedBytes = ProtectedData.Protect(_token, optionalEntropy: null,
                                                       scope: DataProtectionScope.CurrentUser);
            File.WriteAllBytes(tmp, protectedBytes);
            ApplyOwnerOnlyAcl(tmp);
            File.Move(tmp, path, overwrite: true);   // atomic replace; preserves ACL on same volume

            _log.LogInformation("[Auth] launch token written ({Bytes}B, enforce={Enforce}) -> {Path}",
                protectedBytes.Length, Enforce, path);
        }
        catch (Exception ex)
        {
            // Non-fatal: with no token the middleware can't validate, so in warn-only mode it just
            // logs would-rejects; in enforce mode every non-exempt request 401s (fail-closed).
            _log.LogError(ex, "[Auth] failed to write launch token — auth will reject all signed requests");
        }
    }

    private static void ApplyOwnerOnlyAcl(string filePath)
    {
        var sid = WindowsIdentity.GetCurrent().User
                  ?? throw new InvalidOperationException("no current-user SID");
        var sec = new FileSecurity();
        sec.SetAccessRuleProtection(isProtected: true, preserveInheritance: false); // drop inherited ACEs
        sec.SetOwner(sid);
        sec.AddAccessRule(new FileSystemAccessRule(sid, FileSystemRights.FullControl, AccessControlType.Allow));
        new FileInfo(filePath).SetAccessControl(sec);
    }

    // ── Exemption + rate limiting ──────────────────────────────────────────────

    public static bool IsExempt(HttpContext ctx)
    {
        if (HttpMethods.IsOptions(ctx.Request.Method)) return true;   // CORS preflight
        var path   = ctx.Request.Path;
        var method = ctx.Request.Method;
        if (_mailEvalDevUi && path.StartsWithSegments("/api/mail-eval", StringComparison.OrdinalIgnoreCase))
            return true;   // dev-only labeling/report UI (browser can't sign)
        foreach (var (m, prefix) in Exempt)
            if ((m == "*" || string.Equals(m, method, StringComparison.OrdinalIgnoreCase))
                && path.StartsWithSegments(prefix, StringComparison.OrdinalIgnoreCase))
                return true;
        return false;
    }

    /// <summary>True if the request is within the per-second rate cap; false if it should be throttled.</summary>
    public bool CheckRate()
    {
        var sec = NowMs() / 1000;
        lock (_rateLock)
        {
            if (sec != _rateWindowSec) { _rateWindowSec = sec; _rateWindowCount = 0; }
            _rateWindowCount++;
            return _rateWindowCount <= _rateLimitPerSec;
        }
    }

    // ── Verification ───────────────────────────────────────────────────────────

    public enum AuthResult { Ok, MissingToken, MissingHeaders, Skew, BadSig, Replay }

    /// <summary>Validate the signature headers on a non-exempt request.</summary>
    public AuthResult Verify(HttpContext ctx)
    {
        var req = ctx.Request;
        return VerifyCore(req.Method, req.Path.Value ?? "", req.QueryString.Value,
            req.Headers[TsHeader].ToString(), req.Headers[NonceHeader].ToString(), req.Headers[SigHeader].ToString());
    }

    /// <summary>Transport-agnostic core (no HttpContext) — unit-testable with plain strings.</summary>
    internal AuthResult VerifyCore(string method, string path, string? rawQuery, string? ts, string? nonce, string? sig)
    {
        if (_token is null) return AuthResult.MissingToken;
        if (string.IsNullOrEmpty(ts) || string.IsNullOrEmpty(nonce) || string.IsNullOrEmpty(sig))
            return AuthResult.MissingHeaders;
        if (!long.TryParse(ts, out var tsMs)) return AuthResult.MissingHeaders;

        var now = NowMs();
        if (Math.Abs(now - tsMs) > _skewMs) return AuthResult.Skew;

        var canonical = BuildCanonical(method, path, rawQuery, ts, nonce);
        var expected  = ComputeSig(_token, canonical);
        byte[] given;
        try { given = Convert.FromBase64String(sig); } catch { return AuthResult.BadSig; }
        if (!CryptographicOperations.FixedTimeEquals(given, expected)) return AuthResult.BadSig;

        // Signature is valid — only now record the nonce (avoids unauthenticated nonce-store flooding).
        if (Interlocked.Increment(ref _nonceOps) % 256 == 0) EvictExpiredNonces(now);
        if (_nonces.TryGetValue(nonce, out var exp) && exp > now) return AuthResult.Replay;
        _nonces[nonce] = tsMs + _skewMs;
        return AuthResult.Ok;
    }

    /// <summary>Test hook: inject a known token without touching disk/DPAPI.</summary>
    internal void SetTokenForTests(byte[] token) => _token = token;

    private void EvictExpiredNonces(long now)
    {
        foreach (var kvp in _nonces)
            if (kvp.Value <= now) _nonces.TryRemove(kvp.Key, out _);
    }

    // ── Shared crypto helpers (mirrored by the client ProxyAuthHandler) ────────

    /// <summary>Canonical string to sign: METHOD\nPATH\nCANONICAL_QUERY\nTS\nNONCE.</summary>
    public static string BuildCanonical(string method, string path, string? rawQuery, string ts, string nonce)
        => string.Join('\n', method.ToUpperInvariant(), path, CanonicalQuery(rawQuery), ts, nonce);

    /// <summary>
    /// Normalize a query string deterministically: strip the leading '?', split on '&', sort the raw
    /// segments ordinally, rejoin. Operates on the raw (un-decoded) text so both sides agree without
    /// any encode/decode round-trip. Empty/absent query → "".
    /// </summary>
    public static string CanonicalQuery(string? rawQuery)
    {
        if (string.IsNullOrEmpty(rawQuery)) return "";
        var q = rawQuery[0] == '?' ? rawQuery[1..] : rawQuery;
        if (q.Length == 0) return "";
        var parts = q.Split('&', StringSplitOptions.RemoveEmptyEntries);
        Array.Sort(parts, StringComparer.Ordinal);
        return string.Join('&', parts);
    }

    public static byte[] ComputeSig(byte[] token, string canonical)
        => HMACSHA256.HashData(token, Encoding.UTF8.GetBytes(canonical));

    private static long NowMs() => DateTimeOffset.UtcNow.ToUnixTimeMilliseconds();
}
