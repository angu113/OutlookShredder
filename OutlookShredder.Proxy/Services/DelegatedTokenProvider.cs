using Microsoft.Identity.Client;

namespace OutlookShredder.Proxy.Services;

/// <summary>
/// Provides access tokens for the hackensack mailbox via MSAL.
///
/// Token acquisition order:
///   1. In-memory cache (skip if expired within 5 min)
///   2. AcquireTokenSilent  — uses the MSAL cache file, no UI
///   3. AcquireTokenWithDeviceCode — prints URL + code to the log; user visits the URL
///      in any browser and enters the code to complete sign-in (supports MFA).
///      Once authenticated the token is cached to disk and all future calls are silent.
///
/// Config keys (appsettings.json / secrets):
///   HackensackMail:ClientId       — optional override (default: Outlook Mobile app)
///   HackensackMail:TokenCachePath — optional override; default: {exeDir}/hackensack-msal-cache.bin
/// </summary>
public class DelegatedTokenProvider
{
    private const string DefaultClientId = "27922004-5251-4030-b22d-91ecd9a37ea4"; // Outlook Mobile (pre-consented for all Exchange Online tenants)

    private static readonly string[] Scopes =
    [
        "https://outlook.office.com/IMAP.AccessAsUser.All",
        "offline_access",
    ];

    private readonly IPublicClientApplication _msal;
    private readonly string _cachePath;
    private readonly ILogger<DelegatedTokenProvider> _log;
    private readonly SemaphoreSlim _lock = new(1, 1);

    private string _accessToken = "";
    private DateTime _expiresAt = DateTime.MinValue;

    public bool IsConfigured => File.Exists(_cachePath);

    public DelegatedTokenProvider(IConfiguration config, ILogger<DelegatedTokenProvider> log)
    {
        _log = log;

        var clientId = config["HackensackMail:ClientId"] ?? DefaultClientId;
        var appDir   = Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().Location) ?? ".";
        _cachePath   = config["HackensackMail:TokenCachePath"]
                    ?? Path.Combine(appDir, "hackensack-msal-cache.bin");

        _msal = PublicClientApplicationBuilder
            .Create(clientId)
            .WithDefaultRedirectUri()
            .Build();

        RegisterFileCache(_msal.UserTokenCache, _cachePath);

        if (File.Exists(_cachePath))
            _log.LogInformation("[HackensackMail] MSAL cache loaded from {Path}", _cachePath);
        else
            _log.LogInformation("[HackensackMail] No MSAL cache at {Path} — will prompt on first poll", _cachePath);
    }

    public async Task<string> GetAccessTokenAsync(CancellationToken ct = default)
    {
        await _lock.WaitAsync(ct);
        try
        {
            if (!string.IsNullOrEmpty(_accessToken) && DateTime.UtcNow < _expiresAt.AddMinutes(-5))
                return _accessToken;

            var result = await AcquireAsync(ct);
            _accessToken = result.AccessToken;
            _expiresAt   = result.ExpiresOn.UtcDateTime;

            _log.LogInformation("[HackensackMail] Token acquired for {Account}, expires {At:HH:mm:ss} UTC",
                result.Account.Username, _expiresAt);

            return _accessToken;
        }
        finally { _lock.Release(); }
    }

    private async Task<AuthenticationResult> AcquireAsync(CancellationToken ct)
    {
        var accounts = (await _msal.GetAccountsAsync()).ToList();

        // --- Silent path ---
        if (accounts.Count > 0)
        {
            try
            {
                return await _msal
                    .AcquireTokenSilent(Scopes, accounts[0])
                    .ExecuteAsync(ct);
            }
            catch (MsalUiRequiredException ex)
            {
                _log.LogWarning("[HackensackMail] Silent token refresh requires interaction ({Code}): {Msg}",
                    ex.ErrorCode, ex.Message);
            }
        }
        else
        {
            _log.LogInformation("[HackensackMail] No cached account — starting device code flow");
        }

        // --- Device code flow ---
        // Works from any session. User visits the printed URL in any browser and enters the code.
        // Supports MFA. Token is cached to disk on success — subsequent startups are silent.
        return await _msal
            .AcquireTokenWithDeviceCode(Scopes, deviceCode =>
            {
                _log.LogWarning(
                    "[HackensackMail] ACTION REQUIRED — sign in to the Hackensack mailbox:\n" +
                    "  1. Open this URL in any browser: {Url}\n" +
                    "  2. Enter code: {Code}\n" +
                    "  3. Sign in as hackensack@metalsupermarkets.com (MFA will work normally)\n" +
                    "  Code expires in ~15 minutes. This is a one-time step.",
                    deviceCode.VerificationUrl, deviceCode.UserCode);
                return Task.CompletedTask;
            })
            .ExecuteAsync(ct);
    }

    private static void RegisterFileCache(ITokenCache cache, string path)
    {
        cache.SetBeforeAccess(args =>
        {
            if (File.Exists(path))
            {
                try { args.TokenCache.DeserializeMsalV3(File.ReadAllBytes(path)); }
                catch { /* corrupt cache — start fresh */ }
            }
        });

        cache.SetAfterAccess(args =>
        {
            if (args.HasStateChanged)
                File.WriteAllBytes(path, args.TokenCache.SerializeMsalV3());
        });
    }
}
