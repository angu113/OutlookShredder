using System.IO;
using Azure;
using Azure.Core;
using Azure.Identity;
using Azure.Identity.Broker;
using Azure.Security.KeyVault.Secrets;
using Microsoft.Extensions.Configuration;

namespace OutlookShredder.Proxy.Services;

/// <summary>
/// WS1 — fetch secrets from Azure Key Vault at startup using the signed-in user's Entra identity
/// (Windows WAM broker, silent), overlaying them into IConfiguration so existing consumers
/// (`_config["Section:Key"]`) pick them up with no code changes. Replaces cleartext-on-OneDrive.
///
/// Credential: the proxy runs as a per-user scheduled task inside the user's interactive logon
/// session, so WAM already holds the user's Primary Refresh Token (machines are Entra-joined).
/// We request the token SILENTLY (UseDefaultBrokerAccount + DisableAutomaticAuthentication) — no
/// prompt, no certificate. If the vault is unreachable (offline / no RBAC / non-joined), we log
/// and fall back to the existing local secrets file (vault-first, file fallback).
///
/// Vault secret names mirror config keys with '--' for ':' (e.g. SharePoint--ClientSecret ->
/// SharePoint:ClientSecret), so no custom name resolver is needed.
/// </summary>
public static class SecretsBootstrap
{
    /// <summary>The secrets WS1 migrates (config-key form). Vault names are the '--' variant.</summary>
    public static readonly string[] ConfigKeys =
    {
        "Anthropic:ApiKey",
        "Google:ApiKey",
        "SharePoint:TenantId",
        "SharePoint:ClientId",
        "SharePoint:ClientSecret",
        "ServiceBus:ConnectionString",
    };

    /// <summary>"keyvault" or "file" — surfaced by HealthController.</summary>
    public static string Source { get; private set; } = "file";
    /// <summary>Human-readable detail (count / reason) — surfaced by HealthController.</summary>
    public static string Detail { get; private set; } = "Key Vault not attempted";
    /// <summary>True when the vault supplied EVERY expected secret (no key fell back to the local file) —
    /// the precondition for deleting the local cleartext copies.</summary>
    public static bool Complete { get; private set; }

    // Local cleartext secret files this overlay supersedes — deleted once the vault is proven healthy.
    private static readonly List<string> _localSecretFiles = new();
    /// <summary>Registered from Program.cs for each local appsettings.secrets.json config source so the
    /// post-startup cleanup knows which files to remove once the vault is healthy.</summary>
    public static void RegisterLocalSecretFile(string path)
    {
        if (!string.IsNullOrWhiteSpace(path) && !_localSecretFiles.Contains(path)) _localSecretFiles.Add(path);
    }

    public static string ConfigKeyToVaultName(string configKey) => configKey.Replace(":", "--");
    public static string VaultNameToConfigKey(string vaultName) => vaultName.Replace("--", ":");

    /// <summary>
    /// Try to overlay vault secrets onto <paramref name="config"/>. Must run AFTER the JSON file
    /// sources and BEFORE the host is built, so the overlay wins and DI-time readers see vault values.
    /// Never throws — on any failure it leaves the file values in place.
    /// </summary>
    public static void LoadInto(ConfigurationManager config, Serilog.ILogger log)
    {
        var uri = config["KeyVault:Uri"];
        if (string.IsNullOrWhiteSpace(uri))
        {
            Source = "file";
            Detail = "KeyVault:Uri not configured";
            log.Information("[Secrets] source=file ({Detail})", Detail);
            return;
        }

        try
        {
            var cred = BuildSilentBrokerCredential(config["KeyVault:TenantId"]);
            var clientOpts = new SecretClientOptions
            {
                Retry = { MaxRetries = 2, NetworkTimeout = TimeSpan.FromSeconds(10) }
            };
            var client = new SecretClient(new Uri(uri), cred, clientOpts);

            // Bound the whole fetch so a hung token broker / network can't stall startup.
            using var cts = new CancellationTokenSource(TimeSpan.FromSeconds(30));

            var overlay = new Dictionary<string, string?>();
            var missing = new List<string>();
            foreach (var key in ConfigKeys)
            {
                var name = ConfigKeyToVaultName(key);
                try
                {
                    var secret = client.GetSecret(name, cancellationToken: cts.Token);
                    overlay[key] = secret.Value.Value;
                }
                catch (RequestFailedException ex) when (ex.Status == 404)
                {
                    missing.Add(name);   // not in the vault yet — keep the file value for this key
                }
            }

            if (overlay.Count == 0)
                throw new InvalidOperationException("vault reachable but returned none of the expected secrets");

            config.AddInMemoryCollection(overlay);   // added last -> overrides the file sources
            Source = "keyvault";
            Complete = missing.Count == 0;
            Detail = missing.Count == 0
                ? $"{overlay.Count} secrets"
                : $"{overlay.Count} secrets, {missing.Count} missing ({string.Join(", ", missing)}) fell back to file";
            log.Information("[Secrets] source=keyvault ({Detail})", Detail);
        }
        catch (Exception ex)
        {
            Source = "file";
            Detail = $"vault unreachable: {ex.Message}";
            log.Warning("[Secrets] source=file ({Detail})", Detail);
        }
    }

    /// <summary>
    /// WS1 cleanup: once the vault has supplied EVERY secret (<see cref="Complete"/>) AND the proxy has
    /// proven those credentials work at runtime (the caller invokes this only after a successful SharePoint
    /// prewarm), delete the local cleartext secret files registered via <see cref="RegisterLocalSecretFile"/>.
    /// The central OneDrive copy is intentionally left in place for now. Never throws.
    /// </summary>
    public static void CleanupLocalSecretsIfVaultHealthy(Serilog.ILogger log)
    {
        if (Source != "keyvault")
        {
            log.Information("[Secrets] cleanup skipped — source={Source}", Source);
            return;
        }
        if (!Complete)
        {
            log.Information("[Secrets] cleanup skipped — vault incomplete ({Detail}); add the missing secrets to the vault to enable cleanup", Detail);
            return;
        }
        foreach (var path in _localSecretFiles)
        {
            try
            {
                if (File.Exists(path))
                {
                    File.Delete(path);
                    log.Information("[Secrets] deleted local cleartext copy (vault healthy): {Path}", path);
                }
            }
            catch (Exception ex)
            {
                log.Warning("[Secrets] could not delete local copy {Path}: {Err}", path, ex.Message);
            }
        }
    }

    /// <summary>
    /// Signed-in user's token via the Windows WAM broker, SILENT only: UseDefaultBrokerAccount uses the
    /// logged-on Entra account; DisableAutomaticAuthentication makes it throw (caught -> fallback) rather
    /// than ever popping UI from this headless task.
    /// </summary>
    private static TokenCredential BuildSilentBrokerCredential(string? tenantId)
    {
        var opts = new InteractiveBrowserCredentialBrokerOptions(IntPtr.Zero)
        {
            UseDefaultBrokerAccount        = true,
            DisableAutomaticAuthentication = true,
        };
        if (!string.IsNullOrWhiteSpace(tenantId)) opts.TenantId = tenantId;
        return new InteractiveBrowserCredential(opts);
    }
}
