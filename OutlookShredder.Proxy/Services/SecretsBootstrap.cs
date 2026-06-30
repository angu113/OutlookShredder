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
        "SignalWire:ApiToken",                  // SMS send + inbound-webhook signature validation
        "SignalWire:MmsBlobConnectionString",   // outbound-MMS media egress (Azure Blob; SAS handed to carrier)
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

    /// <summary>Config keys whose vault secret name does NOT follow the ':'→'--' convention. Lets a secret be
    /// named for its purpose in the vault (e.g. the SignalWire token lives as <c>silmaril-sms-channel</c>)
    /// while still overlaying onto the standard config key the code reads (<c>SignalWire:ApiToken</c>).</summary>
    private static readonly Dictionary<string, string> VaultNameOverrides = new(StringComparer.Ordinal)
    {
        ["SignalWire:ApiToken"]                = "silmaril-sms-channel",
        ["SignalWire:MmsBlobConnectionString"] = "silmaril-sms-media",
    };

    public static string ConfigKeyToVaultName(string configKey) =>
        VaultNameOverrides.TryGetValue(configKey, out var name) ? name : configKey.Replace(":", "--");

    public static string VaultNameToConfigKey(string vaultName)
    {
        foreach (var kv in VaultNameOverrides)
            if (string.Equals(kv.Value, vaultName, StringComparison.Ordinal)) return kv.Key;
        return vaultName.Replace("--", ":");
    }

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

        var cred = BuildSilentBrokerCredential(config["KeyVault:TenantId"]);
        var clientOpts = new SecretClientOptions
        {
            Retry = { MaxRetries = 2, NetworkTimeout = TimeSpan.FromSeconds(10) }
        };
        var client = new SecretClient(new Uri(uri), cred, clientOpts);

        // The Windows WAM broker is frequently not warm yet on a fresh per-user logon task — its first
        // silent token call can come back "operation was canceled" even though a retry seconds later
        // succeeds (observed 2026-06-29: cold start fell back to file → empty secrets → hard crash; a
        // manual restart then loaded 8 secrets cleanly). So retry the vault fetch a few times with a short
        // backoff BEFORE falling back to the (now-retired) local file. A genuine "interaction required"
        // (no silent token possible) is NOT retried — it can't self-heal — and falls back immediately.
        const int maxAttempts = 4;
        for (int attempt = 1; attempt <= maxAttempts; attempt++)
        {
            try
            {
                // Bound each attempt so a hung token broker / network can't stall startup.
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
                log.Information("[Secrets] source=keyvault ({Detail}){Attempt}",
                    Detail, attempt > 1 ? $" on attempt {attempt}/{maxAttempts}" : "");
                return;   // success
            }
            catch (AuthenticationRequiredException ex)
            {
                // Silent broker has no usable token (not signed in / no PRT). Retrying silent auth can't help.
                Source = "file";
                Detail = $"vault auth unavailable (silent, no interaction): {ex.Message}";
                log.Warning("[Secrets] source=file ({Detail})", Detail);
                return;
            }
            catch (Exception ex) when (attempt < maxAttempts)
            {
                // Transient (cold WAM broker cancellation, network blip) — back off and retry before fallback.
                var delay = TimeSpan.FromSeconds(2 * attempt);
                log.Warning("[Secrets] Key Vault attempt {Attempt}/{Max} failed ({Err}) — retrying in {Delay}s",
                    attempt, maxAttempts, ex.Message, delay.TotalSeconds);
                Thread.Sleep(delay);
            }
            catch (Exception ex)
            {
                // Every attempt failed — fall back to file (the historical behaviour).
                Source = "file";
                Detail = $"vault unreachable after {maxAttempts} attempts: {ex.Message}";
                log.Warning("[Secrets] source=file ({Detail})", Detail);
                return;
            }
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
