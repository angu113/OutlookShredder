using Microsoft.Extensions.Configuration;
using OutlookShredder.Proxy.Services;
using Xunit;

namespace OutlookShredder.Proxy.Tests;

// WS1 secrets → Key Vault. Pins the vault-name <-> config-key mapping (so the stock '--'→':' overlay
// fills IConfiguration with no custom resolver) and the precedence contract the bootstrap relies on:
// a source added AFTER the file sources overrides them (vault-first, file fallback).
public class SecretsBootstrapTests
{
    [Theory]
    [InlineData("Anthropic:ApiKey",            "Anthropic--ApiKey")]
    [InlineData("Google:ApiKey",               "Google--ApiKey")]
    [InlineData("SharePoint:TenantId",         "SharePoint--TenantId")]
    [InlineData("SharePoint:ClientId",         "SharePoint--ClientId")]
    [InlineData("SharePoint:ClientSecret",     "SharePoint--ClientSecret")]
    [InlineData("ServiceBus:ConnectionString", "ServiceBus--ConnectionString")]
    public void ConfigKey_and_VaultName_round_trip(string configKey, string vaultName)
    {
        Assert.Equal(vaultName, SecretsBootstrap.ConfigKeyToVaultName(configKey));
        Assert.Equal(configKey, SecretsBootstrap.VaultNameToConfigKey(vaultName));
    }

    [Fact]
    public void Every_migrated_key_maps_cleanly_both_ways()
    {
        foreach (var key in SecretsBootstrap.ConfigKeys)
        {
            var vaultName = SecretsBootstrap.ConfigKeyToVaultName(key);
            Assert.DoesNotContain(":", vaultName);                 // vault names can't contain ':'
            Assert.Equal(key, SecretsBootstrap.VaultNameToConfigKey(vaultName));
        }
    }

    [Fact]
    public void Overlay_added_after_file_source_overrides_it()
    {
        // Contract the bootstrap relies on: the in-memory vault overlay, added LAST, wins over the
        // earlier "file" source for the same key, while keys only present in the file survive.
        var file = new Dictionary<string, string?>
        {
            ["SharePoint:ClientSecret"] = "file-secret",
            ["Mail:MailboxAddress"]     = "store@example.com",   // not migrated -> must persist
        };
        var vaultOverlay = new Dictionary<string, string?>
        {
            ["SharePoint:ClientSecret"] = "vault-secret",
        };

        var config = new ConfigurationBuilder()
            .AddInMemoryCollection(file)            // stands in for the appsettings.secrets.json layer
            .AddInMemoryCollection(vaultOverlay)    // added after -> higher precedence (the vault overlay)
            .Build();

        Assert.Equal("vault-secret", config["SharePoint:ClientSecret"]);   // vault wins
        Assert.Equal("store@example.com", config["Mail:MailboxAddress"]);  // file-only key preserved
    }
}
