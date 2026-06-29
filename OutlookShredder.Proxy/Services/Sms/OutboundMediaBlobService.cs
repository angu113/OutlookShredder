using Azure.Storage.Blobs;
using Azure.Storage.Blobs.Models;
using Azure.Storage.Sas;

namespace OutlookShredder.Proxy.Services.Sms;

/// <summary>
/// Ephemeral egress for outbound MMS media. SignalWire's REST <c>MediaUrl</c> must be a publicly-fetchable URL
/// (there is no upload-to-carrier path), so we upload the bytes to an Azure Blob container and hand SignalWire a
/// SHORT-LIVED SAS URL it can download. This store is THROWAWAY: the permanent copy of every sent image lives in
/// SharePoint <c>InquiryMedia/</c> (the conversation renders that), the SAS expires in ~1h, and the blobs
/// auto-delete via a lifecycle rule. Configured by the <c>SignalWire:MmsBlobConnectionString</c> secret (Key
/// Vault <c>silmaril-sms-media</c>); when absent, outbound MMS is disabled (<see cref="IsConfigured"/> = false).
/// See memory project_outbound_mms_durability.
/// </summary>
public sealed class OutboundMediaBlobService
{
    private const string   ContainerName = "sms-outbound-media";
    private static readonly TimeSpan SasLifetime = TimeSpan.FromHours(1);

    private readonly string? _connStr;
    private readonly ILogger<OutboundMediaBlobService> _log;
    private BlobContainerClient? _container;

    public OutboundMediaBlobService(IConfiguration config, ILogger<OutboundMediaBlobService> log)
    {
        _connStr = config["SignalWire:MmsBlobConnectionString"];
        _log     = log;
    }

    /// <summary>True when the blob connection string is set — this gates the whole outbound-MMS feature so it
    /// ships dark until the secret exists.</summary>
    public bool IsConfigured => !string.IsNullOrWhiteSpace(_connStr);

    /// <summary>Uploads the bytes under <c>{inquiryId}/{name}</c> and returns a read-only SAS URL valid ~1h for
    /// SignalWire to fetch. Overwrites an existing blob of the same name. Throws when not configured.</summary>
    public async Task<string> UploadAndGetSasUrlAsync(string inquiryId, string name, byte[] bytes, string contentType, CancellationToken ct = default)
    {
        if (!IsConfigured) throw new InvalidOperationException("Outbound MMS media blob is not configured.");

        var container = await GetContainerAsync(ct);
        var blobName  = $"{inquiryId}/{name}";
        var blob      = container.GetBlobClient(blobName);

        using (var ms = new MemoryStream(bytes, writable: false))
        {
            await blob.UploadAsync(ms, new BlobUploadOptions
            {
                HttpHeaders = new BlobHttpHeaders { ContentType = contentType },
            }, ct);
        }

        if (!blob.CanGenerateSasUri)
            throw new InvalidOperationException(
                "Blob client cannot generate a SAS — the connection string must include the account key (AccountKey=...).");

        var sas = new BlobSasBuilder
        {
            BlobContainerName = container.Name,
            BlobName          = blobName,
            Resource          = "b",
            ExpiresOn         = DateTimeOffset.UtcNow.Add(SasLifetime),
        };
        sas.SetPermissions(BlobSasPermissions.Read);
        var url = blob.GenerateSasUri(sas).ToString();
        _log.LogInformation("[MmsBlob] uploaded {Blob} ({Bytes} bytes), SAS valid {Mins}m", blobName, bytes.Length, (int)SasLifetime.TotalMinutes);
        return url;
    }

    private async Task<BlobContainerClient> GetContainerAsync(CancellationToken ct)
    {
        if (_container is not null) return _container;
        var svc       = new BlobServiceClient(_connStr);
        var container = svc.GetBlobContainerClient(ContainerName);
        await container.CreateIfNotExistsAsync(PublicAccessType.None, cancellationToken: ct);   // SAS-only, never anonymous
        return _container = container;
    }
}
