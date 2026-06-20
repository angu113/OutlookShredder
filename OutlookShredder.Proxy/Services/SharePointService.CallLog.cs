using OutlookShredder.Proxy.Models;

namespace OutlookShredder.Proxy.Services;

// Call-log helpers used by the CRM backfill (PhoneController). Kept in their own partial file so the
// 600KB SharePointService.cs doesn't grow. Reuses GetOrCreateCallLogListIdAsync + GetGraph()/site id.
public partial class SharePointService
{
    /// <summary>Reads EVERY PhoneCallLog row (paginated), unlike <see cref="ReadPhoneCallLogAsync"/> which
    /// returns only a recent window. Used by the CRM backfill, which must consider all historic rows.</summary>
    public async Task<List<PhoneCallLogRecord>> ReadAllPhoneCallLogAsync(CancellationToken ct = default)
    {
        var siteId = await GetSiteIdAsync();
        var listId = await GetOrCreateCallLogListIdAsync(ct);

        var results = new List<PhoneCallLogRecord>();
        var page = await GetGraph().Sites[siteId].Lists[listId].Items
            .GetAsync(r =>
            {
                r.QueryParameters.Expand = ["fields($select=Title,CallerPhone,BpName,ContactName,PopupMessage,ReceivedAt,Notes)"];
                r.QueryParameters.Top    = 999;
            }, ct);

        while (page?.Value is not null)
        {
            foreach (var item in page.Value)
            {
                var d = item.Fields?.AdditionalData;
                if (d is null) continue;
                string? Get(string k) => d.TryGetValue(k, out var v) ? v?.ToString() : null;
                results.Add(new PhoneCallLogRecord
                {
                    SpItemId     = item.Id ?? "",
                    CallerName   = Get("Title") ?? "",
                    CallerPhone  = Get("CallerPhone"),
                    BpName       = Get("BpName"),
                    ContactName  = Get("ContactName"),
                    PopupMessage = Get("PopupMessage"),
                    ReceivedAt   = NormalizeInstant(Get("ReceivedAt")),
                    Notes        = Get("Notes"),
                });
            }
            if (page.OdataNextLink is null) break;
            page = await new Microsoft.Graph.Sites.Item.Lists.Item.Items.ItemsRequestBuilder(
                page.OdataNextLink, GetGraph().RequestAdapter).GetAsync(cancellationToken: ct);
        }
        return results;
    }

    /// <summary>Patches BpName/ContactName/PopupMessage on a single PhoneCallLog row by its item ID
    /// (the backfill already holds every row's id, so it patches directly rather than re-querying by phone).</summary>
    public async Task UpdateCallLogCrmByItemIdAsync(
        string spItemId, string? bpName, string? contactName, string? popupMessage,
        CancellationToken ct = default)
    {
        var siteId = await GetSiteIdAsync();
        var listId = await GetOrCreateCallLogListIdAsync(ct);
        await GetGraph().Sites[siteId].Lists[listId].Items[spItemId].Fields.PatchAsync(
            new Microsoft.Graph.Models.FieldValueSet
            {
                AdditionalData = new Dictionary<string, object?>
                {
                    ["BpName"]       = bpName,
                    ["ContactName"]  = contactName,
                    ["PopupMessage"] = popupMessage,
                }
            }, cancellationToken: ct);
    }
}
