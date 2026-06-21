using OutlookShredder.Proxy.Models;

namespace OutlookShredder.Proxy.Services;

/// <summary>
/// Backfills CRM fields (BpName / ContactName / PopupMessage) on existing PhoneCallLog rows using the SAME
/// phone→customer lookup as live incoming-call detection (<see cref="CustomerCacheService.LookupAllByPhone"/>,
/// primary match). Reads every historic row, normalises its CallerPhone, looks it up, and sets the matched
/// CRM values. <c>dryRun=true</c> reports what WOULD change without writing. <c>includeChanged=false</c> fills
/// only blanks (a gain) and never overwrites an existing BpName.
///
/// Extracted from PhoneController so the same scan can run after a customer/contact import
/// (<see cref="Controllers.ImportController"/>) as well as on the manual <c>/api/phone/call-log/backfill-crm</c>
/// endpoint.
/// </summary>
public sealed class CallLogCrmBackfillService(
    SharePointService sp,
    CustomerCacheService crm,
    ILogger<CallLogCrmBackfillService> log)
{
    public enum CrmBackfillAction { Unchanged, AddBp, ChangeBp }

    public sealed record BackfillResult(
        int TotalRecords, int WouldUpdate, int AddBlank, int ChangeExisting, int Unchanged,
        int NoMatch, int NoPhone, int MultiMatchPhones, int Updated, int Failed,
        IReadOnlyList<string> AddSamples, IReadOnlyList<string> ChangeSamples);

    /// <summary>Decides the backfill action for one call-log row given its matched CRM values. Pure +
    /// testable: Unchanged when all three CRM fields already equal the match (null/blank- and trim-
    /// insensitive); AddBp when BpName is currently blank (a gain); otherwise ChangeBp (would overwrite).</summary>
    public static CrmBackfillAction Classify(
        PhoneCallLogRecord rec, string? matchBp, string? matchContact, string? matchPopup)
    {
        static bool Same(string? a, string? b) =>
            string.Equals((a ?? "").Trim(), (b ?? "").Trim(), StringComparison.Ordinal);

        if (Same(rec.BpName, matchBp) && Same(rec.ContactName, matchContact) && Same(rec.PopupMessage, matchPopup))
            return CrmBackfillAction.Unchanged;

        return string.IsNullOrWhiteSpace(rec.BpName) ? CrmBackfillAction.AddBp : CrmBackfillAction.ChangeBp;
    }

    public async Task<BackfillResult> RunAsync(bool dryRun, bool includeChanged, CancellationToken ct = default)
    {
        var records = await sp.ReadAllPhoneCallLogAsync(ct);

        int noPhone = 0, noMatch = 0, unchanged = 0, addBlank = 0, changeExisting = 0, multiMatch = 0;
        var plan          = new List<(PhoneCallLogRecord Rec, CustomerLookupResult Match)>();
        var addSamples    = new List<string>();
        var changeSamples = new List<string>();

        foreach (var rec in records)
        {
            var digits = CustomerImportService.NormalizePhone(rec.CallerPhone);
            if (digits is null) { noPhone++; continue; }

            var matches = crm.LookupAllByPhone(digits);
            if (matches.Count == 0) { noMatch++; continue; }
            if (matches.Count > 1) multiMatch++;
            var m = matches[0];   // primary match — same as live detection's call-log write

            switch (Classify(rec, m.BusinessPartner, m.ContactName, m.PopupMessage))
            {
                case CrmBackfillAction.Unchanged:
                    unchanged++;
                    break;
                case CrmBackfillAction.AddBp:
                    addBlank++;
                    plan.Add((rec, m));
                    if (addSamples.Count < 40)
                        addSamples.Add($"{rec.CallerPhone} -> '{m.BusinessPartner}' (was blank)");
                    break;
                case CrmBackfillAction.ChangeBp:
                    changeExisting++;
                    if (includeChanged) plan.Add((rec, m));
                    if (changeSamples.Count < 40)
                        changeSamples.Add($"{rec.CallerPhone}: '{rec.BpName}' -> '{m.BusinessPartner}'");
                    break;
            }
        }

        int updated = 0, failed = 0;
        if (!dryRun)
        {
            foreach (var (rec, m) in plan)
            {
                try
                {
                    await sp.UpdateCallLogCrmByItemIdAsync(
                        rec.SpItemId, m.BusinessPartner, m.ContactName, m.PopupMessage, ct);
                    updated++;
                }
                catch (Exception ex)
                {
                    failed++;
                    log.LogWarning(ex, "[Phone] backfill patch failed for item {Id}", rec.SpItemId);
                }
            }
            log.LogInformation("[Phone] CRM backfill: {Updated} updated, {Failed} failed (of {Plan} planned)",
                updated, failed, plan.Count);
        }

        return new BackfillResult(
            records.Count, plan.Count, addBlank, changeExisting, unchanged,
            noMatch, noPhone, multiMatch, updated, failed, addSamples, changeSamples);
    }
}
