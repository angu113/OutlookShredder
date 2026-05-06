using OutlookShredder.Proxy.Services;

namespace OutlookShredder.Proxy.Services;

/// <summary>
/// In-memory cache for CRM contacts and business partners.
/// Eliminates repeated SharePoint Graph calls for phone-number lookup during Zoom calls
/// and for the contact-mapping dialog in Shredder.
/// Refreshes every 5 minutes; invalidates on write operations.
/// </summary>
public class CustomerCacheService : IHostedService
{
    private readonly SharePointService                _sp;
    private readonly ILogger<CustomerCacheService>    _log;

    // Atomically-replaced read snapshot — readers never block writers
    private volatile CrmSnapshot _snap = CrmSnapshot.Empty;

    private Timer? _timer;

    public CustomerCacheService(SharePointService sp, ILogger<CustomerCacheService> log)
    {
        _sp  = sp;
        _log = log;
    }

    public async Task StartAsync(CancellationToken ct)
    {
        await RefreshAsync(ct);
        _timer = new Timer(_ => _ = SafeRefreshAsync(), null,
            TimeSpan.FromMinutes(5), TimeSpan.FromMinutes(5));
    }

    public Task StopAsync(CancellationToken ct)
    {
        _timer?.Dispose();
        return Task.CompletedTask;
    }

    // ── Public API (all return immediately from the in-memory snapshot) ────────

    public List<CustomerContactRow> GetAllContacts()     => _snap.Contacts;
    public List<string>             GetAllPartnerNames() => _snap.PartnerNames;

    public CustomerLookupResult? LookupByPhone(string rawPhone)
    {
        var phone = CustomerImportService.NormalizePhone(rawPhone);
        if (phone is null) return null;

        var snap = _snap;
        if (!snap.PhoneIndex.TryGetValue(phone, out var contacts) || contacts.Count == 0)
            return null;

        var bpName      = contacts[0].CustomerName;
        var contactName = contacts.Count == 1 ? contacts[0].ContactName : null;
        snap.BpPopups.TryGetValue(bpName, out var popup);
        return new CustomerLookupResult(bpName, popup, contactName);
    }

    /// <summary>Call after any write to CRM data so the cache stays fresh.</summary>
    public void Invalidate() => _ = SafeRefreshAsync();

    // ── Private ───────────────────────────────────────────────────────────────

    private async Task SafeRefreshAsync()
    {
        try { await RefreshAsync(CancellationToken.None); }
        catch (Exception ex) { _log.LogWarning(ex, "[CRM] Cache refresh failed — old data retained"); }
    }

    private async Task RefreshAsync(CancellationToken ct)
    {
        var contactsTask = _sp.ReadAllContactsAsync(ct);
        var partnersTask = _sp.ReadAllCustomersAsync(ct);
        await Task.WhenAll(contactsTask, partnersTask);

        var contacts = await contactsTask;
        var partners = await partnersTask;

        var phoneIndex = contacts
            .GroupBy(c => CustomerImportService.NormalizePhone(c.Phone) ?? c.Phone,
                     StringComparer.OrdinalIgnoreCase)
            .ToDictionary(g => g.Key, g => g.ToList(), StringComparer.OrdinalIgnoreCase);

        var bpPopups = partners.ToDictionary(
            p => p.Name, p => p.PopupMessage, StringComparer.OrdinalIgnoreCase);

        var partnerNames = partners
            .Select(p => p.Name)
            .OrderBy(n => n, StringComparer.OrdinalIgnoreCase)
            .ToList();

        _snap = new CrmSnapshot(contacts, partnerNames, phoneIndex, bpPopups);
        _log.LogInformation("[CRM] Cache refreshed — {C} contacts, {B} business partners",
            contacts.Count, partnerNames.Count);
    }

    // ── Snapshot ──────────────────────────────────────────────────────────────

    private sealed class CrmSnapshot(
        List<CustomerContactRow>                          contacts,
        List<string>                                      partnerNames,
        Dictionary<string, List<CustomerContactRow>>     phoneIndex,
        Dictionary<string, string?>                      bpPopups)
    {
        public List<CustomerContactRow>                      Contacts     { get; } = contacts;
        public List<string>                                  PartnerNames { get; } = partnerNames;
        public Dictionary<string, List<CustomerContactRow>> PhoneIndex   { get; } = phoneIndex;
        public Dictionary<string, string?>                  BpPopups     { get; } = bpPopups;

        public static CrmSnapshot Empty { get; } = new([], [], [], []);
    }
}

public sealed record CustomerBpRow(string Name, string? PopupMessage);
