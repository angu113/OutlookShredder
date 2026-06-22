using Microsoft.Extensions.Configuration;
using OutlookShredder.Proxy.Services;

namespace OutlookShredder.Proxy.Services;

/// <summary>
/// In-memory cache for CRM contacts and business partners.
/// Eliminates repeated SharePoint Graph calls for phone-number lookup during Zoom calls
/// and for the contact-mapping dialog in Shredder.
/// Refreshes every 5 minutes; invalidates on write operations.
/// </summary>
public class CustomerCacheService : IHostedService, ICacheStatusProvider
{
    private readonly SharePointService                _sp;
    private readonly ILogger<CustomerCacheService>    _log;
    private readonly IConfiguration                   _config;

    // Atomically-replaced read snapshot — readers never block writers
    private volatile CrmSnapshot _snap = CrmSnapshot.Empty;

    private Timer? _timer;
    private DateTime? _builtUtc;     // set after the first successful refresh — drives readiness
    private DateTime? _lastRefresh;

    public CustomerCacheService(SharePointService sp, ILogger<CustomerCacheService> log, IConfiguration config)
    {
        _sp     = sp;
        _log    = log;
        _config = config;
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

    /// <summary>Customer name → its primary contact (name + phone), one per customer. Used to stamp a
    /// contact onto worklist/delivery cards when the slip doesn't carry the order's own contact.</summary>
    public Dictionary<string, CustomerPrimaryContact> GetPrimaryContacts() => _snap.PrimaryContacts;

    /// <summary>Customer name → ERP payment terms (e.g. "Net 30 Days"). Covers all customers, incl. inactive,
    /// so terms resolve for any document/statement. Empty string/blank when the customer has no terms set.</summary>
    public Dictionary<string, string?> GetAllTerms()     => _snap.BpTerms;

    public CustomerLookupResult? LookupByPhone(string rawPhone) =>
        LookupAllByPhone(rawPhone).FirstOrDefault();

    /// <summary>
    /// Returns all CRM matches for the phone number, one per distinct business partner.
    /// A phone that appears at multiple companies (e.g. a contact who works at two firms)
    /// produces multiple entries rather than silently dropping all but the first.
    /// </summary>
    public List<CustomerLookupResult> LookupAllByPhone(string rawPhone)
    {
        var phone = CustomerImportService.NormalizePhone(rawPhone);
        if (phone is null) return [];

        var snap = _snap;
        if (!snap.PhoneIndex.TryGetValue(phone, out var contacts) || contacts.Count == 0)
            return [];

        return contacts
            .GroupBy(c => c.CustomerName, StringComparer.OrdinalIgnoreCase)
            .Select(g =>
            {
                snap.BpPopups.TryGetValue(g.Key, out var popup);
                var distinctContacts = g.DistinctBy(c => c.ContactName, StringComparer.OrdinalIgnoreCase).ToList();
                var contactName = distinctContacts.Count == 1 ? distinctContacts[0].ContactName : null;
                return new CustomerLookupResult(g.Key, popup, contactName);
            })
            .ToList();
    }

    /// <summary>Call after any write to CRM data so the cache stays fresh.</summary>
    public void Invalidate() => _ = SafeRefreshAsync();

    /// <summary>Awaitable refresh — use when a caller must read CRM data that reflects a just-completed write
    /// (e.g. a targeted backfill right after a phone is linked to a contact).</summary>
    public Task RefreshNowAsync(CancellationToken ct = default) => RefreshAsync(ct);

    // ── ICacheStatusProvider (readiness) ───────────────────────────────────────
    public string Name => "customers";
    public string DisplayName => "Customers / CRM";
    public int SchemaVersion => 1;
    public int ItemCount => _snap.PartnerNames.Count + _snap.Contacts.Count;
    public DateTime? CacheBuiltUtc => _builtUtc;
    public DateTime? LastDeltaUtc => _lastRefresh;
    public bool IsLoading => false;
    public Task ForceRebuildAsync(CancellationToken ct = default) => RefreshAsync(ct);
    public Task ForceDeltaAsync(CancellationToken ct = default) => RefreshAsync(ct);

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

        // Honor the ERP Active (logical-delete) flag: by default, inactive customers — AND their contacts
        // (a contact inherits its customer's active status; there's no per-contact flag) — are excluded
        // from every lookup, keeping worklists/pickers clutter-free. Flip Customers:IncludeInactive=true to
        // surface them. Filtering at snapshot-build means all consumers (names, phone lookup, contacts) are
        // filtered uniformly. The import path reads SP directly, so it still sees/marks inactive records.
        var includeInactive = _config.GetValue("Customers:IncludeInactive", false);
        var activePartners  = includeInactive ? partners : partners.Where(p => p.Active).ToList();

        // Hide a contact ONLY when its parent customer is known AND inactive. Orphan contacts — whose
        // customer isn't in the Customers list at all — are left visible (absence ≠ inactivity), preserving
        // the pre-existing additive behavior for them.
        var inactiveNames = includeInactive
            ? new HashSet<string>(StringComparer.OrdinalIgnoreCase)
            : new HashSet<string>(partners.Where(p => !p.Active).Select(p => p.Name), StringComparer.OrdinalIgnoreCase);

        var visibleContacts = inactiveNames.Count == 0
            ? contacts
            : contacts.Where(c => !inactiveNames.Contains(c.CustomerName)).ToList();

        var phoneIndex = visibleContacts
            .GroupBy(c => CustomerImportService.NormalizePhone(c.Phone) ?? c.Phone,
                     StringComparer.OrdinalIgnoreCase)
            .ToDictionary(g => g.Key, g => g.ToList(), StringComparer.OrdinalIgnoreCase);

        var bpPopups = activePartners.ToDictionary(
            p => p.Name, p => p.PopupMessage, StringComparer.OrdinalIgnoreCase);

        var partnerNames = activePartners
            .Select(p => p.Name)
            .OrderBy(n => n, StringComparer.OrdinalIgnoreCase)
            .ToList();

        // Payment terms keyed by customer name — built from ALL customers (incl. inactive) so terms still
        // resolve for any document/statement that references a customer regardless of active state.
        var bpTerms = new Dictionary<string, string?>(StringComparer.OrdinalIgnoreCase);
        foreach (var p in partners) bpTerms[p.Name] = p.PaymentTerms;

        // One "primary" contact per customer (name + phone), derived from the visible CustomerContacts rows:
        // prefer the first row that has BOTH a contact name and a phone, else the first with a phone, else the
        // first row. Used to stamp a customer's contact onto worklist/delivery cards (CRM fallback when the
        // slip itself doesn't carry the order's contact). Keyed by customer name, case-insensitive.
        var primaryContacts = visibleContacts
            .GroupBy(c => c.CustomerName, StringComparer.OrdinalIgnoreCase)
            .ToDictionary(
                g => g.Key,
                g =>
                {
                    var best = g.FirstOrDefault(c => !string.IsNullOrWhiteSpace(c.ContactName) && !string.IsNullOrWhiteSpace(c.Phone))
                            ?? g.FirstOrDefault(c => !string.IsNullOrWhiteSpace(c.Phone))
                            ?? g.First();
                    return new CustomerPrimaryContact(best.ContactName, best.Phone);
                },
                StringComparer.OrdinalIgnoreCase);

        _snap = new CrmSnapshot(visibleContacts, partnerNames, phoneIndex, bpPopups, bpTerms, primaryContacts);
        _lastRefresh = DateTime.UtcNow;
        _builtUtc ??= _lastRefresh;
        var hiddenBp = partners.Count - activePartners.Count;
        _log.LogInformation(
            "[CRM] Cache refreshed — {C} contacts, {B} business partners{Hidden}",
            visibleContacts.Count, partnerNames.Count,
            (hiddenBp > 0 || contacts.Count != visibleContacts.Count)
                ? $" (hid {hiddenBp} inactive customers, {contacts.Count - visibleContacts.Count} of their contacts)"
                : "");
    }

    // ── Snapshot ──────────────────────────────────────────────────────────────

    private sealed class CrmSnapshot(
        List<CustomerContactRow>                          contacts,
        List<string>                                      partnerNames,
        Dictionary<string, List<CustomerContactRow>>     phoneIndex,
        Dictionary<string, string?>                      bpPopups,
        Dictionary<string, string?>                      bpTerms,
        Dictionary<string, CustomerPrimaryContact>       primaryContacts)
    {
        public List<CustomerContactRow>                      Contacts        { get; } = contacts;
        public List<string>                                  PartnerNames    { get; } = partnerNames;
        public Dictionary<string, List<CustomerContactRow>> PhoneIndex      { get; } = phoneIndex;
        public Dictionary<string, string?>                  BpPopups        { get; } = bpPopups;
        public Dictionary<string, string?>                  BpTerms         { get; } = bpTerms;
        public Dictionary<string, CustomerPrimaryContact>   PrimaryContacts { get; } = primaryContacts;

        public static CrmSnapshot Empty { get; } = new([], [], [], [], [], []);
    }
}

public sealed record CustomerBpRow(string Name, string? PopupMessage, bool Active = true, string? PaymentTerms = null);

/// <summary>A customer's primary contact for card display: contact name + phone (10-digit, unformatted).</summary>
public sealed record CustomerPrimaryContact(string ContactName, string Phone);
