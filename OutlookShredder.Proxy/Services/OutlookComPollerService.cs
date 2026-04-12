using System.Runtime.InteropServices;
using System.Runtime.Versioning;
using System.Text.RegularExpressions;

namespace OutlookShredder.Proxy.Services;

/// <summary>
/// Background service that polls a local Outlook account for outbound Purchase Order emails
/// by connecting to the running Outlook process via COM automation.
/// Used for mailboxes that are not accessible via Microsoft Graph (e.g. hackensack@metalsupermarkets.com).
///
/// Config keys (appsettings.secrets.json):
///   OutlookCom:Mailbox              — display name or email of the Outlook store to read, e.g. "hackensack@metalsupermarkets.com"
///   OutlookCom:PollIntervalSeconds  — how often to poll (default: 60)
///   OutlookCom:ProcessedCategory    — Outlook category stamped on processed items (default: "PO-COM-Processed")
/// </summary>
public class OutlookComPollerService : BackgroundService
{
    private readonly MailPollerService                  _poller;
    private readonly IConfiguration                     _config;
    private readonly ILogger<OutlookComPollerService>   _log;

    // Matches RE:, FW:, FWD:, [EXTERNAL] etc.
    private static readonly Regex SubjectPrefixRegex =
        new(@"^(\s*(RE|FW|FWD)\s*:\s*|\s*\[.*?\]\s*)+", RegexOptions.Compiled | RegexOptions.IgnoreCase);

    public OutlookComPollerService(
        MailPollerService                  poller,
        IConfiguration                     config,
        ILogger<OutlookComPollerService>   log)
    {
        _poller = poller;
        _config = config;
        _log    = log;
    }

    [SupportedOSPlatform("windows")]
    protected override async Task ExecuteAsync(CancellationToken ct)
    {
        var mailbox = _config["OutlookCom:Mailbox"];
        if (string.IsNullOrWhiteSpace(mailbox))
        {
            _log.LogInformation("[OutlookCOM] OutlookCom:Mailbox not configured — poller disabled");
            return;
        }

        var interval = int.TryParse(_config["OutlookCom:PollIntervalSeconds"], out var i) ? i : 60;
        _log.LogInformation("[OutlookCOM] Polling {Mailbox} every {Interval}s via COM", mailbox, interval);

        // Give Outlook time to start before the first poll
        await Task.Delay(TimeSpan.FromSeconds(15), ct);

        while (!ct.IsCancellationRequested)
        {
            try
            {
                await PollAsync(mailbox, ct);
            }
            catch (OperationCanceledException) { break; }
            catch (Exception ex)
            {
                _log.LogError(ex, "[OutlookCOM] Poll error");
            }

            await Task.Delay(TimeSpan.FromSeconds(interval), ct);
        }
    }

    [SupportedOSPlatform("windows")]
    private async Task PollAsync(string mailbox, CancellationToken ct)
    {
        var processedCategory = _config["OutlookCom:ProcessedCategory"] ?? "PO-COM-Processed";
        var lookbackDays      = int.TryParse(_config["OutlookCom:LookbackDays"], out var ld) ? ld : 7;

        // All COM work happens synchronously on a dedicated STA thread.
        // We extract data into plain .NET objects before returning to the async context.
        List<OutlookPoMessage> messages;
        try
        {
            messages = await RunOnStaThreadAsync(() => CollectPoMessages(mailbox, processedCategory, lookbackDays));
        }
        catch (COMException ex) when (ex.HResult == unchecked((int)0x800401E3))
        {
            // MK_E_UNAVAILABLE — Outlook is not running
            _log.LogDebug("[OutlookCOM] Outlook not running — skipping poll");
            return;
        }
        catch (Exception ex)
        {
            _log.LogWarning(ex, "[OutlookCOM] Could not read from Outlook");
            return;
        }

        if (messages.Count == 0) return;
        _log.LogInformation("[OutlookCOM] Found {Count} unprocessed PO email(s)", messages.Count);

        foreach (var msg in messages)
        {
            if (ct.IsCancellationRequested) break;

            var processed = await _poller.ProcessPurchaseOrderCoreAsync(
                msg.Subject,
                msg.PlainBody,
                msg.PdfAttachments,
                msg.EntryId,
                msg.SentOn,
                msg.SenderHint,
                ct);

            // Stamp the item on an STA thread regardless of extraction success
            var category = processed ? processedCategory : "PO-COM-NoExtract";
            await RunOnStaThreadAsync(() => StampMessage(msg.EntryId, category));
        }
    }

    // ── Public fetch API ─────────────────────────────────────────────────────

    /// <summary>
    /// Fetches an Outlook mail item by its MAPI EntryID and returns it as a plain-data snapshot.
    /// Returns null if Outlook is not running or the item cannot be found.
    /// </summary>
    [SupportedOSPlatform("windows")]
    public async Task<OutlookPoMessage?> FetchByEntryIdAsync(string entryId)
    {
        try
        {
            return await RunOnStaThreadAsync(() => FetchMessageByEntryId(entryId));
        }
        catch (COMException ex) when (ex.HResult == unchecked((int)0x800401E3))
        {
            _log.LogWarning("[OutlookCOM] MK_E_UNAVAILABLE (0x800401E3) fetching EntryId {EntryId} — Outlook ROT not accessible from proxy process", entryId);
            return null;
        }
        catch (Exception ex)
        {
            _log.LogWarning(ex, "[OutlookCOM] Failed to fetch EntryId {EntryId}", entryId);
            return null;
        }
    }

    [SupportedOSPlatform("windows")]
    private OutlookPoMessage? FetchMessageByEntryId(string entryId)
    {
        dynamic outlook = GetActiveComObject("Outlook.Application");
        dynamic? item;
        try   { item = outlook.Session.GetItemFromID(entryId); }
        catch { return null; }
        if (item is null) return null;

        string subject = item.Subject ?? "";
        string body    = item.Body    ?? "";

        var pdfs = new List<(string Name, byte[] Bytes)>();
        dynamic attachments = item.Attachments;
        for (int i = 1; i <= (int)attachments.Count; i++)
        {
            dynamic att     = attachments[i];
            string  attName = att.FileName ?? att.DisplayName ?? "";
            if (!attName.EndsWith(".pdf", StringComparison.OrdinalIgnoreCase)) continue;

            var tmp = Path.Combine(Path.GetTempPath(), $"shredder_po_rx_{Guid.NewGuid():N}.pdf");
            try
            {
                att.SaveAsFile(tmp);
                pdfs.Add((attName, File.ReadAllBytes(tmp)));
            }
            finally
            {
                if (File.Exists(tmp)) File.Delete(tmp);
            }
            break; // first PDF only
        }

        string sentOn = "";
        try { sentOn = ((DateTime)item.SentOn).ToUniversalTime().ToString("o"); }
        catch { sentOn = DateTime.UtcNow.ToString("o"); }

        string senderName = "";
        try { senderName = item.SenderEmailAddress ?? item.SenderName ?? ""; }
        catch { }

        return new OutlookPoMessage(
            EntryId:        entryId,
            Subject:        subject,
            PlainBody:      body,
            SentOn:         sentOn,
            SenderHint:     senderName,
            PdfAttachments: pdfs);
    }

    // ── COM helpers ───────────────────────────────────────────────────────────

    [SupportedOSPlatform("windows")]
    private List<OutlookPoMessage> CollectPoMessages(string mailbox, string processedCategory, int lookbackDays = 7)
    {
        // GetActiveObject throws COMException if Outlook is not running
        dynamic outlook = GetActiveComObject("Outlook.Application");
        dynamic session = outlook.Session;

        // Find the store matching the configured mailbox display name or SMTP address
        dynamic? store = null;
        foreach (dynamic s in session.Stores)
        {
            string displayName = s.DisplayName ?? "";
            if (displayName.Equals(mailbox, StringComparison.OrdinalIgnoreCase))
            {
                store = s;
                break;
            }
            // Also check the account SMTP address if available
            try
            {
                dynamic? account = s.ExchangeAccount;
                string smtpAddress = account?.SmtpAddress ?? "";
                if (smtpAddress.Equals(mailbox, StringComparison.OrdinalIgnoreCase))
                {
                    store = s;
                    break;
                }
            }
            catch { /* not all stores have ExchangeAccount */ }
        }

        if (store is null)
        {
            _log.LogWarning("[OutlookCOM] Store '{Mailbox}' not found in Outlook", mailbox);
            return [];
        }

        // olFolderSentMail = 5
        dynamic sentItems = store.GetDefaultFolder(5);
        var result   = new List<OutlookPoMessage>();
        var cutoff   = DateTime.Now.AddDays(-lookbackDays);

        foreach (dynamic item in sentItems.Items)
        {
            try
            {
                // olMail = 43
                if ((int)item.Class != 43) continue;

                // Skip items older than the lookback window
                try { if ((DateTime)item.SentOn < cutoff) continue; }
                catch { continue; }

                string subject = item.Subject ?? "";
                var cleanSubject = SubjectPrefixRegex.Replace(subject, "").Trim();
                if (!cleanSubject.StartsWith("Purchase Order #HSK-PO", StringComparison.OrdinalIgnoreCase))
                    continue;

                // Skip already-stamped items
                string categories = item.Categories ?? "";
                if (categories.Contains(processedCategory, StringComparison.OrdinalIgnoreCase) ||
                    categories.Contains("PO-COM-NoExtract",  StringComparison.OrdinalIgnoreCase))
                    continue;

                // Extract body (plain text)
                string body = item.Body ?? "";

                // Extract PDF attachments — save to temp file to get bytes
                var pdfs = new List<(string Name, byte[] Bytes)>();
                dynamic attachments = item.Attachments;
                for (int i = 1; i <= (int)attachments.Count; i++)
                {
                    dynamic att = attachments[i];
                    string attName = att.FileName ?? att.DisplayName ?? "";
                    if (!attName.EndsWith(".pdf", StringComparison.OrdinalIgnoreCase)) continue;

                    var tmp = Path.Combine(Path.GetTempPath(), $"shredder_po_{Guid.NewGuid():N}.pdf");
                    try
                    {
                        att.SaveAsFile(tmp);
                        pdfs.Add((attName, File.ReadAllBytes(tmp)));
                    }
                    finally
                    {
                        if (File.Exists(tmp)) File.Delete(tmp);
                    }
                    break; // first PDF only
                }

                string sentOn = "";
                try { sentOn = ((DateTime)item.SentOn).ToUniversalTime().ToString("o"); }
                catch { sentOn = DateTime.UtcNow.ToString("o"); }

                string senderName = "";
                try { senderName = item.SenderEmailAddress ?? item.SenderName ?? ""; }
                catch { }

                result.Add(new OutlookPoMessage(
                    EntryId:        item.EntryID,
                    Subject:        subject,
                    PlainBody:      body,
                    SentOn:         sentOn,
                    SenderHint:     senderName,
                    PdfAttachments: pdfs));
            }
            catch (Exception ex)
            {
                _log.LogWarning(ex, "[OutlookCOM] Error reading Outlook item — skipping");
            }
        }

        return result;
    }

    [SupportedOSPlatform("windows")]
    private void StampMessage(string entryId, string category)
    {
        try
        {
            dynamic outlook = GetActiveComObject("Outlook.Application");
            dynamic item    = outlook.Session.GetItemFromID(entryId);
            string  current = item.Categories ?? "";

            if (!current.Contains(category, StringComparison.OrdinalIgnoreCase))
            {
                item.Categories = string.IsNullOrEmpty(current)
                    ? category
                    : $"{current}; {category}";
                item.Save();
            }
        }
        catch (Exception ex)
        {
            _log.LogWarning(ex, "[OutlookCOM] Failed to stamp category '{Category}' on EntryId {EntryId}",
                category, entryId);
        }
    }

    /// <summary>
    /// Runs <paramref name="func"/> on a new STA thread (required for COM automation)
    /// and returns its result as a Task.
    /// </summary>
    [SupportedOSPlatform("windows")]
    private static Task<T> RunOnStaThreadAsync<T>(Func<T> func)
    {
        var tcs = new TaskCompletionSource<T>(TaskCreationOptions.RunContinuationsAsynchronously);
        var thread = new Thread(() =>
        {
            try   { tcs.SetResult(func()); }
            catch (Exception ex) { tcs.SetException(ex); }
        });
        thread.SetApartmentState(ApartmentState.STA);
        thread.IsBackground = true;
        thread.Start();
        return tcs.Task;
    }

    [SupportedOSPlatform("windows")]
    private static Task RunOnStaThreadAsync(Action action) =>
        RunOnStaThreadAsync<bool>(() => { action(); return true; });

    // ── P/Invoke: GetActiveObject (removed from Marshal in .NET 5+) ──────────

    [DllImport("oleaut32.dll")]
    private static extern int GetActiveObject(
        ref Guid rclsid,
        IntPtr   pvReserved,
        [MarshalAs(UnmanagedType.IUnknown)] out object ppunk);

    [SupportedOSPlatform("windows")]
    private static dynamic GetActiveComObject(string progId)
    {
        var type = Type.GetTypeFromProgID(progId)
            ?? throw new COMException($"COM class '{progId}' not registered", unchecked((int)0x800401E3));
        var clsid = type.GUID;
        var hr    = GetActiveObject(ref clsid, IntPtr.Zero, out var obj);
        if (hr != 0) Marshal.ThrowExceptionForHR(hr);
        return (dynamic)obj;
    }
}

/// <summary>Plain-data snapshot of an Outlook PO email, extracted on the STA thread.</summary>
public record OutlookPoMessage(
    string                          EntryId,
    string                          Subject,
    string                          PlainBody,
    string                          SentOn,
    string                          SenderHint,
    List<(string Name, byte[] Bytes)> PdfAttachments);
