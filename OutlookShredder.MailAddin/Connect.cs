using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using System.Threading;
using System.Windows.Forms;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace OutlookShredder.MailAddin;

[ComVisible(true)]
[ClassInterface(ClassInterfaceType.None)]
[Guid("7F3A1D5C-4E2B-4A8F-9C6D-1B3E5F7A9C2D")]
[ProgId("OutlookShredder.MailAddin.Connect")]
public class Connect : IDTExtensibility2
{
    // Static constructor runs before the first instance is created, before the CLR
    // JITs the instance constructor that references Outlook.Items etc.
    // Outlook's CLR host sets AppBase to Outlook's own directory, so sibling DLLs
    // (e.g. Microsoft.Office.Interop.Outlook.dll) are invisible to the default binder.
    // This handler resolves them from the directory of this DLL instead.
    // Writes to log without touching any external type (safe for static-init context)
    private static void EarlyLog(string msg)
    {
        try
        {
            var loc  = typeof(Connect).Assembly.Location;
            var dir  = System.IO.Path.GetDirectoryName(loc) ?? ".";
            var path = System.IO.Path.Combine(dir, "addin.log");
            System.IO.File.AppendAllText(path,
                $"{DateTime.Now:yyyy-MM-dd HH:mm:ss} [STATIC] {msg}\r\n");
        }
        catch { }
    }

    static Connect()
    {
        EarlyLog("Static ctor entered");
        var dir = System.IO.Path.GetDirectoryName(typeof(Connect).Assembly.Location) ?? "";
        EarlyLog($"DLL dir: {dir}");
        AppDomain.CurrentDomain.AssemblyResolve += (_, args) =>
        {
            var name      = new System.Reflection.AssemblyName(args.Name).Name;
            var candidate = System.IO.Path.Combine(dir, name + ".dll");
            var exists    = System.IO.File.Exists(candidate);
            EarlyLog($"AssemblyResolve: {name} -> {(exists ? "FOUND" : "not found")}");
            return exists ? System.Reflection.Assembly.LoadFrom(candidate) : null;
        };
        EarlyLog("Static ctor complete, resolver registered");
    }

    private Outlook.Application?      _app;
    private AddinConfig               _config      = new AddinConfig();
    private AddinHttpServer?          _server;
    private ProxyPushClient?          _pushClient;
    private SynchronizationContext?   _staCtx;

    // COM event sources — must stay rooted or events silently stop firing
    private readonly List<Outlook.Items>                    _monitoredItems  = new List<Outlook.Items>();
    private readonly Dictionary<Outlook.Items, string>      _itemsToStoreId  = new Dictionary<Outlook.Items, string>();
    private readonly Dictionary<Outlook.Items, string>      _itemsToName     = new Dictionary<Outlook.Items, string>();

    public Connect()
    {
        EarlyLog("Instance ctor called");
    }

    public void OnConnection(object Application, ext_ConnectMode ConnectMode, object AddInInst, ref Array custom)
    {
        EarlyLog($"OnConnection called: mode={ConnectMode}");
        try
        {
            _app = (Outlook.Application)Application;
            EarlyLog("OnConnection: _app assigned ok");
        }
        catch (Exception ex)
        {
            EarlyLog($"OnConnection error: {ex}");
            throw;
        }
    }

    public void OnDisconnection(ext_DisconnectMode RemoveMode, ref Array custom)
    {
        _server?.Stop();
        foreach (var items in _monitoredItems)
        {
            try { Marshal.ReleaseComObject(items); }
            catch { }
        }
        _monitoredItems.Clear();
        _itemsToStoreId.Clear();
        _itemsToName.Clear();
    }

    public void OnAddInsUpdate(ref Array custom) { }

    public void OnStartupComplete(ref Array custom)
    {
        // Ensure we have an STA SynchronizationContext so background threads
        // can marshal Outlook COM calls back here via _staCtx.Post(...)
        if (SynchronizationContext.Current is null)
            SynchronizationContext.SetSynchronizationContext(new WindowsFormsSynchronizationContext());
        _staCtx = SynchronizationContext.Current!;

        try
        {
            _config     = AddinConfig.Load();
            _pushClient = new ProxyPushClient(_config.ProxyUrl);
            _server     = new AddinHttpServer(_config.ListenPort, _app!, _staCtx);
            _server.Start();
            HookMailboxes();
            ProxyPushClient.Log("Add-in started");
        }
        catch (Exception ex)
        {
            ProxyPushClient.Log($"OnStartupComplete error: {ex}");
        }
    }

    public void OnBeginShutdown(ref Array custom) { }

    private void HookMailboxes()
    {
        if (_app is null) return;
        if (_config.MonitoredMailboxes.Count == 0)
        {
            ProxyPushClient.Log("No monitored mailboxes configured — nothing to hook");
            return;
        }

        foreach (Outlook.Store store in _app.Session.Stores)
        {
            var displayName = store.DisplayName ?? string.Empty;
            foreach (var target in _config.MonitoredMailboxes)
            {
                if (!displayName.Equals(target, StringComparison.OrdinalIgnoreCase))
                    continue;

                try
                {
                    var inbox = (Outlook.MAPIFolder)store.GetDefaultFolder(Outlook.OlDefaultFolders.olFolderInbox);
                    var items = inbox.Items;

                    items.ItemAdd += OnItemAdded;
                    _monitoredItems.Add(items);
                    _itemsToStoreId[items] = store.StoreID;
                    _itemsToName[items]    = displayName;

                    var idSnippet = store.StoreID.Length > 8 ? store.StoreID.Substring(0, 8) : store.StoreID;
                    ProxyPushClient.Log($"Hooked inbox of '{displayName}' (StoreID={idSnippet}...)");
                }
                catch (Exception ex)
                {
                    ProxyPushClient.Log($"Failed to hook '{target}': {ex.Message}");
                }
            }
        }
    }

    private void OnItemAdded(object Item)
    {
        if (Item is not Outlook.MailItem mail) return;

        string? storeId     = null;
        string? mailboxName = null;
        try
        {
            if (mail.Parent is Outlook.MAPIFolder parentFolder)
            {
                var parentStoreId = parentFolder.StoreID;
                foreach (var kvp in _itemsToStoreId)
                {
                    if (kvp.Value == parentStoreId)
                    {
                        storeId = kvp.Value;
                        _itemsToName.TryGetValue(kvp.Key, out mailboxName);
                        break;
                    }
                }
            }
        }
        catch { }

        try
        {
            var payload = OutlookReader.BuildPayload(mail, storeId, mailboxName);
            // Fire-and-forget on a thread-pool thread — don't block Outlook's STA
            _ = _pushClient!.SendAsync(payload);
        }
        catch (Exception ex)
        {
            ProxyPushClient.Log($"OnItemAdded error: {ex.Message}");
        }
    }
}
