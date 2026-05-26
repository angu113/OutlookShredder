using System.Diagnostics;
using System.Runtime.InteropServices;
using System.Text;
using System.Windows.Automation;

namespace OutlookShredder.Proxy.Services;

/// <summary>
/// Uses SetWinEventHook (EVENT_OBJECT_SHOW, WINEVENT_OUTOFCONTEXT) to detect new
/// windows from Zoom.exe and dumps their UIA element tree to the proxy log.
/// Runs on a dedicated STA thread with a native Win32 message loop.
/// Toggle with Zoom:WatcherEnabled in appsettings (default true).
/// Set Zoom:DebugAllWindows=true to log every window open (for diagnosis).
/// </summary>
public class ZoomCallWatcherService : BackgroundService
{
    private const int MaxTreeDepth = 8;

    // WinEvent
    private const uint EVENT_OBJECT_SHOW        = 0x8002;
    private const uint EVENT_OBJECT_NAMECHANGE  = 0x800C;
    private const uint WINEVENT_OUTOFCONTEXT    = 0x0000;
    private const uint WINEVENT_SKIPOWNPROCESS  = 0x0002;
    private const int  OBJID_WINDOW            = 0;
    private const uint WM_QUIT                 = 0x0012;

    [DllImport("user32.dll", SetLastError = true)]
    private static extern IntPtr SetWinEventHook(
        uint eventMin, uint eventMax,
        IntPtr hmodWinEventProc,
        WinEventProc lpfnWinEventProc,
        uint idProcess, uint idThread,
        uint dwFlags);

    [DllImport("user32.dll")]
    private static extern bool UnhookWinEvent(IntPtr hWinEventHook);

    [DllImport("user32.dll")]
    private static extern int GetMessage(out NativeMsg msg, IntPtr hWnd, uint min, uint max);

    [DllImport("user32.dll")]
    private static extern bool TranslateMessage(ref NativeMsg msg);

    [DllImport("user32.dll")]
    private static extern IntPtr DispatchMessage(ref NativeMsg msg);

    [DllImport("user32.dll")]
    private static extern bool PostThreadMessage(int threadId, uint msg, IntPtr wp, IntPtr lp);

    // PeekMessage with PM_NOREMOVE — used to force creation of the thread message queue
    // before SetWinEventHook registers the hook. WINEVENT_OUTOFCONTEXT requires the
    // calling thread to have a message queue before the hook is registered.
    private const uint PM_NOREMOVE = 0x0000;
    [DllImport("user32.dll")]
    private static extern bool PeekMessage(out NativeMsg msg, IntPtr hWnd, uint min, uint max, uint removeMsg);

    [DllImport("kernel32.dll")]
    private static extern int GetCurrentThreadId();

    [DllImport("user32.dll")]
    private static extern bool EnumChildWindows(IntPtr hwndParent, EnumWindowsProc lpEnumFunc, IntPtr lParam);

    [DllImport("user32.dll")]
    private static extern bool EnumWindows(EnumWindowsProc lpEnumFunc, IntPtr lParam);

    [DllImport("user32.dll", CharSet = CharSet.Auto)]
    private static extern int GetWindowText(IntPtr hWnd, StringBuilder lpString, int nMaxCount);

    [DllImport("user32.dll")]
    private static extern uint GetWindowThreadProcessId(IntPtr hWnd, out uint lpdwProcessId);

    private delegate bool EnumWindowsProc(IntPtr hwnd, IntPtr lParam);

    [StructLayout(LayoutKind.Sequential)]
    private struct NativeMsg
    {
        public IntPtr hwnd;
        public uint   message;
        public IntPtr wParam;
        public IntPtr lParam;
        public uint   time;
        public int    ptX, ptY;
    }

    private delegate void WinEventProc(
        IntPtr hWinEventHook, uint eventType, IntPtr hwnd,
        int idObject, int idChild,
        uint dwEventThread, uint dwmsEventTime);

    private readonly ILogger<ZoomCallWatcherService> _log;
    private readonly IConfiguration                  _config;
    private readonly RfqNotificationService          _notify;
    private readonly SharePointService               _sp;
    private readonly CustomerCacheService            _crmCache;
    private readonly ProxyLeaseService               _lease;
    // Held as a field so the GC never collects the delegate while the hook is live
    private WinEventProc? _winEventCallback;

    public ZoomCallWatcherService(
        ILogger<ZoomCallWatcherService> log,
        IConfiguration config,
        RfqNotificationService notify,
        SharePointService sp,
        CustomerCacheService crmCache,
        ProxyLeaseService lease)
    {
        _log      = log;
        _config   = config;
        _crmCache = crmCache;
        _notify   = notify;
        _sp       = sp;
        _lease    = lease;
    }

    protected override async Task ExecuteAsync(CancellationToken ct)
    {
        if (!_config.GetValue("Zoom:WatcherEnabled", true))
        {
            _log.LogInformation("[Zoom] Watcher disabled via Zoom:WatcherEnabled");
            return;
        }

        // Steal the lease on startup — the most-recently-started proxy always wins.
        // The previous holder stops its hook within ~15 s when its renewal detects loss.
        // During that brief overlap both hooks may fire; the client deduplicates by phone+30s.
        await _lease.StealLeaseAsync(ct);

        while (!ct.IsCancellationRequested)
        {
            if (!_lease.IsLeaseHolder)
            {
                _log.LogInformation("[Zoom] Waiting for lease (another proxy is active)");
                try { await Task.Delay(TimeSpan.FromSeconds(15), ct); } catch (OperationCanceledException) { return; }
                continue;
            }

            _log.LogInformation("[Zoom] Lease held — starting WinEvent hook");

            // Run the hook; returns when lease is lost or ct is cancelled.
            await RunHookAsync(ct);

            if (!ct.IsCancellationRequested)
            {
                _log.LogInformation("[Zoom] Hook stopped (lease lost) — will re-check in 15 s");
                try { await Task.Delay(TimeSpan.FromSeconds(15), ct); } catch (OperationCanceledException) { return; }
            }
        }
    }

    private async Task RunHookAsync(CancellationToken ct)
    {
        // Combine the host ct with a lease-loss cancellation so we can stop the
        // STA thread either when the app shuts down OR when we lose the lease.
        // NOTE: must be async + await tcs.Task so the using block stays alive
        // until the STA thread exits. A non-async Task return disposes leaseCts
        // immediately on return, before the STA thread reaches Token.Register.
        using var leaseCts = CancellationTokenSource.CreateLinkedTokenSource(ct);

        var tcs = new TaskCompletionSource();

        var thread = new Thread(() =>
        {
            var threadId  = GetCurrentThreadId();
            IntPtr hookShow = IntPtr.Zero;
            IntPtr hookName = IntPtr.Zero;

            _winEventCallback = (_, eventType, hwnd, idObject, _, _, _) =>
                OnWinEvent(eventType, hwnd, idObject);
            var callback = _winEventCallback;

            try
            {
                PeekMessage(out _, IntPtr.Zero, 0, 0, PM_NOREMOVE);

                var flags = WINEVENT_OUTOFCONTEXT | WINEVENT_SKIPOWNPROCESS;

                hookShow = SetWinEventHook(
                    EVENT_OBJECT_SHOW, EVENT_OBJECT_SHOW,
                    IntPtr.Zero, callback, 0, 0, flags);

                // Second hook for name-change events only — avoids capturing the
                // noisy range of events between SHOW and NAMECHANGE (focus, state, etc.)
                hookName = SetWinEventHook(
                    EVENT_OBJECT_NAMECHANGE, EVENT_OBJECT_NAMECHANGE,
                    IntPtr.Zero, callback, 0, 0, flags);

                if (hookShow == IntPtr.Zero || hookName == IntPtr.Zero)
                {
                    _log.LogError("[Zoom] SetWinEventHook failed (error {Err})", Marshal.GetLastWin32Error());
                    tcs.TrySetResult();
                    return;
                }

                _log.LogInformation("[Zoom] WinEvent hook active (threadId={ThreadId})", threadId);

                // Quit the message loop on app shutdown OR lease loss.
                leaseCts.Token.Register(() =>
                    PostThreadMessage(threadId, WM_QUIT, IntPtr.Zero, IntPtr.Zero));

                while (GetMessage(out var msg, IntPtr.Zero, 0, 0) > 0)
                {
                    TranslateMessage(ref msg);
                    DispatchMessage(ref msg);
                }
            }
            catch (Exception ex)
            {
                _log.LogError(ex, "[Zoom] WinEvent hook setup failed");
            }
            finally
            {
                if (hookShow != IntPtr.Zero) UnhookWinEvent(hookShow);
                if (hookName != IntPtr.Zero) UnhookWinEvent(hookName);
                tcs.TrySetResult();
            }
        });

        thread.SetApartmentState(ApartmentState.STA);
        thread.IsBackground = true;
        thread.Name         = "ZoomWinEventWatcher";
        thread.Start();

        // Monitor lease on a background thread while the STA message loop runs.
        _ = Task.Run(async () =>
        {
            while (!leaseCts.Token.IsCancellationRequested)
            {
                try { await Task.Delay(TimeSpan.FromSeconds(15), leaseCts.Token); }
                catch (OperationCanceledException) { return; }

                if (!_lease.IsLeaseHolder)
                {
                    _log.LogInformation("[Zoom] Lease no longer held — stopping hook");
                    leaseCts.Cancel();
                    return;
                }
            }
        }, leaseCts.Token);

        await tcs.Task;
    }

    private void OnWinEvent(uint eventType, IntPtr hwnd, int idObject)
    {
        // Only top-level window objects (idObject == OBJID_WINDOW)
        if (idObject != OBJID_WINDOW || hwnd == IntPtr.Zero) return;

        try
        {
            int pid;
            try
            {
                var el = AutomationElement.FromHandle(hwnd);
                pid = el.Current.ProcessId;
            }
            catch { return; }

            string procName;
            try   { procName = Process.GetProcessById(pid).ProcessName; }
            catch { return; }

            var debugAll = _config.GetValue("Zoom:DebugAllWindows", false);
            if (debugAll)
            {
                string? winName  = null;
                string? winClass = null;
                try { winName  = AutomationElement.FromHandle(hwnd).Current.Name;      } catch { }
                try { winClass = AutomationElement.FromHandle(hwnd).Current.ClassName; } catch { }
                var evtLabel = eventType == EVENT_OBJECT_NAMECHANGE ? "NameChange" : "WindowShown";
                _log.LogInformation("[Zoom:dbg] {Evt} proc='{Proc}' name='{Name}' class='{Class}'",
                    evtLabel, procName, winName ?? "", winClass ?? "");
            }

            if (!procName.StartsWith("Zoom", StringComparison.OrdinalIgnoreCase)) return;

            try
            {
                var el   = AutomationElement.FromHandle(hwnd);
                var name = TryGet(() => el.Current.Name)      ?? "";
                var cls  = TryGet(() => el.Current.ClassName) ?? "";

                // Always log every Zoom window event so we can see the detection pipeline.
                var evtLabel = eventType == EVENT_OBJECT_NAMECHANGE ? "NameChange" : "WindowShown";
                var namePreview = name.Length > 100 ? name[..100] + "..." : name;
                _log.LogInformation("[Zoom] {Evt} class='{Class}' name='{Name}'", evtLabel, cls, namePreview);

                // Fast path: "is calling you" text is already on the window name
                // (works for both EVENT_OBJECT_SHOW and EVENT_OBJECT_NAMECHANGE).
                if (name.Contains("is calling you", StringComparison.OrdinalIgnoreCase))
                {
                    _log.LogInformation("[Zoom] Fast path triggered — firing incoming call");
                    FireIncomingCall(name);
                }
                // Delayed path (SHOW only): Zoom wraps the call UI inside a ZoomShadowFrameClass window.
                // UIA children aren't populated at EVENT_OBJECT_SHOW time, so try three Win32
                // fallbacks after a short delay: UIA subtree → EnumChildWindows → EnumWindows.
                // For NAMECHANGE events the name is already current — skip the delayed search.
                else if (eventType == EVENT_OBJECT_SHOW &&
                         (cls == "ZoomShadowFrameClass" || cls == "SipCallNormalIncomingCallWindow"))
                {
                    _log.LogInformation("[Zoom] Delayed search starting for class '{Class}' (400ms delay)", cls);
                    var capturedHwnd = hwnd;
                    var capturedPid  = pid;
                    var capturedEl   = el;
                    var capturedCls  = cls;
                    var capturedDebug = debugAll;
                    _ = Task.Run(async () =>
                    {
                        await Task.Delay(400);
                        var callText = FindCallTextInTree(capturedEl)
                                    ?? FindCallTextInChildWindows(capturedHwnd, capturedDebug)
                                    ?? FindCallTextInZoomTopLevel(capturedPid, capturedDebug);
                        if (callText is not null)
                        {
                            _log.LogInformation("[Zoom] Delayed search found call text: '{Text}'", callText);
                            FireIncomingCall(callText);
                        }
                        else
                        {
                            _log.LogWarning("[Zoom] Delayed search: no call text found in '{Class}' children or siblings after 400ms", capturedCls);
                        }
                    });
                }
                else if (eventType == EVENT_OBJECT_SHOW)
                {
                    _log.LogDebug("[Zoom] WindowShown ignored — class '{Class}' is not a known call window", cls);
                }

                if (debugAll && eventType == EVENT_OBJECT_SHOW)
                {
                    var sb = new StringBuilder();
                    sb.AppendLine($"[Zoom] WindowShown proc='{procName}' name='{name}' class='{cls}'");
                    DumpTree(el, sb, depth: 0);
                    _log.LogInformation("{Tree}", sb.ToString());
                }
            }
            catch (Exception ex)
            {
                _log.LogDebug(ex, "[Zoom] tree walk failed (element may be stale)");
            }
        }
        catch (Exception ex)
        {
            _log.LogDebug(ex, "[Zoom] OnWinEvent error");
        }
    }

    /// <summary>
    /// Recursively searches the UIA element tree for any node whose Name contains
    /// "is calling you". Returns the first matching name, or null if not found.
    /// </summary>
    private string? FindCallTextInTree(AutomationElement root, int depth = 0)
    {
        if (depth > MaxTreeDepth) return null;
        try
        {
            var name = TryGet(() => root.Current.Name) ?? "";
            if (name.Contains("is calling you", StringComparison.OrdinalIgnoreCase)) return name;
            var walker = TreeWalker.RawViewWalker;
            var child  = walker.GetFirstChild(root);
            while (child != null)
            {
                var found = FindCallTextInTree(child, depth + 1);
                if (found is not null) return found;
                child = walker.GetNextSibling(child);
            }
        }
        catch { }
        return null;
    }

    private void FireIncomingCall(string callTitle)
    {
        var (callerName, callerPhone) = ParseIncomingCallTitle(callTitle);
        _log.LogInformation("[Zoom] Incoming call — raw='{Raw}' → name='{Name}' phone='{Phone}'",
            callTitle.Length > 120 ? callTitle[..120] + "..." : callTitle, callerName, callerPhone);

        _ = Task.Run(async () =>
        {
            // CRM lookup
            var allCrm = !string.IsNullOrWhiteSpace(callerPhone)
                ? _crmCache.LookupAllByPhone(callerPhone)
                : [];
            var crm = allCrm.Count > 0 ? allCrm[0] : null;

            if (allCrm.Count == 0)
                _log.LogInformation("[Zoom] CRM: no match for phone '{Phone}'", callerPhone);
            else
                _log.LogInformation("[Zoom] CRM: {Count} match(es) — primary bp='{Bp}' contact='{Contact}'",
                    allCrm.Count, crm!.BusinessPartner, crm.ContactName ?? "(none)");

            // SharePoint call log write
            string spItemId = "";
            try
            {
                spItemId = await _sp.WritePhoneCallLogAsync(
                    callerName, callerPhone ?? "",
                    crm?.BusinessPartner, crm?.ContactName, crm?.PopupMessage,
                    DateTimeOffset.UtcNow, CancellationToken.None);
                _log.LogInformation("[Zoom] Call log written — SpItemId={SpItemId}", spItemId);
            }
            catch (Exception ex)
            {
                _log.LogWarning(ex, "[Zoom] Call log write FAILED for '{Name}' / '{Phone}'", callerName, callerPhone);
            }

            // Publish to Service Bus + SSE
            _notify.NotifyIncomingCall(callerName, callerPhone,
                crm?.BusinessPartner, crm?.PopupMessage, crm?.ContactName,
                callLogSpItemId: spItemId,
                allMatches: allCrm.Count > 1 ? allCrm : null);
            _log.LogInformation("[Zoom] Notification published — name='{Name}' phone='{Phone}' spItemId='{SpItemId}'",
                callerName, callerPhone, spItemId);
        });
    }

    /// Enumerates Win32 child HWNDs of <paramref name="parentHwnd"/> via GetWindowText.
    /// Catches the case where the call content is a child window with no UIA exposure.
    private string? FindCallTextInChildWindows(IntPtr parentHwnd, bool debugAll = false)
    {
        string? result = null;
        var childTexts = debugAll ? new List<string>() : null;
        EnumWindowsProc cb = (hwnd, _) =>
        {
            var sb = new StringBuilder(512);
            var len = GetWindowText(hwnd, sb, sb.Capacity);
            if (len > 0)
            {
                var text = sb.ToString();
                childTexts?.Add(text);
                if (text.Contains("is calling you", StringComparison.OrdinalIgnoreCase))
                { result = text; return false; }
            }
            return true;
        };
        EnumChildWindows(parentHwnd, cb, IntPtr.Zero);
        GC.KeepAlive(cb);
        if (debugAll)
            _log.LogInformation("[Zoom:dbg] EnumChildWindows found {N} text(s): [{Texts}]",
                childTexts!.Count, string.Join(", ", childTexts.Select(t => $"'{t}'")));
        return result;
    }

    /// Enumerates ALL top-level windows belonging to the Zoom process via GetWindowText.
    /// Catches the case where the call content window is a sibling (not a child) of the
    /// shadow frame — i.e. a separate top-level HWND that didn't fire EVENT_OBJECT_SHOW.
    private string? FindCallTextInZoomTopLevel(int zoomPid, bool debugAll = false)
    {
        string? result = null;
        var topTexts = debugAll ? new List<string>() : null;
        EnumWindowsProc cb = (hwnd, _) =>
        {
            GetWindowThreadProcessId(hwnd, out var winPid);
            if (winPid != (uint)zoomPid) return true;
            var sb = new StringBuilder(512);
            if (GetWindowText(hwnd, sb, sb.Capacity) > 0)
            {
                var text = sb.ToString();
                topTexts?.Add(text);
                if (text.Contains("is calling you", StringComparison.OrdinalIgnoreCase))
                { result = text; return false; }
            }
            return true;
        };
        EnumWindows(cb, IntPtr.Zero);
        GC.KeepAlive(cb);
        if (debugAll)
            _log.LogInformation("[Zoom:dbg] EnumWindows (Zoom PID {Pid}) found {N} window(s): [{Texts}]",
                zoomPid, topTexts!.Count, string.Join(", ", topTexts.Select(t => $"'{t}'")));
        return result;
    }

    /// <summary>
    /// Parses "Angus Wathen (9 7 3 ) 7 5 2 -2 1 9 3  is calling you…"
    /// into ("Angus Wathen", "(973) 752-2193").
    /// </summary>
    private static (string name, string phone) ParseIncomingCallTitle(string title)
    {
        const string suffix = " is calling you";
        var suffixIdx = title.IndexOf(suffix, StringComparison.OrdinalIgnoreCase);
        if (suffixIdx < 0) return (title.Trim(), "");

        var before    = title[..suffixIdx].Trim();  // "Angus Wathen (9 7 3 ) 7 5 2 -2 1 9 3"
        var parenIdx  = before.IndexOf('(');
        if (parenIdx <= 0) return (before, "");

        var callerName  = before[..parenIdx].Trim();
        // Strip all spaces within the raw phone string, then reformat "(NXX)NXX-XXXX" → "(NXX) NXX-XXXX"
        var rawPhone    = before[parenIdx..].Trim();
        var digits      = string.Concat(rawPhone.Where(char.IsDigit));  // "9737522193"
        var callerPhone = digits.Length == 10
            ? $"({digits[..3]}) {digits[3..6]}-{digits[6..]}"
            : rawPhone.Replace(" ", "");

        return (callerName, callerPhone);
    }

    private static void DumpTree(AutomationElement el, StringBuilder sb, int depth)
    {
        if (depth > MaxTreeDepth) return;

        var pad   = new string(' ', depth * 2);
        var name  = TryGet(() => el.Current.Name)        ?? "";
        var cls   = TryGet(() => el.Current.ClassName)   ?? "";
        var type  = TryGet(() => el.Current.ControlType.ProgrammaticName) ?? "";
        var value = "";

        if (el.TryGetCurrentPattern(ValuePattern.Pattern, out var obj) && obj is ValuePattern vp)
            value = TryGet(() => vp.Current.Value) ?? "";

        sb.AppendLine(value.Length > 0
            ? $"{pad}[{type}] name='{name}' class='{cls}' value='{value}'"
            : $"{pad}[{type}] name='{name}' class='{cls}'");

        try
        {
            var walker = TreeWalker.RawViewWalker;
            var child  = walker.GetFirstChild(el);
            while (child != null)
            {
                DumpTree(child, sb, depth + 1);
                child = walker.GetNextSibling(child);
            }
        }
        catch { /* element became stale mid-walk */ }
    }

    private static string? TryGet(Func<string> fn)
    {
        try   { return fn(); }
        catch { return null; }
    }
}
