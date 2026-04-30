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
    private const uint EVENT_OBJECT_SHOW       = 0x8002;
    private const uint WINEVENT_OUTOFCONTEXT   = 0x0000;
    private const uint WINEVENT_SKIPOWNPROCESS = 0x0002;
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
    // Held as a field so the GC never collects the delegate while the hook is live
    private WinEventProc? _winEventCallback;

    public ZoomCallWatcherService(
        ILogger<ZoomCallWatcherService> log,
        IConfiguration config,
        RfqNotificationService notify)
    {
        _log    = log;
        _config = config;
        _notify = notify;
    }

    protected override Task ExecuteAsync(CancellationToken ct)
    {
        if (!_config.GetValue("Zoom:WatcherEnabled", true))
        {
            _log.LogInformation("[Zoom] Watcher disabled via Zoom:WatcherEnabled");
            return Task.CompletedTask;
        }

        var tcs    = new TaskCompletionSource();
        var thread = new Thread(() =>
        {
            var threadId = GetCurrentThreadId();
            IntPtr hook  = IntPtr.Zero;

            _winEventCallback = (_, eventType, hwnd, idObject, _, _, _) =>
                OnWinEvent(eventType, hwnd, idObject);
            var callback = _winEventCallback;

            try
            {
                // Force creation of the thread message queue before registering the hook.
                // WINEVENT_OUTOFCONTEXT posts callbacks to this queue; without it the
                // hook registers but never delivers events.
                PeekMessage(out _, IntPtr.Zero, 0, 0, PM_NOREMOVE);

                hook = SetWinEventHook(
                    EVENT_OBJECT_SHOW, EVENT_OBJECT_SHOW,
                    IntPtr.Zero,
                    callback,
                    0,    // all processes
                    0,    // all threads
                    WINEVENT_OUTOFCONTEXT | WINEVENT_SKIPOWNPROCESS);

                if (hook == IntPtr.Zero)
                {
                    _log.LogError("[Zoom] SetWinEventHook failed (error {Err})", Marshal.GetLastWin32Error());
                    tcs.TrySetResult();
                    return;
                }

                _log.LogInformation("[Zoom] WinEvent hook active (threadId={ThreadId}) — Zoom windows will be logged on incoming calls", threadId);

                ct.Register(() => PostThreadMessage(threadId, WM_QUIT, IntPtr.Zero, IntPtr.Zero));

                // Native message loop — WinEvent callbacks are delivered here
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
                if (hook != IntPtr.Zero) UnhookWinEvent(hook);
                tcs.TrySetResult();
            }
        });

        thread.SetApartmentState(ApartmentState.STA);
        thread.IsBackground = true;
        thread.Name         = "ZoomWinEventWatcher";
        thread.Start();

        return tcs.Task;
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

            if (_config.GetValue("Zoom:DebugAllWindows", false))
            {
                string? winName  = null;
                string? winClass = null;
                try { winName  = AutomationElement.FromHandle(hwnd).Current.Name;      } catch { }
                try { winClass = AutomationElement.FromHandle(hwnd).Current.ClassName; } catch { }
                _log.LogInformation("[Zoom:dbg] WindowShown proc='{Proc}' name='{Name}' class='{Class}'",
                    procName, winName ?? "", winClass ?? "");
            }

            if (!procName.StartsWith("Zoom", StringComparison.OrdinalIgnoreCase)) return;

            try
            {
                var el    = AutomationElement.FromHandle(hwnd);
                var name  = TryGet(() => el.Current.Name)      ?? "";
                var cls   = TryGet(() => el.Current.ClassName) ?? "";

                // Incoming Zoom Phone call notification
                if (cls == "SipCallNormalIncomingCallWindow" && name.Contains("is calling you", StringComparison.OrdinalIgnoreCase))
                {
                    var (callerName, callerPhone) = ParseIncomingCallTitle(name);
                    _log.LogInformation("[Zoom] Incoming call — name='{Name}' phone='{Phone}'", callerName, callerPhone);
                    _notify.NotifyIncomingCall(callerName, callerPhone);
                }

                var sb = new StringBuilder();
                sb.AppendLine($"[Zoom] WindowShown proc='{procName}' name='{name}' class='{cls}'");
                DumpTree(el, sb, depth: 0);
                _log.LogInformation("{Tree}", sb.ToString());
            }
            catch (Exception ex)
            {
                _log.LogDebug(ex, "[Zoom] tree dump failed (element may be stale)");
            }
        }
        catch (Exception ex)
        {
            _log.LogDebug(ex, "[Zoom] OnWinEvent error");
        }
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
