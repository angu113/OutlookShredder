using System.Diagnostics;
using System.Text;
using System.Windows.Automation;

namespace OutlookShredder.Proxy.Services;

/// <summary>
/// Listens for Windows UI Automation WindowOpenedEvents and logs the full element
/// tree for any window belonging to Zoom.exe.  Used to capture the incoming-call
/// notification structure so we can later extract the caller number.
///
/// Runs on a dedicated STA thread (required for UIA COM initialisation).
/// Toggle with Zoom:WatcherEnabled in appsettings (default true).
/// </summary>
public class ZoomCallWatcherService : BackgroundService
{
    private const int MaxTreeDepth = 8;

    private readonly ILogger<ZoomCallWatcherService> _log;
    private readonly IConfiguration                  _config;

    public ZoomCallWatcherService(
        ILogger<ZoomCallWatcherService> log,
        IConfiguration config)
    {
        _log    = log;
        _config = config;
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
            try
            {
                var handler = new AutomationEventHandler((sender, _) =>
                {
                    if (sender is AutomationElement el) OnWindowOpened(el);
                });

                Automation.AddAutomationEventHandler(
                    WindowPattern.WindowOpenedEvent,
                    AutomationElement.RootElement,
                    TreeScope.Subtree,
                    handler);

                _log.LogInformation("[Zoom] UIA WindowOpened hook active — Zoom window trees will be logged on incoming calls");

                using (ct.Register(() =>
                {
                    try { Automation.RemoveAutomationEventHandler(WindowPattern.WindowOpenedEvent, AutomationElement.RootElement, handler); }
                    catch { /* best-effort cleanup */ }
                    tcs.TrySetResult();
                }))
                {
                    tcs.Task.GetAwaiter().GetResult();
                }
            }
            catch (Exception ex)
            {
                _log.LogError(ex, "[Zoom] UIA hook setup failed");
                tcs.TrySetResult();
            }
        });

        thread.SetApartmentState(ApartmentState.STA);
        thread.IsBackground = true;
        thread.Name         = "ZoomUiaWatcher";
        thread.Start();

        return tcs.Task;
    }

    private void OnWindowOpened(AutomationElement el)
    {
        try
        {
            int pid;
            try   { pid = el.Current.ProcessId; }
            catch { return; }

            string procName;
            try   { procName = Process.GetProcessById(pid).ProcessName; }
            catch { return; }

            if (!procName.Equals("Zoom", StringComparison.OrdinalIgnoreCase)) return;

            var sb = new StringBuilder();
            sb.AppendLine($"[Zoom] WindowOpened name='{TryGet(() => el.Current.Name)}' class='{TryGet(() => el.Current.ClassName)}'");
            DumpTree(el, sb, depth: 0);
            _log.LogInformation("{Tree}", sb.ToString());
        }
        catch (Exception ex)
        {
            _log.LogDebug(ex, "[Zoom] OnWindowOpened error (element may be stale)");
        }
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
