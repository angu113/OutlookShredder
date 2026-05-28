namespace OutlookShredder.Proxy.Services;

/// <summary>
/// Watches %LOCALAPPDATA%\ShredderData\shutdown.signal for a GUID that matches
/// the persistent key in shutdown.key. When it matches, stops the proxy gracefully.
///
/// Used by reinstall.ps1 to avoid force-killing the process during upgrades.
/// The key persists across installs so the signal works even when the running
/// build differs from the installer (e.g. dev instance vs prod reinstall).
/// </summary>
public sealed class ShutdownWatcherService : IHostedService, IDisposable
{
    private static readonly string DataDir    = Path.Combine(
        Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
        "ShredderData");
    private static readonly string KeyFile    = Path.Combine(DataDir, "shutdown.key");
    private static readonly string SignalFile = Path.Combine(DataDir, "shutdown.signal");

    private readonly IHostApplicationLifetime       _lifetime;
    private readonly ILogger<ShutdownWatcherService> _log;
    private FileSystemWatcher? _watcher;
    private string _key = "";
    private int    _handled;

    public ShutdownWatcherService(
        IHostApplicationLifetime        lifetime,
        ILogger<ShutdownWatcherService> log)
    {
        _lifetime = lifetime;
        _log      = log;
    }

    public Task StartAsync(CancellationToken ct)
    {
        Directory.CreateDirectory(DataDir);
        _key = EnsureKey();

        try { if (File.Exists(SignalFile)) File.Delete(SignalFile); } catch { }

        _watcher = new FileSystemWatcher(DataDir, "shutdown.signal")
        {
            NotifyFilter        = NotifyFilters.FileName | NotifyFilters.LastWrite,
            EnableRaisingEvents = true,
        };
        _watcher.Created += OnSignal;
        _watcher.Changed += OnSignal;

        _log.LogInformation("[Shutdown] Watcher ready");
        return Task.CompletedTask;
    }

    public Task StopAsync(CancellationToken ct)
    {
        _watcher?.Dispose();
        try { if (File.Exists(SignalFile)) File.Delete(SignalFile); } catch { }
        return Task.CompletedTask;
    }

    private void OnSignal(object sender, FileSystemEventArgs e)
    {
        if (Interlocked.CompareExchange(ref _handled, 1, 0) != 0) return;
        try
        {
            var content = File.ReadAllText(SignalFile).Trim();
            if (!string.Equals(content, _key, StringComparison.OrdinalIgnoreCase))
            {
                _log.LogWarning("[Shutdown] Signal GUID mismatch — ignoring");
                Interlocked.Exchange(ref _handled, 0);
                return;
            }
            _log.LogInformation("[Shutdown] Valid signal — stopping proxy");
            _lifetime.StopApplication();
        }
        catch (Exception ex)
        {
            _log.LogWarning(ex, "[Shutdown] Error reading signal file");
            Interlocked.Exchange(ref _handled, 0);
        }
    }

    private static string EnsureKey()
    {
        if (File.Exists(KeyFile))
            return File.ReadAllText(KeyFile).Trim();
        var key = Guid.NewGuid().ToString("N");
        try   { File.WriteAllText(KeyFile, key); }
        catch { key = File.ReadAllText(KeyFile).Trim(); } // lost race — use winner's key
        return key;
    }

    public void Dispose() => _watcher?.Dispose();
}
