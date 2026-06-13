namespace OutlookShredder.Proxy.Services;

/// <summary>
/// In-memory relay shared between SteveController (HTTP relay) and FileWatcherService (CSV detection).
/// Uses plain statics + Volatile.Read/Write — the same pattern as PhoneSearchController.
/// </summary>
public static class SteveState
{
    private static string? _pendingTask;
    private static string? _exportResultPath;

    public static void    SetPending(string task)  => Volatile.Write(ref _pendingTask, task);
    public static string? GetPending()             => Volatile.Read(ref _pendingTask);
    public static void    ClearPending()           => Volatile.Write(ref _pendingTask, null);

    public static void    SetExportResult(string path) => Volatile.Write(ref _exportResultPath, path);
    public static string? GetExportResult()            => Volatile.Read(ref _exportResultPath);
    public static void    ClearExportResult()          => Volatile.Write(ref _exportResultPath, null);
}
