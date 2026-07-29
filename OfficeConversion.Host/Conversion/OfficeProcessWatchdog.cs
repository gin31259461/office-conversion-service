using System.Diagnostics;

namespace OfficeConversion.Conversion;

internal sealed class OfficeProcessWatchdog : IDisposable
{
    private readonly ILogger _logger;
    private readonly string _processName;
    private readonly HashSet<int> _existingProcessIds;
    private readonly Timer _timer;
    private int _processId;
    private int _timedOut;

    public OfficeProcessWatchdog(
        TimeSpan timeout,
        string processName,
        ILogger logger)
    {
        _logger = logger;
        _processName = processName;
        _existingProcessIds = GetProcessIds(processName);
        _timer = new Timer(OnTimeout, null, timeout, Timeout.InfiniteTimeSpan);
    }

    public bool HasTimedOut => Volatile.Read(ref _timedOut) != 0;

    public void TrackNewProcess()
    {
        var processId = FindNewProcessId();
        if (processId != 0)
        {
            Volatile.Write(ref _processId, processId);
        }
    }

    public void ThrowIfTimedOut()
    {
        if (HasTimedOut)
        {
            throw new TimeoutException("Microsoft Office conversion timed out.");
        }
    }

    public void Dispose() => _timer.Dispose();

    private void OnTimeout(object? state)
    {
        Interlocked.Exchange(ref _timedOut, 1);

        var processId = Volatile.Read(ref _processId);
        if (processId == 0)
        {
            processId = FindNewProcessId();
        }

        if (processId == 0)
        {
            _logger.LogError(
                "Office conversion timed out before its {ProcessName} process could be identified",
                _processName);
            return;
        }

        try
        {
            using var process = Process.GetProcessById(processId);
            _logger.LogError(
                "Office conversion timed out; terminating process {ProcessId}",
                processId);
            process.Kill(entireProcessTree: true);
        }
        catch (Exception exception)
        {
            _logger.LogError(
                exception,
                "Unable to terminate timed-out Office process {ProcessId}",
                processId);
        }
    }

    private int FindNewProcessId()
    {
        return GetProcessIds(_processName)
            .Where(processId => !_existingProcessIds.Contains(processId))
            .OrderByDescending(processId => processId)
            .FirstOrDefault();
    }

    private static HashSet<int> GetProcessIds(string processName)
    {
        var processes = Process.GetProcessesByName(processName);
        try
        {
            return processes.Select(process => process.Id).ToHashSet();
        }
        finally
        {
            foreach (var process in processes)
            {
                process.Dispose();
            }
        }
    }
}
