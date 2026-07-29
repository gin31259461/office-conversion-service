namespace OfficeConversion.Conversion;

internal sealed class OfficeConversionWorker(
    OfficeConversionQueue queue,
    IOfficeDocumentConverter converter,
    ILogger<OfficeConversionWorker> logger) : IHostedService
{
    private readonly TaskCompletionSource _stopped =
        new(TaskCreationOptions.RunContinuationsAsynchronously);

    private Thread? _thread;

    public Task StartAsync(CancellationToken cancellationToken)
    {
        _thread = new Thread(ProcessQueue)
        {
            IsBackground = true,
            Name = "OfficeConversion STA Worker"
        };
        _thread.SetApartmentState(ApartmentState.STA);
        _thread.Start();

        return Task.CompletedTask;
    }

    public async Task StopAsync(CancellationToken cancellationToken)
    {
        queue.CompleteAdding();
        await _stopped.Task.WaitAsync(cancellationToken);
    }

    private void ProcessQueue()
    {
        try
        {
            foreach (var job in queue.GetConsumingEnumerable())
            {
                queue.MarkDequeued();

                if (job.CancellationToken.IsCancellationRequested)
                {
                    job.Completion.TrySetCanceled(job.CancellationToken);
                    continue;
                }

                try
                {
                    converter.Convert(job.InputPath, job.OutputPath, job.Target);
                    job.Completion.TrySetResult();
                }
                catch (Exception exception)
                {
                    logger.LogError(
                        exception,
                        "Office conversion to {Target} failed",
                        job.Target);
                    job.Completion.TrySetException(exception);
                }
            }
        }
        catch (Exception exception)
        {
            logger.LogCritical(exception, "The Office conversion worker stopped unexpectedly");
        }
        finally
        {
            _stopped.TrySetResult();
        }
    }
}
