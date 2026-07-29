using System.Collections.Concurrent;
using Microsoft.Extensions.Options;

namespace OfficeConversion.Conversion;

internal sealed class OfficeConversionQueue : IDisposable
{
    private readonly BlockingCollection<ConversionJob> _jobs = new();
    private readonly SemaphoreSlim _availableSlots;

    public OfficeConversionQueue(IOptions<ConversionOptions> options)
    {
        _availableSlots = new SemaphoreSlim(
            options.Value.QueueCapacity,
            options.Value.QueueCapacity);
    }

    public IEnumerable<ConversionJob> GetConsumingEnumerable() =>
        _jobs.GetConsumingEnumerable();

    public async Task EnqueueAsync(
        ConversionJob job,
        CancellationToken cancellationToken)
    {
        await _availableSlots.WaitAsync(cancellationToken);

        if (!_jobs.TryAdd(job))
        {
            _availableSlots.Release();
            throw new InvalidOperationException("The conversion queue is stopping.");
        }
    }

    public void MarkDequeued() => _availableSlots.Release();

    public void CompleteAdding() => _jobs.CompleteAdding();

    public void Dispose()
    {
        _jobs.Dispose();
        _availableSlots.Dispose();
    }
}
