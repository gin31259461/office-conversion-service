namespace OfficeConversion.Conversion;

internal sealed record ConversionJob(
    string InputPath,
    string OutputPath,
    ConversionTarget Target,
    CancellationToken CancellationToken)
{
    public TaskCompletionSource Completion { get; } =
        new(TaskCreationOptions.RunContinuationsAsynchronously);
}
