namespace OfficeConversion.Conversion;

internal sealed class ConversionService(
    OfficeConversionQueue queue,
    ILogger<ConversionService> logger) : IConversionService
{
    public async Task<byte[]> ConvertAsync(
        Stream input,
        string inputFileName,
        ConversionTarget target,
        CancellationToken cancellationToken)
    {
        var workDirectory = Path.Combine(
            Path.GetTempPath(),
            "office-conversion",
            Guid.NewGuid().ToString("N"));

        Directory.CreateDirectory(workDirectory);

        var inputExtension = GetSafeExtension(inputFileName);
        var inputPath = Path.Combine(workDirectory, $"input{inputExtension}");
        var outputPath = Path.Combine(workDirectory, $"output{target.OutputExtension()}");

        try
        {
            await using (var file = new FileStream(
                inputPath,
                FileMode.CreateNew,
                FileAccess.Write,
                FileShare.None,
                bufferSize: 81920,
                useAsync: true))
            {
                await input.CopyToAsync(file, cancellationToken);
            }

            var job = new ConversionJob(
                inputPath,
                outputPath,
                target,
                cancellationToken);

            await queue.EnqueueAsync(job, cancellationToken);
            await job.Completion.Task;

            return await File.ReadAllBytesAsync(outputPath, CancellationToken.None);
        }
        finally
        {
            try
            {
                if (Directory.Exists(workDirectory))
                {
                    Directory.Delete(workDirectory, recursive: true);
                }
            }
            catch (Exception exception)
            {
                logger.LogWarning(
                    exception,
                    "Unable to remove conversion directory {WorkDirectory}",
                    workDirectory);
            }
        }
    }

    private static string GetSafeExtension(string fileName)
    {
        var extension = Path.GetExtension(Path.GetFileName(fileName));
        if (string.IsNullOrWhiteSpace(extension) ||
            extension.Length > 10 ||
            extension.Any(character => !char.IsLetterOrDigit(character) && character != '.'))
        {
            return ".tmp";
        }

        return extension.ToLowerInvariant();
    }
}
