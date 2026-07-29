namespace OfficeConversion.Conversion;

public interface IConversionService
{
    Task<byte[]> ConvertAsync(
        Stream input,
        string inputFileName,
        ConversionTarget target,
        CancellationToken cancellationToken);
}
