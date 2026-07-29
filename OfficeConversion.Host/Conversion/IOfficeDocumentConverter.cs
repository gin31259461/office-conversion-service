namespace OfficeConversion.Conversion;

internal interface IOfficeDocumentConverter
{
    void Convert(
        string inputPath,
        string outputPath,
        ConversionTarget target);
}
