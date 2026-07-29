namespace OfficeConversion.Conversion;

public enum ConversionTarget
{
    WordPdf,
    WordOdt,
    ExcelPdf,
    ExcelOds
}

public static class ConversionTargetExtensions
{
    public static string OutputExtension(this ConversionTarget target) => target switch
    {
        ConversionTarget.WordPdf or ConversionTarget.ExcelPdf => ".pdf",
        ConversionTarget.WordOdt => ".odt",
        ConversionTarget.ExcelOds => ".ods",
        _ => throw new ArgumentOutOfRangeException(nameof(target), target, null)
    };

    public static string ContentType(this ConversionTarget target) => target switch
    {
        ConversionTarget.WordPdf or ConversionTarget.ExcelPdf => "application/pdf",
        ConversionTarget.WordOdt => "application/vnd.oasis.opendocument.text",
        ConversionTarget.ExcelOds => "application/vnd.oasis.opendocument.spreadsheet",
        _ => throw new ArgumentOutOfRangeException(nameof(target), target, null)
    };

    public static string DownloadFileName(this ConversionTarget target) => target switch
    {
        ConversionTarget.WordPdf => "wordtopdf.pdf",
        ConversionTarget.WordOdt => "wordtoodt.odt",
        ConversionTarget.ExcelPdf => "exceltopdf.pdf",
        ConversionTarget.ExcelOds => "exceltoods.ods",
        _ => throw new ArgumentOutOfRangeException(nameof(target), target, null)
    };
}
