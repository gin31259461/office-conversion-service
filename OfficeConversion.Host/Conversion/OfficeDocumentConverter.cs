using System.Runtime.InteropServices;
using Microsoft.Extensions.Options;

namespace OfficeConversion.Conversion;

internal sealed class OfficeDocumentConverter(
    IOptions<ConversionOptions> options,
    ILogger<OfficeDocumentConverter> logger) : IOfficeDocumentConverter
{
    private const int DoNotSaveChanges = 0;
    private const int WordExportPdf = 17;
    private const int WordOpenDocumentText = 23;
    private const int ExcelFixedFormatPdf = 0;
    private const int ExcelOpenDocumentSpreadsheet = 60;

    private readonly TimeSpan _timeout =
        TimeSpan.FromSeconds(options.Value.TimeoutSeconds);

    public void Convert(
        string inputPath,
        string outputPath,
        ConversionTarget target)
    {
        switch (target)
        {
            case ConversionTarget.WordPdf:
            case ConversionTarget.WordOdt:
                ConvertWord(inputPath, outputPath, target);
                break;
            case ConversionTarget.ExcelPdf:
            case ConversionTarget.ExcelOds:
                ConvertExcel(inputPath, outputPath, target);
                break;
            default:
                throw new ArgumentOutOfRangeException(nameof(target), target, null);
        }
    }

    private void ConvertWord(
        string inputPath,
        string outputPath,
        ConversionTarget target)
    {
        object? application = null;
        object? documents = null;
        object? document = null;
        using var watchdog = new OfficeProcessWatchdog(
            _timeout,
            "WINWORD",
            logger);

        try
        {
            var applicationType =
                Type.GetTypeFromProgID("Word.Application", throwOnError: true)!;
            application = Activator.CreateInstance(applicationType)
                ?? throw new InvalidOperationException("Unable to create Microsoft Word.");

            dynamic word = application;
            word.Visible = false;
            word.DisplayAlerts = 0;
            watchdog.TrackNewProcess();

            documents = word.Documents;
            dynamic wordDocuments = documents;
            document = wordDocuments.Open(
                FileName: inputPath,
                ReadOnly: true,
                AddToRecentFiles: false,
                Visible: false,
                OpenAndRepair: true,
                NoEncodingDialog: true);

            dynamic wordDocument = document;
            if (target == ConversionTarget.WordPdf)
            {
                wordDocument.ExportAsFixedFormat(outputPath, WordExportPdf);
            }
            else
            {
                wordDocument.SaveAs2(outputPath, WordOpenDocumentText);
            }

            watchdog.ThrowIfTimedOut();
        }
        catch (Exception exception) when (watchdog.HasTimedOut)
        {
            throw new TimeoutException(
                "Microsoft Word conversion timed out.",
                exception);
        }
        finally
        {
            CloseWordDocument(document);
            ReleaseComObject(documents);
            QuitWordApplication(application);
        }
    }

    private void ConvertExcel(
        string inputPath,
        string outputPath,
        ConversionTarget target)
    {
        object? application = null;
        object? workbooks = null;
        object? workbook = null;
        using var watchdog = new OfficeProcessWatchdog(
            _timeout,
            "EXCEL",
            logger);

        try
        {
            var applicationType =
                Type.GetTypeFromProgID("Excel.Application", throwOnError: true)!;
            application = Activator.CreateInstance(applicationType)
                ?? throw new InvalidOperationException("Unable to create Microsoft Excel.");

            dynamic excel = application;
            excel.Visible = false;
            excel.DisplayAlerts = false;
            watchdog.TrackNewProcess();

            workbooks = excel.Workbooks;
            dynamic excelWorkbooks = workbooks;
            workbook = excelWorkbooks.Open(
                Filename: inputPath,
                ReadOnly: true,
                AddToMru: false,
                IgnoreReadOnlyRecommended: true);

            dynamic excelWorkbook = workbook;
            if (target == ConversionTarget.ExcelPdf)
            {
                excelWorkbook.ExportAsFixedFormat(
                    ExcelFixedFormatPdf,
                    outputPath);
            }
            else
            {
                excelWorkbook.SaveAs(
                    outputPath,
                    ExcelOpenDocumentSpreadsheet);
            }

            watchdog.ThrowIfTimedOut();
        }
        catch (Exception exception) when (watchdog.HasTimedOut)
        {
            throw new TimeoutException(
                "Microsoft Excel conversion timed out.",
                exception);
        }
        finally
        {
            CloseExcelWorkbook(workbook);
            ReleaseComObject(workbooks);
            QuitExcelApplication(application);
        }
    }

    private static void CloseWordDocument(object? document)
    {
        if (document is null)
        {
            return;
        }

        try
        {
            dynamic wordDocument = document;
            wordDocument.Close(DoNotSaveChanges);
        }
        catch (COMException)
        {
            // The watchdog may already have terminated Word.
        }
        finally
        {
            ReleaseComObject(document);
        }
    }

    private static void QuitWordApplication(object? application)
    {
        if (application is null)
        {
            return;
        }

        try
        {
            dynamic word = application;
            word.Quit(DoNotSaveChanges);
        }
        catch (COMException)
        {
            // The watchdog may already have terminated Word.
        }
        finally
        {
            ReleaseComObject(application);
        }
    }

    private static void CloseExcelWorkbook(object? workbook)
    {
        if (workbook is null)
        {
            return;
        }

        try
        {
            dynamic excelWorkbook = workbook;
            excelWorkbook.Close(SaveChanges: false);
        }
        catch (COMException)
        {
            // The watchdog may already have terminated Excel.
        }
        finally
        {
            ReleaseComObject(workbook);
        }
    }

    private static void QuitExcelApplication(object? application)
    {
        if (application is null)
        {
            return;
        }

        try
        {
            dynamic excel = application;
            excel.Quit();
        }
        catch (COMException)
        {
            // The watchdog may already have terminated Excel.
        }
        finally
        {
            ReleaseComObject(application);
        }
    }

    private static void ReleaseComObject(object? value)
    {
        if (value is not null && Marshal.IsComObject(value))
        {
            Marshal.FinalReleaseComObject(value);
        }
    }
}
