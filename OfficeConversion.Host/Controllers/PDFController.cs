using Microsoft.AspNetCore.Mvc;
using OfficeConversion.Conversion;

namespace OfficeConversion.Controllers;

[ApiController]
public sealed class PDFController(
    IConversionService conversionService,
    ILogger<PDFController> logger) : ControllerBase
{
    [HttpPost("api/pdf")]
    public Task<IActionResult> ConvertWordToPdf(CancellationToken cancellationToken) =>
        ConvertUploadedFile(ConversionTarget.WordPdf, cancellationToken);

    [HttpPost("api/word/{type}")]
    public Task<IActionResult> ConvertWord(
        string type,
        CancellationToken cancellationToken)
    {
        var target = type.ToLowerInvariant() switch
        {
            "pdf" => ConversionTarget.WordPdf,
            "odt" => ConversionTarget.WordOdt,
            _ => (ConversionTarget?)null
        };

        return target is null
            ? Task.FromResult<IActionResult>(NotFound())
            : ConvertUploadedFile(target.Value, cancellationToken);
    }

    [HttpPost("api/excel/{type}")]
    public Task<IActionResult> ConvertExcel(
        string type,
        CancellationToken cancellationToken)
    {
        var target = type.ToLowerInvariant() switch
        {
            "pdf" => ConversionTarget.ExcelPdf,
            "ods" => ConversionTarget.ExcelOds,
            _ => (ConversionTarget?)null
        };

        return target is null
            ? Task.FromResult<IActionResult>(NotFound())
            : ConvertUploadedFile(target.Value, cancellationToken);
    }

    private async Task<IActionResult> ConvertUploadedFile(
        ConversionTarget target,
        CancellationToken cancellationToken)
    {
        if (!Request.HasFormContentType)
        {
            return StatusCode(
                StatusCodes.Status415UnsupportedMediaType,
                new ProblemDetails
                {
                    Status = StatusCodes.Status415UnsupportedMediaType,
                    Title = "multipart/form-data is required."
                });
        }

        var form = await Request.ReadFormAsync(cancellationToken);
        var uploadedFile = form.Files.FirstOrDefault();
        if (uploadedFile is null || uploadedFile.Length == 0)
        {
            return BadRequest(new ProblemDetails
            {
                Status = StatusCodes.Status400BadRequest,
                Title = "A non-empty file is required."
            });
        }

        try
        {
            await using var input = uploadedFile.OpenReadStream();
            var output = await conversionService.ConvertAsync(
                input,
                uploadedFile.FileName,
                target,
                cancellationToken);

            return File(
                output,
                target.ContentType(),
                target.DownloadFileName());
        }
        catch (OperationCanceledException) when (cancellationToken.IsCancellationRequested)
        {
            return StatusCode(499);
        }
        catch (TimeoutException exception)
        {
            logger.LogWarning(exception, "Conversion to {Target} timed out", target);
            return Problem(
                statusCode: StatusCodes.Status504GatewayTimeout,
                title: "Microsoft Office conversion timed out.");
        }
        catch (Exception exception)
        {
            logger.LogError(exception, "Conversion to {Target} failed", target);
            return Problem(
                statusCode: StatusCodes.Status500InternalServerError,
                title: "Microsoft Office conversion failed.");
        }
    }
}
