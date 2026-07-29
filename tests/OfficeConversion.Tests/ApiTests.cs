using System.Net;
using System.Net.Http.Headers;
using Microsoft.AspNetCore.Hosting;
using Microsoft.AspNetCore.Mvc.Testing;
using Microsoft.Extensions.DependencyInjection;
using Microsoft.Extensions.DependencyInjection.Extensions;
using OfficeConversion.Conversion;
using Xunit;

namespace OfficeConversion.Tests;

public sealed class ApiTests : IClassFixture<OfficeConversionApplicationFactory>
{
    private readonly OfficeConversionApplicationFactory _factory;
    private readonly HttpClient _client;

    public ApiTests(OfficeConversionApplicationFactory factory)
    {
        _factory = factory;
        _client = factory.CreateClient();
    }

    [Fact]
    public async Task Health_ReturnsHealthy()
    {
        var response = await _client.GetAsync("/health");

        Assert.Equal(HttpStatusCode.OK, response.StatusCode);
        Assert.Contains("\"healthy\"", await response.Content.ReadAsStringAsync());
    }

    [Fact]
    public async Task Pdf_WithMultipartFile_ReturnsConvertedDownload()
    {
        using var request = CreateFileRequest("/api/pdf");

        var response = await _client.SendAsync(request);

        Assert.Equal(HttpStatusCode.OK, response.StatusCode);
        Assert.Equal("application/pdf", response.Content.Headers.ContentType?.MediaType);
        Assert.Equal(
            "wordtopdf.pdf",
            response.Content.Headers.ContentDisposition?.FileNameStar);
        Assert.Equal(
            FakeConversionService.ConvertedContent,
            await response.Content.ReadAsByteArrayAsync());
        Assert.Equal(ConversionTarget.WordPdf, _factory.Converter.LastTarget);
    }

    [Fact]
    public async Task Pdf_WithoutMultipartForm_ReturnsUnsupportedMediaType()
    {
        var response = await _client.PostAsync(
            "/api/pdf",
            new StringContent("not multipart"));

        Assert.Equal(HttpStatusCode.UnsupportedMediaType, response.StatusCode);
    }

    [Theory]
    [InlineData("/api/word/pdf", ConversionTarget.WordPdf, "application/pdf")]
    [InlineData(
        "/api/word/odt",
        ConversionTarget.WordOdt,
        "application/vnd.oasis.opendocument.text")]
    [InlineData("/api/excel/pdf", ConversionTarget.ExcelPdf, "application/pdf")]
    [InlineData(
        "/api/excel/ods",
        ConversionTarget.ExcelOds,
        "application/vnd.oasis.opendocument.spreadsheet")]
    public async Task SupportedRoutes_MapToExpectedConversion(
        string path,
        ConversionTarget expectedTarget,
        string expectedContentType)
    {
        using var request = CreateFileRequest(path);

        var response = await _client.SendAsync(request);

        Assert.Equal(HttpStatusCode.OK, response.StatusCode);
        Assert.Equal(expectedContentType, response.Content.Headers.ContentType?.MediaType);
        Assert.Equal(expectedTarget, _factory.Converter.LastTarget);
    }

    [Theory]
    [InlineData("/api/word/docx")]
    [InlineData("/api/excel/xlsx")]
    public async Task UnsupportedOutputType_ReturnsNotFound(string path)
    {
        using var request = CreateFileRequest(path);

        var response = await _client.SendAsync(request);

        Assert.Equal(HttpStatusCode.NotFound, response.StatusCode);
    }

    private static HttpRequestMessage CreateFileRequest(string path)
    {
        var multipart = new MultipartFormDataContent();
        var file = new ByteArrayContent([1, 2, 3]);
        file.Headers.ContentType =
            new MediaTypeHeaderValue("application/octet-stream");
        multipart.Add(file, "anything", "sample.docx");

        return new HttpRequestMessage(HttpMethod.Post, path)
        {
            Content = multipart
        };
    }
}

public sealed class OfficeConversionApplicationFactory :
    WebApplicationFactory<Program>
{
    public FakeConversionService Converter { get; } = new();

    protected override void ConfigureWebHost(IWebHostBuilder builder)
    {
        builder.ConfigureServices(services =>
        {
            services.RemoveAll<IConversionService>();
            services.AddSingleton<IConversionService>(Converter);
        });
    }
}

public sealed class FakeConversionService : IConversionService
{
    public static readonly byte[] ConvertedContent = [4, 5, 6];

    public ConversionTarget? LastTarget { get; private set; }

    public Task<byte[]> ConvertAsync(
        Stream input,
        string inputFileName,
        ConversionTarget target,
        CancellationToken cancellationToken)
    {
        LastTarget = target;
        return Task.FromResult(ConvertedContent);
    }
}
