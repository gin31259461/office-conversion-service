using System.Net;
using System.Net.Http.Headers;
using Microsoft.AspNetCore.Mvc.Testing;
using Xunit;

namespace OfficeConversion.Tests;

public sealed class OfficeSmokeTests
{
    [OfficeFact]
    public async Task Word_ConvertsRtfToPdf()
    {
        await using var factory = new WebApplicationFactory<Program>();
        using var client = factory.CreateClient();
        await using var fileStream = File.OpenRead(
            Path.Combine(AppContext.BaseDirectory, "Fixtures", "smoke.rtf"));

        using var multipart = new MultipartFormDataContent();
        using var file = new StreamContent(fileStream);
        file.Headers.ContentType = new MediaTypeHeaderValue("application/rtf");
        multipart.Add(file, "file", "smoke.rtf");

        var response = await client.PostAsync("/api/pdf", multipart);

        var responseBody = await response.Content.ReadAsByteArrayAsync();
        Assert.True(
            response.StatusCode == HttpStatusCode.OK,
            $"Expected 200 but received {(int)response.StatusCode}: " +
            System.Text.Encoding.UTF8.GetString(responseBody));
        var pdf = responseBody;
        Assert.True(pdf.Length > 4);
        Assert.Equal("%PDF", System.Text.Encoding.ASCII.GetString(pdf, 0, 4));
    }

    [OfficeFact]
    public async Task Excel_ConvertsCsvToPdf()
    {
        await using var factory = new WebApplicationFactory<Program>();
        using var client = factory.CreateClient();
        await using var fileStream = File.OpenRead(
            Path.Combine(AppContext.BaseDirectory, "Fixtures", "smoke.csv"));

        using var multipart = new MultipartFormDataContent();
        using var file = new StreamContent(fileStream);
        file.Headers.ContentType = new MediaTypeHeaderValue("text/csv");
        multipart.Add(file, "file", "smoke.csv");

        var response = await client.PostAsync("/api/excel/pdf", multipart);

        var responseBody = await response.Content.ReadAsByteArrayAsync();
        Assert.True(
            response.StatusCode == HttpStatusCode.OK,
            $"Expected 200 but received {(int)response.StatusCode}: " +
            System.Text.Encoding.UTF8.GetString(responseBody));
        Assert.True(responseBody.Length > 4);
        Assert.Equal(
            "%PDF",
            System.Text.Encoding.ASCII.GetString(responseBody, 0, 4));
    }
}

public sealed class OfficeFactAttribute : FactAttribute
{
    public OfficeFactAttribute()
    {
        if (!string.Equals(
                Environment.GetEnvironmentVariable(
                    "OFFICECONVERSION_RUN_OFFICE_TESTS"),
                "1",
                StringComparison.Ordinal))
        {
            Skip =
                "Set OFFICECONVERSION_RUN_OFFICE_TESTS=1 to run Microsoft Office tests.";
        }
    }
}
