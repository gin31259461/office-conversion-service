using System.Diagnostics;
using System.Net;
using System.Net.Sockets;
using System.Text.Json;
using Xunit;

namespace OfficeConversion.Tests;

public sealed class ConsoleHostTests
{
    [Fact]
    public async Task Executable_StartsKestrelAndServesHealthEndpoint()
    {
        var port = GetAvailablePort();
        var executablePath = Path.Combine(
            AppContext.BaseDirectory,
            OperatingSystem.IsWindows()
                ? "OfficeConversion.Host.exe"
                : "OfficeConversion.Host");

        var startInfo = new ProcessStartInfo(executablePath)
        {
            UseShellExecute = false,
            CreateNoWindow = true
        };
        startInfo.Environment["ASPNETCORE_URLS"] = $"http://127.0.0.1:{port}";

        using var process = Process.Start(startInfo)
            ?? throw new InvalidOperationException(
                "Unable to start OfficeConversion.Host.");

        try
        {
            using var client = new HttpClient
            {
                Timeout = TimeSpan.FromMilliseconds(500)
            };

            HttpResponseMessage? response = null;
            for (var attempt = 0; attempt < 20; attempt++)
            {
                try
                {
                    response = await client.GetAsync($"http://127.0.0.1:{port}/health");
                    break;
                }
                catch (HttpRequestException)
                {
                    await Task.Delay(250);
                }
                catch (TaskCanceledException)
                {
                    await Task.Delay(250);
                }
            }

            Assert.NotNull(response);
            Assert.Equal(HttpStatusCode.OK, response.StatusCode);

            var body = JsonDocument.Parse(await response.Content.ReadAsStreamAsync());
            Assert.Equal("healthy", body.RootElement.GetProperty("status").GetString());
        }
        finally
        {
            if (!process.HasExited)
            {
                process.Kill(entireProcessTree: true);
                await process.WaitForExitAsync();
            }
        }
    }

    private static int GetAvailablePort()
    {
        var listener = new TcpListener(IPAddress.Loopback, 0);
        listener.Start();
        try
        {
            return ((IPEndPoint)listener.LocalEndpoint).Port;
        }
        finally
        {
            listener.Stop();
        }
    }
}
