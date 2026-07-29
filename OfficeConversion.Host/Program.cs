using OfficeConversion.Conversion;

var builder = WebApplication.CreateBuilder(args);

var configuredUrls = builder.Configuration[WebHostDefaults.ServerUrlsKey];
if (string.IsNullOrWhiteSpace(configuredUrls))
{
    var appSettingsUrls = builder.Configuration["Hosting:Urls"];
    builder.WebHost.UseUrls(
        string.IsNullOrWhiteSpace(appSettingsUrls)
            ? "http://localhost:8085"
            : appSettingsUrls);
}

builder.Services.AddWindowsService(options => options.ServiceName = "OfficeConversion");
builder.Services.AddControllers();
builder.Services
    .AddOptions<ConversionOptions>()
    .Bind(builder.Configuration.GetSection(ConversionOptions.SectionName))
    .ValidateDataAnnotations()
    .ValidateOnStart();

builder.Services.AddSingleton<OfficeConversionQueue>();
builder.Services.AddSingleton<IOfficeDocumentConverter, OfficeDocumentConverter>();
builder.Services.AddSingleton<IConversionService, ConversionService>();
builder.Services.AddHostedService<OfficeConversionWorker>();

var app = builder.Build();

app.MapControllers();
app.MapGet("/health", () => Results.Ok(new { status = "healthy" }));

app.Run();

public partial class Program;
