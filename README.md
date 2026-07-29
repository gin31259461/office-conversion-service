# Office Conversion Service

A Windows-hosted HTTP API that converts documents with the locally installed
Microsoft Word and Excel rendering engines. It is intended for conversions
where Microsoft Office layout fidelity matters more than cross-platform
execution.

The same application can run as:

- a console application backed by Kestrel;
- a Windows Service;
- an ASP.NET Core application behind IIS.

## Features

- Word to PDF
- Word to OpenDocument Text (`.odt`)
- Excel to PDF
- Excel to OpenDocument Spreadsheet (`.ods`)
- Bounded request queue with configurable capacity
- A single dedicated STA worker for Office COM automation
- Per-conversion timeout with an Office process watchdog
- Isolated temporary directories with best-effort cleanup
- Health endpoint and integration tests
- Optional real Word and Excel smoke tests

> [!IMPORTANT]
> Microsoft does not support unattended Office automation from ASP.NET, IIS, or
> Windows Services. This project adds serialization, timeouts, process tracking,
> and cleanup to reduce operational risk, but it cannot make unattended Office
> automation fully supported or eliminate every dialog, hang, or activation
> issue.

> [!WARNING]
> The API currently has no authentication, authorization, file extension
> allow-list, explicit upload-size policy, or application-level rate limiting.
> Do not expose it directly to an untrusted network without adding those
> controls at the application or reverse-proxy layer.

## How it works

```text
multipart/form-data request
          │
          ▼
ASP.NET Core controller
          │
          ▼
unique temporary work directory
          │
          ▼
bounded conversion queue
          │
          ▼
single dedicated STA thread
          │
          ▼
Word.Application or Excel.Application
          │
          ▼
download response + temporary-file cleanup
```

Office is created through Windows COM late binding. The project therefore does
not need legacy Office PIA NuGet packages, but Word or Excel must still be
installed, registered, and activated for the account running the process.

Only one conversion runs at a time. This is deliberate: Word and Excel are
desktop applications with shared state and are unreliable when automated
concurrently under one service identity.

## Requirements

- Windows 10/11 or Windows Server
- .NET 10 SDK for development and building
- .NET 10 ASP.NET Core Runtime for framework-dependent deployment
- Microsoft Word and Excel installed and activated
- Required document fonts and Office language packs
- For IIS: IIS plus the .NET 10 Hosting Bundle

The application targets `net10.0-windows` and is not intended to run on Linux
or macOS.

## Quick start

From the repository root:

```powershell
dotnet restore .\OfficeConversion.sln
dotnet run --project .\OfficeConversion.Host\OfficeConversion.Host.csproj
```

By default, the host listens on the URL configured by `Hosting:Urls` in
`OfficeConversion.Host/appsettings.json`:

```text
http://localhost:8085
```

Check the host:

```powershell
Invoke-RestMethod http://localhost:8085/health
```

Expected response:

```json
{
  "status": "healthy"
}
```

Convert a Word document:

```powershell
curl.exe `
  --fail-with-body `
  -X POST `
  -F "file=@C:\Documents\sample.docx" `
  -o sample.pdf `
  http://localhost:8085/api/pdf
```

The multipart field name is not significant; the first uploaded file is used.

## API

| Method | Route | Input | Output |
| --- | --- | --- | --- |
| `GET` | `/health` | None | JSON health status |
| `POST` | `/api/pdf` | Word-readable document | PDF |
| `POST` | `/api/word/pdf` | Word-readable document | PDF |
| `POST` | `/api/word/odt` | Word-readable document | ODT |
| `POST` | `/api/excel/pdf` | Excel-readable workbook | PDF |
| `POST` | `/api/excel/ods` | Excel-readable workbook | ODS |

Common Word inputs include `.doc`, `.docx`, and `.rtf`. Common Excel inputs
include `.xls`, `.xlsx`, and `.csv`. Actual compatibility is determined by the
installed Office version and its file filters.

### Responses

Successful conversions return a file download with the appropriate media type.

| Status | Meaning |
| --- | --- |
| `200` | Conversion succeeded |
| `400` | The multipart request contains no non-empty file |
| `404` | The requested output type is unsupported |
| `415` | The request is not `multipart/form-data` |
| `499` | The client canceled the request |
| `500` | Microsoft Office conversion failed |
| `504` | Microsoft Office conversion exceeded the configured timeout |

## Configuration

The default configuration is in
[`OfficeConversion.Host/appsettings.json`](OfficeConversion.Host/appsettings.json):

```json
{
  "Hosting": {
    "Urls": "http://localhost:8085"
  },
  "Conversion": {
    "QueueCapacity": 20,
    "TimeoutSeconds": 120
  }
}
```

| Setting | Valid range | Purpose |
| --- | --- | --- |
| `Hosting:Urls` | One or more semicolon-separated HTTP(S) URLs | Default addresses on which the standalone Kestrel host listens |
| `Conversion:QueueCapacity` | `1`–`1000` | Maximum number of jobs waiting to be processed |
| `Conversion:TimeoutSeconds` | `10`–`3600` | Maximum Office processing time before the watchdog terminates the tracked process |

ASP.NET Core configuration providers can override JSON settings:

```powershell
$env:Conversion__QueueCapacity = "10"
$env:Conversion__TimeoutSeconds = "180"
$env:ASPNETCORE_URLS = "http://127.0.0.1:5057"
```

`ASPNETCORE_URLS` and the `--urls` command-line option override
`Hosting:Urls`. If none of these settings is present, the application falls
back to `http://localhost:8085`.
[`launchSettings.json`](OfficeConversion.Host/Properties/launchSettings.json)
is for local development only. IIS supplies its own binding when hosting the
application.

Temporary work is stored under:

```text
<system temp>\office-conversion\<job-guid>\
```

The output is currently loaded into memory before it is returned. Consider
streaming and explicit request-size limits before using this service for very
large documents.

## Run as a console application

Development:

```powershell
dotnet run --project .\OfficeConversion.Host\OfficeConversion.Host.csproj
```

Published executable:

```powershell
dotnet publish `
  .\OfficeConversion.Host\OfficeConversion.Host.csproj `
  -c Release `
  -r win-x64 `
  --self-contained false `
  -o .\OfficeConversion.Host\bin\publish\win-x64

.\OfficeConversion.Host\bin\publish\win-x64\OfficeConversion.Host.exe `
  --urls http://127.0.0.1:5057
```

## Run as a Windows Service

Publish:

```powershell
dotnet publish `
  .\OfficeConversion.Host\OfficeConversion.Host.csproj `
  -c Release `
  -r win-x64 `
  --self-contained false `
  -o .\OfficeConversion.Host\bin\publish\service
```

Copy the publish directory to a permanent location such as:

```text
C:\Services\OfficeConversion
```

Use a dedicated Windows account. Sign in as that account at least once, activate
Word and Excel, dismiss first-run prompts, install required fonts, and verify a
manual conversion before registering the service.

From an elevated PowerShell session:

```powershell
$credential = Get-Credential
$binaryPath = '"C:\Services\OfficeConversion\OfficeConversion.Host.exe"'

New-Service `
  -Name "OfficeConversion" `
  -DisplayName "Office Conversion Service" `
  -Description "Converts Word and Excel documents through Microsoft Office." `
  -BinaryPathName $binaryPath `
  -Credential $credential `
  -StartupType Automatic

Start-Service OfficeConversion
Get-Service OfficeConversion
```

The service account needs:

- **Log on as a service** permission;
- read and execute permission on the deployment directory;
- write access to its temporary directory;
- a valid Office installation, activation, and user profile;
- access to every font required for faithful rendering.

## Run behind IIS

Install the .NET 10 Hosting Bundle, then publish using the out-of-process hosting
model:

```powershell
dotnet publish `
  .\OfficeConversion.Host\OfficeConversion.Host.csproj `
  -c Release `
  -r win-x64 `
  --self-contained false `
  -p:AspNetCoreHostingModel=OutOfProcess `
  -o .\OfficeConversion.Host\bin\publish\iis
```

Create an IIS site pointing to the publish directory. Configure a dedicated
application pool:

- **.NET CLR version:** No Managed Code
- **Managed pipeline:** Integrated
- **Enable 32-bit applications:** False
- **Identity:** a dedicated account with an initialized and activated Office
  profile

The publish command generates the required `web.config` and configures the
ASP.NET Core Module to launch `OfficeConversion.Host.exe`.

> [!TIP]
> For long-running production conversions, the more predictable arrangement is
> to run Office Conversion Service as a Windows Service and place IIS in front
> of it as a reverse proxy. IIS application-pool recycling can otherwise
> interrupt active work.

## Testing

Run the normal suite:

```powershell
dotnet test .\OfficeConversion.sln
```

The normal suite:

- replaces Office conversion with an in-memory fake for API contract tests;
- exercises all supported routes and media types;
- starts the compiled console executable and checks the Kestrel health endpoint;
- skips tests that require installed Office.

Run the real Word and Excel smoke tests:

```powershell
$env:OFFICECONVERSION_RUN_OFFICE_TESTS = "1"

dotnet test `
  .\tests\OfficeConversion.Tests\OfficeConversion.Tests.csproj `
  --filter "FullyQualifiedName~OfficeSmokeTests"
```

The smoke tests launch the installed Word and Excel applications and verify that
both produce a PDF. Run them only on a machine prepared for Office automation.

Verify formatting without changing files:

```powershell
dotnet format .\OfficeConversion.sln `
  --verify-no-changes `
  --no-restore
```

## Repository layout

```text
.
├── OfficeConversion.sln
├── OfficeConversion.Host
│   ├── Controllers
│   │   └── PDFController.cs
│   ├── Conversion
│   │   ├── ConversionService.cs
│   │   ├── OfficeConversionQueue.cs
│   │   ├── OfficeConversionWorker.cs
│   │   ├── OfficeDocumentConverter.cs
│   │   └── OfficeProcessWatchdog.cs
│   ├── Program.cs
│   └── appsettings.json
└── tests
    └── OfficeConversion.Tests
        ├── ApiTests.cs
        ├── ConsoleHostTests.cs
        ├── OfficeSmokeTests.cs
        └── Fixtures
```

## Operational notes

- Use a dedicated Windows identity for Office automation.
- Do not use that identity for interactive desktop Office work.
- Keep Word and Excel versions, fonts, language packs, and printer settings
  consistent between environments; they can affect document pagination.
- The worker is intentionally single-threaded. Scaling requires isolated
  service instances or machines, not additional worker threads in one process.
- Monitor application logs for timeout, process termination, COM, activation,
  and temporary-directory cleanup errors.
- A healthy `/health` response confirms that the web host is running; it does
  not launch Word or Excel or prove that Office conversion is healthy.

## Troubleshooting

### `Class not registered` or Office cannot be created

Confirm that Word and Excel are installed for the server architecture and that
their COM ProgIDs are registered:

```powershell
Get-ItemProperty Registry::HKEY_CLASSES_ROOT\Word.Application\CurVer
Get-ItemProperty Registry::HKEY_CLASSES_ROOT\Excel.Application\CurVer
```

### Conversion works interactively but fails as a service

Run the service under the same dedicated account used to activate and initialize
Office. Check profile creation, file-system permissions, fonts, language packs,
and any hidden first-run or recovery dialogs.

### Requests return `504`

Check logs for the tracked `WINWORD` or `EXCEL` process termination. Increase
`Conversion__TimeoutSeconds` only after verifying that the document legitimately
needs more time and is not blocked by a dialog.

### Layout differs between machines

Compare Office versions, update channels, fonts, language packs, default printer
drivers, and document-linked resources. Using the Microsoft Office rendering
engine improves fidelity but does not make different Windows environments
identical.
