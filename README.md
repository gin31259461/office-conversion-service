# Office Conversion Service

Windows document-conversion API that uses the locally installed Microsoft Word
and Excel rendering engines. The same application can run as a console app, a
Windows Service, or behind IIS.

## Requirements

- Windows 10/11 or Windows Server
- .NET 10 SDK for building
- .NET 10 ASP.NET Core Runtime for framework-dependent deployment
- Microsoft Word and Excel installed and activated
- The Office fonts and language packs required by the source documents

> [!IMPORTANT]
> Microsoft does not support unattended Office automation in IIS or Windows
> Services. This project serializes Office work, applies a timeout, and isolates
> each Office process as safeguards, but deployments must still monitor for
> Office dialogs, hangs, and profile-specific activation issues.

## Architecture

```text
HTTP request
    │
ASP.NET Core controller
    │
bounded conversion queue
    │
single STA worker
    │
installed Word or Excel
    │
download response
```

Only one Office conversion is processed at a time. This avoids concurrent COM
automation against the same service account. Uploaded and generated files are
stored in a unique system temporary directory and removed after the request.

## Configuration

`appsettings.json` contains only conversion behavior:

```json
{
  "Conversion": {
    "QueueCapacity": 20,
    "TimeoutSeconds": 120
  }
}
```

ASP.NET Core configuration overrides are supported. For example:

```powershell
$env:Conversion__TimeoutSeconds = "180"
$env:ASPNETCORE_URLS = "http://127.0.0.1:5057"
```

`Properties/launchSettings.json` is used only for local development.

## Run as a console app

```powershell
dotnet run
```

The development profile listens on `http://localhost:5057`. Verify it with:

```powershell
Invoke-RestMethod http://localhost:5057/health
```

## Run as a Windows Service

Publish the executable:

```powershell
dotnet publish -c Release -r win-x64 --self-contained false -o publish
```

Copy `publish` to a permanent directory, then register it from an elevated
PowerShell session:

```powershell
$binaryPath = '"C:\Services\OfficeConversion\OfficeConversion.Host.exe" --urls http://127.0.0.1:5057'

New-Service `
  -Name "OfficeConversion" `
  -DisplayName "Office Conversion Service" `
  -BinaryPathName $binaryPath `
  -StartupType Automatic

Start-Service OfficeConversion
```

Use a dedicated service account that has Office activated and access to all
required fonts. Do not use the same account for interactive Office work.

## Run behind IIS

Install the .NET 10 Hosting Bundle, then publish with the IIS out-of-process
hosting model:

```powershell
dotnet publish `
  -c Release `
  -r win-x64 `
  --self-contained false `
  -p:AspNetCoreHostingModel=OutOfProcess `
  -o publish-iis
```

Create an IIS site whose physical path is `publish-iis`, and configure its
application pool with:

- **.NET CLR version:** No Managed Code
- **Managed pipeline:** Integrated
- **Enable 32-bit applications:** False
- A dedicated application-pool identity with an initialized Office profile

For production environments with long-running conversions, prefer running
OfficeConversion as a Windows Service and placing IIS in front of that service as a
reverse proxy. IIS application-pool recycling can otherwise interrupt active
Office conversions.

## API

All conversion endpoints accept `multipart/form-data`. The first uploaded file
is converted, so the form field name is not significant.

| Endpoint | Output |
| --- | --- |
| `POST /api/pdf` | Word to PDF |
| `POST /api/word/pdf` | Word to PDF |
| `POST /api/word/odt` | Word to ODT |
| `POST /api/excel/pdf` | Excel to PDF |
| `POST /api/excel/ods` | Excel to ODS |
| `GET /health` | Host health |

Example:

```powershell
curl.exe `
  -X POST `
  -F "file=@C:\Documents\sample.docx" `
  -o sample.pdf `
  http://localhost:5057/api/pdf
```

Invalid requests return standard HTTP errors. Office failures return `500`, and
conversion timeouts return `504`.

## Tests

Run the normal test suite:

```powershell
dotnet test OfficeConversion.sln
```

The normal suite uses an in-memory converter and also starts the compiled
console executable to verify Kestrel hosting. To run the real Microsoft Word
smoke test:

```powershell
$env:OFFICECONVERSION_RUN_OFFICE_TESTS = "1"
dotnet test tests\OfficeConversion.Tests\OfficeConversion.Tests.csproj `
  --filter "FullyQualifiedName~OfficeSmokeTests"
```
