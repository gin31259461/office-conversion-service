# AGENTS.md

## Project overview

Office Conversion Service is a Windows-only ASP.NET Core API that converts Word
and Excel files with the locally installed Microsoft Office rendering engines.
It targets `net10.0-windows` and can run as a console application, a Windows
Service, or behind IIS.

The repository contains two projects:

- `OfficeConversion.Host/OfficeConversion.Host.csproj`: the web host and Office
  conversion pipeline;
- `tests/OfficeConversion.Tests/OfficeConversion.Tests.csproj`: xUnit API,
  executable-host, and opt-in real Office tests.

Read `README.md` before changing hosting, deployment, API behavior, or Office
automation.

## Architecture and invariants

The request flow is:

```text
PDFController
  -> ConversionService
  -> OfficeConversionQueue
  -> OfficeConversionWorker (one dedicated STA thread)
  -> OfficeDocumentConverter
  -> Word.Application or Excel.Application
```

Preserve these invariants unless the task explicitly changes the architecture:

1. Controllers must not create or invoke Office COM objects directly.
2. Office work must pass through the bounded queue.
3. A single dedicated STA thread processes jobs serially. Do not replace this
   with parallel task-pool execution.
4. `OfficeDocumentConverter` uses late-bound COM ProgIDs. Do not add legacy
   Office PIA packages merely for enum names or early binding.
5. Release documents, workbook collections, and Office applications in reverse
   ownership order. Keep cleanup in `finally`.
6. The watchdog may terminate only the Office process identified as newly
   created for the current conversion. Never kill all `WINWORD` or `EXCEL`
   processes.
7. Each job uses a unique directory below
   `<system temp>/office-conversion/<guid>` and must clean it up on success,
   failure, and cancellation.
8. Keep the queue capacity and timeout validated at startup.
9. Existing API routes and download names are compatibility-sensitive. Do not
   rename them incidentally.
10. Keep `Program.cs` top-level statements and the global
    `public partial class Program` test seam unless a task explicitly requires a
    different startup pattern.

Microsoft does not support unattended Office automation in server processes.
Treat timeout, process cleanup, logging, and the dedicated identity guidance as
operational safety requirements, not optional polish.

## Repository layout

```text
OfficeConversion.sln
OfficeConversion.Host/
  Controllers/PDFController.cs
  Conversion/
  Program.cs
  appsettings.json
  Properties/launchSettings.json
tests/OfficeConversion.Tests/
  ApiTests.cs
  ConsoleHostTests.cs
  OfficeSmokeTests.cs
  Fixtures/
README.md
AGENTS.md
```

There are no generated source files. Never edit or commit `bin/`, `obj/`,
publish output, test results, Office output files, or temporary conversion
directories.

## Setup and development commands

Run all commands from the repository root.

Restore:

```powershell
dotnet restore .\OfficeConversion.sln
```

Build:

```powershell
dotnet build .\OfficeConversion.sln --no-restore
```

Run the development host:

```powershell
dotnet run --project .\OfficeConversion.Host\OfficeConversion.Host.csproj
```

The launch profile listens on `http://localhost:5057`. Check it with:

```powershell
Invoke-RestMethod http://localhost:5057/health
```

Configuration follows normal ASP.NET Core precedence. Environment-variable
names use double underscores:

```powershell
$env:Conversion__QueueCapacity = "20"
$env:Conversion__TimeoutSeconds = "120"
$env:ASPNETCORE_URLS = "http://127.0.0.1:5057"
```

`ConversionOptions` validation currently permits:

- `QueueCapacity`: `1` through `1000`;
- `TimeoutSeconds`: `10` through `3600`.

## Testing instructions

Run the default suite after every code change:

```powershell
dotnet test .\OfficeConversion.sln --no-restore
```

The default suite must not require Microsoft Office. It uses
`FakeConversionService` for API behavior and skips `OfficeSmokeTests`.

Run a focused test:

```powershell
dotnet test `
  .\tests\OfficeConversion.Tests\OfficeConversion.Tests.csproj `
  --no-restore `
  --filter "FullyQualifiedName~ApiTests"
```

Run real Office tests only on a prepared Windows machine:

```powershell
$env:OFFICECONVERSION_RUN_OFFICE_TESTS = "1"

dotnet test `
  .\tests\OfficeConversion.Tests\OfficeConversion.Tests.csproj `
  --no-restore `
  --filter "FullyQualifiedName~OfficeSmokeTests"
```

After real Office tests, verify that no test-created processes remain:

```powershell
Get-Process -Name OfficeConversion.Host,WINWORD,EXCEL -ErrorAction SilentlyContinue
```

Do not terminate a pre-existing interactive Word or Excel process. Investigate
the process owner and start time before stopping anything.

Test responsibilities:

- `ApiTests.cs`: route mapping, status, content type, download metadata, and
  request validation without Office;
- `ConsoleHostTests.cs`: starts the compiled executable, checks Kestrel, and
  terminates only that child process;
- `OfficeSmokeTests.cs`: opt-in Word and Excel rendering verification.

When adding a conversion target, update all of the following:

1. `ConversionTarget`;
2. extension, media type, and download-name mappings;
3. controller route mapping;
4. `OfficeDocumentConverter`;
5. fake API tests;
6. real smoke tests when a practical fixture exists;
7. the API table in `README.md`.

## Code style

- Use C# file-scoped namespaces.
- Keep nullable reference types and implicit usings enabled.
- Prefer primary constructors where dependencies are simple and immutable.
- Use `async` I/O for request and file operations.
- Keep COM calls synchronous on the STA worker.
- Use `CancellationToken` on request, queue, and file-I/O boundaries.
- Use structured `ILogger` messages with named properties.
- Keep public interfaces narrow; conversion implementation types should remain
  `internal` unless consumers need them.
- Prefer switch expressions for exhaustive `ConversionTarget` mappings.
- Do not expose raw exception details in HTTP responses. Log server details and
  return stable `ProblemDetails`.
- Preserve top-level `Program.cs`; move substantial logic into focused classes
  rather than expanding startup code.

Verify formatting:

```powershell
dotnet format .\OfficeConversion.sln `
  --verify-no-changes `
  --no-restore
```

Do not make unrelated formatting changes.

## API behavior

The supported routes are:

- `GET /health`
- `POST /api/pdf`
- `POST /api/word/pdf`
- `POST /api/word/odt`
- `POST /api/excel/pdf`
- `POST /api/excel/ods`

Conversion requests use `multipart/form-data`; the first uploaded file is used
and must be non-empty. Successful responses are downloads. Maintain these error
semantics:

- `400`: missing or empty file;
- `404`: unsupported output type;
- `415`: non-multipart request;
- `499`: client canceled;
- `500`: Office conversion failure;
- `504`: Office timeout.

`/health` proves that the ASP.NET Core host is alive. It intentionally does not
start Office and is not an Office-readiness probe.

## Office COM guidance

`OfficeDocumentConverter` uses:

- `Word.Application` for Word conversions;
- `Excel.Application` for Excel conversions;
- numeric constants matching Office enum values;
- `Marshal.FinalReleaseComObject` for explicit release.

When touching COM code:

- test the default suite first;
- run the relevant real Office smoke test;
- use a dedicated Office account;
- keep applications invisible and alerts disabled;
- open source files read-only where possible;
- never save changes back to the uploaded source;
- release child COM objects before parent applications;
- confirm output existence and process cleanup.

Do not assume a timeout can safely abort an arbitrary managed thread. The
watchdog terminates the newly identified Office process so that a blocked COM
call can unwind. Changes to process identification are high risk and require
real Office testing.

## Security considerations

The service processes uploaded Office files. Current code has no authentication,
file allow-list, explicit upload limit, or application-level rate limiter.

For security-related changes:

- validate at both reverse-proxy and application boundaries;
- do not return local paths, COM details, or stack traces;
- do not log document contents;
- use a least-privileged dedicated Windows identity;
- keep deployment and temporary directories inaccessible to untrusted users;
- treat macro-enabled and externally linked documents as untrusted input;
- never commit tokens, credentials, certificates, or machine-specific secrets.

Document any new security controls in `README.md`.

## Build and deployment

Console or Windows Service publish:

```powershell
dotnet publish `
  .\OfficeConversion.Host\OfficeConversion.Host.csproj `
  -c Release `
  -r win-x64 `
  --self-contained false `
  -o .\OfficeConversion.Host\bin\publish\service
```

IIS out-of-process publish:

```powershell
dotnet publish `
  .\OfficeConversion.Host\OfficeConversion.Host.csproj `
  -c Release `
  -r win-x64 `
  --self-contained false `
  -p:AspNetCoreHostingModel=OutOfProcess `
  -o .\OfficeConversion.Host\bin\publish\iis
```

Publishing is validation only unless the task explicitly authorizes deployment.
Do not register, start, stop, or remove a Windows Service; modify IIS; or copy
files into a production directory without explicit user authorization.

## Required validation before handoff

For documentation-only changes:

```powershell
dotnet format .\OfficeConversion.sln --verify-no-changes --no-restore
dotnet test .\OfficeConversion.sln --no-restore
```

For normal code changes, also build Release:

```powershell
dotnet build .\OfficeConversion.sln -c Release --no-restore
```

For Office conversion changes, additionally run the relevant opt-in smoke tests
and inspect for orphan processes.

Before reporting completion:

- run `git diff --check`;
- review `git status --short`;
- remove temporary debug instrumentation and generated artifacts;
- update `README.md` and this file if commands, architecture, routes,
  configuration, or hosting behavior changed.

There is no repository-specific commit-message or pull-request title convention
yet. Use concise imperative commit messages and report exactly which validation
commands passed.

## Troubleshooting

- `Class not registered`: confirm Office COM registration and process
  architecture.
- Conversion works interactively but not as a service: verify the service
  identity, Office activation/profile, filesystem permissions, fonts, and
  hidden dialogs.
- `504`: inspect timeout and process-termination logs before increasing the
  timeout.
- Build output locked: find the exact `OfficeConversion.Host` or test process
  that owns it; terminate only a process created by the current test run.
- Layout differs across machines: compare Office versions, fonts, language
  packs, printer drivers, and linked document resources.
