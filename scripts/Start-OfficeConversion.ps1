$ErrorActionPreference = "Stop"

$repositoryRoot = Split-Path -Parent $PSScriptRoot
$publishDirectory = Join-Path `
    $repositoryRoot `
    "OfficeConversion.Host\bin\Publish"
$executablePath = Join-Path `
    $publishDirectory `
    "OfficeConversion.Host.exe"
$logDirectory = Join-Path $publishDirectory "log"
$standardOutputPath = Join-Path $logDirectory "out.txt"
$standardErrorPath = Join-Path $logDirectory "err.txt"

if (-not (Test-Path -LiteralPath $executablePath -PathType Leaf)) {
    throw "Office Conversion executable was not found: $executablePath"
}

New-Item `
    -ItemType Directory `
    -Path $logDirectory `
    -Force |
    Out-Null

$process = Start-Process `
    -FilePath $executablePath `
    -WorkingDirectory $publishDirectory `
    -WindowStyle Hidden `
    -RedirectStandardOutput $standardOutputPath `
    -RedirectStandardError $standardErrorPath `
    -PassThru

$process.WaitForExit()
exit $process.ExitCode
