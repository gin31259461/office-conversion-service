Option Explicit

Dim fileSystem
Dim shell
Dim scriptDirectory
Dim powershellPath
Dim powershellScriptPath
Dim command
Dim exitCode

Set fileSystem = CreateObject("Scripting.FileSystemObject")
Set shell = CreateObject("WScript.Shell")

scriptDirectory = fileSystem.GetParentFolderName(WScript.ScriptFullName)
powershellPath = shell.ExpandEnvironmentStrings( _
    "%SystemRoot%\System32\WindowsPowerShell\v1.0\powershell.exe")
powershellScriptPath = fileSystem.BuildPath( _
    scriptDirectory, _
    "Start-OfficeConversion.ps1")

shell.CurrentDirectory = scriptDirectory

command = """" & powershellPath & """" & _
    " -NoLogo -NoProfile -NonInteractive" & _
    " -WindowStyle Hidden -ExecutionPolicy Bypass" & _
    " -File """ & powershellScriptPath & """"

exitCode = shell.Run(command, 0, True)
WScript.Quit exitCode
