@echo off
setlocal
rem Build with the Windows .NET Framework compiler; no SDK or PowerShell required.
set "agregatorCompiler=%SystemRoot%\Microsoft.NET\Framework64\v4.0.30319\csc.exe"
if not exist "%agregatorCompiler%" (
    echo Error: .NET Framework 4.x x64 C# compiler not found.
    set "agregatorExitCode=1"
    goto finish
)

echo Building Agregator.exe...
"%agregatorCompiler%" /nologo /optimize+ /target:exe /platform:x64 /codepage:65001 ^
    /out:"%~dp0Agregator.build.exe" ^
    /reference:System.Core.dll ^
    /reference:System.Xml.Linq.dll ^
    /reference:System.IO.Compression.dll ^
    /reference:System.IO.Compression.FileSystem.dll ^
    /reference:System.Drawing.dll ^
    /reference:System.Windows.Forms.dll ^
    /reference:Microsoft.CSharp.dll ^
    "%~dp0sources\Agregator.cs" "%~dp0sources\SnapshotStyles.cs"
if errorlevel 1 (
    set "agregatorExitCode=1"
    goto finish
)

rem Replace the executable only after a successful build.
move /y "%~dp0Agregator.build.exe" "%~dp0Agregator.exe" >nul
if errorlevel 1 (
    echo Error: cannot replace Agregator.exe. Close any running instance.
    set "agregatorExitCode=1"
    goto finish
)

echo Build succeeded: "%~dp0Agregator.exe"
set "agregatorExitCode=0"

:finish
rem Preserve the build status across the interactive pause; do not run the program.
pause
exit /b %agregatorExitCode%