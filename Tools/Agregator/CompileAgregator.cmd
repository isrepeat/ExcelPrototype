@echo off
setlocal
set "agregatorDirectory=%~dp0"
rem PUSHD maps UNC locations to a temporary drive and saves the caller's directory.
set "agregatorDirectoryPushed=0"
pushd "%agregatorDirectory%"
if errorlevel 1 (
    echo Error: cannot access the build script directory.
    set "agregatorExitCode=1"
    goto finish
)
set "agregatorDirectoryPushed=1"

rem Build with the Windows .NET Framework compiler; no SDK or PowerShell required.
set "agregatorCompiler=%SystemRoot%\Microsoft.NET\Framework64\v4.0.30319\csc.exe"
if not exist "%agregatorCompiler%" (
    echo Error: .NET Framework 4.x x64 C# compiler not found.
    set "agregatorExitCode=1"
    goto finish
)

echo Building Agregator.exe...
"%agregatorCompiler%" /nologo /optimize+ /target:exe /platform:x64 /codepage:65001 ^
    /out:"%agregatorDirectory%Agregator.build.exe" ^
    /reference:System.Core.dll ^
    /reference:System.Xml.Linq.dll ^
    /reference:System.IO.Compression.dll ^
    /reference:System.IO.Compression.FileSystem.dll ^
    /reference:System.Drawing.dll ^
    /reference:System.Windows.Forms.dll ^
    /reference:Microsoft.CSharp.dll ^
    "%agregatorDirectory%sources\Agregator.cs" "%agregatorDirectory%sources\SnapshotStyles.cs"
if errorlevel 1 (
    set "agregatorExitCode=1"
    goto finish
)

rem Replace the executable only after a successful build.
move /y "%agregatorDirectory%Agregator.build.exe" "%agregatorDirectory%Agregator.exe" >nul
if errorlevel 1 (
    echo Error: cannot replace Agregator.exe. Close any running instance.
    set "agregatorExitCode=1"
    goto finish
)

echo Build succeeded: "%agregatorDirectory%Agregator.exe"
set "agregatorExitCode=0"

:finish
rem Release the temporary network drive and restore the original directory.
if "%agregatorDirectoryPushed%"=="1" popd
rem Preserve the build status across the interactive pause; do not run the program.
pause
exit /b %agregatorExitCode%