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
set "agregatorOutputDirectory=%agregatorDirectory%build"
set "agregatorBuildDirectory=%agregatorOutputDirectory%\.staging"
call "%agregatorDirectory%..\..\common\CompileCommon.cmd" "%agregatorBuildDirectory%" --no-pause
if errorlevel 1 (
    set "agregatorExitCode=1"
    goto finish
)
"%agregatorCompiler%" /nologo /optimize+ /target:exe /platform:x64 /codepage:65001 ^
    /out:"%agregatorBuildDirectory%\Agregator.exe" ^
    /reference:"%agregatorBuildDirectory%\ExcelTools.Tooling.dll" ^
    /reference:"%agregatorBuildDirectory%\ExcelTools.Tooling.Console.dll" ^
    /reference:"%agregatorBuildDirectory%\ExcelTools.Excel.Tables.dll" ^
    /reference:"%agregatorBuildDirectory%\ExcelTools.Excel.Interop.dll" ^
    /reference:System.Core.dll ^
    /reference:System.Xml.Linq.dll ^
    /reference:System.IO.Compression.dll ^
    /reference:System.IO.Compression.FileSystem.dll ^
    /reference:System.Drawing.dll ^
    /reference:System.Windows.Forms.dll ^
    /reference:Microsoft.CSharp.dll ^
    "%agregatorDirectory%sources\Program.cs" ^
    "%agregatorDirectory%sources\Configuration\ConfigurationReader.cs" ^
    "%agregatorDirectory%sources\Aggregation\AggregationService.cs" ^
    "%agregatorDirectory%sources\SnapshotStyles.cs"
if errorlevel 1 (
    set "agregatorExitCode=1"
    goto finish
)

rem Publish the complete compiled set only after all compilations succeed.
copy /y "%agregatorBuildDirectory%\ExcelTools.*.dll" "%agregatorOutputDirectory%\" >nul
if errorlevel 1 (
    echo Error: cannot replace shared libraries. Close any running instance.
    set "agregatorExitCode=1"
    goto finish
)
move /y "%agregatorBuildDirectory%\Agregator.exe" "%agregatorOutputDirectory%\Agregator.exe" >nul
if errorlevel 1 (
    echo Error: cannot replace Agregator.exe. Close any running instance.
    set "agregatorExitCode=1"
    goto finish
)

echo Build succeeded: "%agregatorOutputDirectory%\Agregator.exe"
set "agregatorExitCode=0"

:finish
rem Remove intermediate binaries on both successful and failed builds.
if "%agregatorDirectoryPushed%"=="1" (
    if exist "build\.staging\" rmdir /s /q "build\.staging"
    if exist "build\.staging\" (
        echo Error: cannot remove the staging directory.
        set "agregatorExitCode=1"
    )
)
rem Release the temporary network drive and restore the original directory.
if "%agregatorDirectoryPushed%"=="1" popd
rem Preserve the build status across the interactive pause; do not run the program.
pause
exit /b %agregatorExitCode%