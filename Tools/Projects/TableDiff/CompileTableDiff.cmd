@echo off
setlocal
set "tableDiffDirectory=%~dp0"
rem PUSHD maps UNC locations to a temporary drive and saves the caller's directory.
set "tableDiffDirectoryPushed=0"
pushd "%tableDiffDirectory%"
if errorlevel 1 (
    echo Error: cannot access the build script directory.
    set "tableDiffExitCode=1"
    goto finish
)
set "tableDiffDirectoryPushed=1"

rem Build with the Windows .NET Framework compiler; no SDK or PowerShell required.
set "tableDiffCompiler=%SystemRoot%\Microsoft.NET\Framework64\v4.0.30319\csc.exe"
if not exist "%tableDiffCompiler%" (
    echo Error: .NET Framework 4.x x64 C# compiler not found.
    set "tableDiffExitCode=1"
    goto finish
)

echo Building TableDiff.exe...
set "tableDiffOutputDirectory=%tableDiffDirectory%build"
set "tableDiffBuildDirectory=%tableDiffOutputDirectory%\.staging"
call "%tableDiffDirectory%..\..\common\CompileCommon.cmd" "%tableDiffBuildDirectory%"
if errorlevel 1 (
    set "tableDiffExitCode=1"
    goto finish
)
"%tableDiffCompiler%" /nologo /optimize+ /target:exe /platform:x64 /codepage:65001 ^
    /out:"%tableDiffBuildDirectory%\TableDiff.exe" ^
    /reference:"%tableDiffBuildDirectory%\ExcelTools.Tooling.dll" ^
    /reference:"%tableDiffBuildDirectory%\ExcelTools.Tooling.Console.dll" ^
    /reference:"%tableDiffBuildDirectory%\ExcelTools.Excel.Tables.dll" ^
    /reference:"%tableDiffBuildDirectory%\ExcelTools.Excel.OpenXml.dll" ^
    /reference:System.Core.dll ^
    /reference:System.Xml.Linq.dll ^
    /reference:System.IO.Compression.dll ^
    /reference:System.IO.Compression.FileSystem.dll ^
    "%tableDiffDirectory%sources\Program.cs" ^
    "%tableDiffDirectory%sources\Configuration\ConfigurationReader.cs" ^
    "%tableDiffDirectory%sources\Comparison\TableComparer.cs" ^
    "%tableDiffDirectory%sources\Reporting\HtmlDiffWriter.cs" ^
    /resource:"%tableDiffDirectory%sources\Reporting\report.css",report.css ^
    /resource:"%tableDiffDirectory%sources\Reporting\report.js",report.js
if errorlevel 1 (
    set "tableDiffExitCode=1"
    goto finish
)

rem Publish the complete compiled set only after all compilations succeed.
copy /y "%tableDiffBuildDirectory%\ExcelTools.*.dll" "%tableDiffOutputDirectory%\" >nul
if errorlevel 1 (
    echo Error: cannot replace shared libraries. Close any running instance.
    set "tableDiffExitCode=1"
    goto finish
)
move /y "%tableDiffBuildDirectory%\TableDiff.exe" "%tableDiffOutputDirectory%\TableDiff.exe" >nul
if errorlevel 1 (
    echo Error: cannot replace TableDiff.exe. Close any running instance.
    set "tableDiffExitCode=1"
    goto finish
)

echo Build succeeded: "%tableDiffOutputDirectory%\TableDiff.exe"
set "tableDiffExitCode=0"

:finish
rem Remove intermediate binaries on both successful and failed builds.
if "%tableDiffDirectoryPushed%"=="1" (
    if exist "build\.staging\" rmdir /s /q "build\.staging"
    if exist "build\.staging\" (
        echo Error: cannot remove the staging directory.
        set "tableDiffExitCode=1"
    )
)
rem Release the temporary network drive and restore the original directory.
if "%tableDiffDirectoryPushed%"=="1" popd
rem Preserve the build status across the interactive pause; do not run the program.
pause
exit /b %tableDiffExitCode%