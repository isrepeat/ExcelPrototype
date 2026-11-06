@echo off
setlocal
rem Build reusable libraries into the destination supplied by the calling tool.
if "%~1"=="" (
    echo Error: specify the library output directory.
    goto failed
)
set "commonOutput=%~f1"
set "commonCompiler=%SystemRoot%\Microsoft.NET\Framework64\v4.0.30319\csc.exe"
if not exist "%commonCompiler%" (
    echo Error: .NET Framework 4.x x64 C# compiler not found.
    goto failed
)
if not exist "%commonOutput%\" mkdir "%commonOutput%"
if errorlevel 1 goto failed
echo Building ExcelTools.Tooling.dll...
"%commonCompiler%" /nologo /optimize+ /target:library /platform:x64 /codepage:65001 ^
    /out:"%commonOutput%\ExcelTools.Tooling.dll" ^
    /reference:System.dll ^
    /reference:System.Core.dll ^
    "%~dp0Tooling\ToolPaths.cs" ^
    "%~dp0Tooling\OutputPublisher.cs"
if errorlevel 1 goto failed

echo Building ExcelTools.Tooling.Console.dll...
"%commonCompiler%" /nologo /optimize+ /target:library /platform:x64 /codepage:65001 ^
    /out:"%commonOutput%\ExcelTools.Tooling.Console.dll" ^
    /reference:System.dll ^
    /reference:System.Core.dll ^
    "%~dp0Tooling.Console\ConsoleTool.cs"
if errorlevel 1 goto failed

echo Building ExcelTools.Excel.Tables.dll...
"%commonCompiler%" /nologo /optimize+ /target:library /platform:x64 /codepage:65001 ^
    /out:"%commonOutput%\ExcelTools.Excel.Tables.dll" ^
    /reference:System.dll ^
    /reference:System.Core.dll ^
    "%~dp0Excel.Tables\WorkbookSource.cs" ^
    "%~dp0Excel.Tables\TableSnapshot.cs" ^
    "%~dp0Excel.Tables\TableMatching.cs"
if errorlevel 1 goto failed

echo Building ExcelTools.Excel.Changes.dll...
"%commonCompiler%" /nologo /optimize+ /target:library /platform:x64 /codepage:65001 ^
    /out:"%commonOutput%\ExcelTools.Excel.Changes.dll" ^
    /reference:System.Core.dll ^
    /reference:System.Xml.dll ^
    /reference:System.Xml.Linq.dll ^
    /reference:"%commonOutput%\ExcelTools.Tooling.dll" ^
    /reference:"%commonOutput%\ExcelTools.Excel.Tables.dll" ^
    "%~dp0Excel.Changes\ConfigurationReader.cs" ^
    "%~dp0Excel.Changes\PreparedReport.cs"
if errorlevel 1 goto failed

echo Building ExcelTools.Excel.Interop.dll...
"%commonCompiler%" /nologo /optimize+ /target:library /platform:x64 /codepage:65001 ^
    /out:"%commonOutput%\ExcelTools.Excel.Interop.dll" ^
    /reference:System.dll ^
    /reference:System.Core.dll ^
    /reference:Microsoft.CSharp.dll ^
    "%~dp0Excel.Interop\ComObjects.cs" ^
    "%~dp0Excel.Interop\ExcelSession.cs" ^
    "%~dp0Excel.Interop\ExcelValues.cs"
if errorlevel 1 goto failed
echo Building ExcelTools.Excel.OpenXml.dll...
"%commonCompiler%" /nologo /optimize+ /target:library /platform:x64 /codepage:65001 ^
    /out:"%commonOutput%\ExcelTools.Excel.OpenXml.dll" ^
    /reference:System.dll ^
    /reference:System.Core.dll ^
    /reference:System.Xml.dll ^
    /reference:System.Xml.Linq.dll ^
    /reference:System.IO.Compression.dll ^
    /reference:"%commonOutput%\ExcelTools.Excel.Tables.dll" ^
    "%~dp0Excel.OpenXml\WorkbookReader.cs" ^
    "%~dp0Excel.OpenXml\FormulaTranslator.cs" ^
    "%~dp0Excel.OpenXml\WorkbookWriter.cs"
if errorlevel 1 goto failed
echo Shared libraries built successfully.
set "commonExitCode=0"
goto finish

:failed
echo Shared library build failed.
set "commonExitCode=1"

:finish
if /i not "%~2"=="--no-pause" pause
exit /b %commonExitCode%