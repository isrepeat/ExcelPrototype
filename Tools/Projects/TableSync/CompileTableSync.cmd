@echo off
setlocal
set "tableSyncOutput=%~dp0build"
call "%~dp0..\TableDiff\CompileTableDiff.cmd" --no-pause
if errorlevel 1 goto failed
if not exist "%tableSyncOutput%\" mkdir "%tableSyncOutput%"
if errorlevel 1 goto failed
copy /y "%~dp0..\TableDiff\build\ExcelTools.*.dll" "%tableSyncOutput%\" >nul
if errorlevel 1 goto failed
if not exist "%tableSyncOutput%\plugins\" mkdir "%tableSyncOutput%\plugins"
if errorlevel 1 goto failed
"%SystemRoot%\Microsoft.NET\Framework64\v4.0.30319\csc.exe" /nologo /target:library /codepage:65001 ^
    /out:"%tableSyncOutput%\TableSync.Contracts.dll" ^
    "%~dp0..\..\common\TableSync.Contracts\SyncPlugin.cs"
if errorlevel 1 goto failed
"%SystemRoot%\Microsoft.NET\Framework64\v4.0.30319\csc.exe" /nologo /target:library /codepage:65001 ^
    /out:"%tableSyncOutput%\plugins\TableSync.Plugins.Personnel.dll" ^
    /reference:"%tableSyncOutput%\TableSync.Contracts.dll" ^
    "%~dp0plugins\Personnel\PersonnelPlugin.cs"
if errorlevel 1 goto failed
"%SystemRoot%\Microsoft.NET\Framework64\v4.0.30319\csc.exe" /nologo /optimize+ /target:exe /platform:x64 /codepage:65001 ^
    /out:"%tableSyncOutput%\TableSync.exe" ^
    /reference:System.Core.dll ^
    /reference:System.Xml.dll ^
    /reference:System.Xml.Linq.dll ^
    /reference:"%tableSyncOutput%\ExcelTools.Excel.Changes.dll" ^
    /reference:"%tableSyncOutput%\ExcelTools.Tooling.dll" ^
    /reference:"%tableSyncOutput%\ExcelTools.Tooling.Console.dll" ^
    /reference:"%tableSyncOutput%\ExcelTools.Excel.Tables.dll" ^
    /reference:"%tableSyncOutput%\ExcelTools.Excel.OpenXml.dll" ^
    /reference:"%tableSyncOutput%\TableSync.Contracts.dll" ^
    "%~dp0sources\Program.cs"
if errorlevel 1 goto failed
echo Build succeeded: "%tableSyncOutput%\TableSync.exe"
set "tableSyncExitCode=0"
goto finish

:failed
set "tableSyncExitCode=%errorlevel%"
echo Build failed. Exit code: %tableSyncExitCode%

:finish
pause
exit /b %tableSyncExitCode%