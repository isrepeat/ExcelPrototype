@echo off
powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%~dp0PowerShell\Update-WorkbookModules.ps1" -WorkbookPath "%~dp0..\PROJECTS_EXCEL\WorkbookUpdater\WorkbookUpdater.xlam" -VbaFolderPath "%~dp0..\PROJECTS_EXCEL\WorkbookUpdater" -Profile WorkbookUpdater
set "updateExitCode=%errorlevel%"
pause
exit /b %updateExitCode%