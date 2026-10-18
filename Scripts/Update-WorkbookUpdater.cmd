@echo off
powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%~dp0..\PROJECTS_EXCEL\WorkbookUpdater\Build-Updater.ps1" -UpdateExisting
set "updateExitCode=%errorlevel%"
pause
exit /b %updateExitCode%