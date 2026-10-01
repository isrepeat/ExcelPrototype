@echo off
powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%~dp0PowerShell\Update-WorkbookModules.ps1" -WorkbookName "Personal event builder.xlsm" -VbaFolderPath "%~dp0..\PROJECTS_EXCEL\vba" -Profile PersonalEventBuilder -InitializeMacro ex_PersonalEventBuilder.fn_Initialize
set "updateExitCode=%errorlevel%"
pause
exit /b %updateExitCode%