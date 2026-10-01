@echo off
powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%~dp0PowerShell\Update-WorkbookModules.ps1" -WorkbookName PERSONAL.XLSB -VbaFolderPath "%~dp0..\MacrosExcel\PERSONAL" -Profile PERSONAL -InitializeMacro ex_Core.fn_ReloadPersonalRuntime
pause