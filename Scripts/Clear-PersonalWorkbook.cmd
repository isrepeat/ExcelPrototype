@echo off
powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%~dp0PowerShell\Update-WorkbookModules.ps1" -WorkbookName PERSONAL.XLSB -Mode Clear
pause