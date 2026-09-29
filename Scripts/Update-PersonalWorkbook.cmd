@echo off
powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%~dp0PowerShell\Update-PersonalWorkbook.ps1" -Mode Update
pause