@echo off
setlocal
set "elevation_module=%~dp0PowerShell\Modules\ElevationModule\ElevationModule.psm1"
set "elevated_command=/c ""%~f0"""

powershell.exe -NoProfile -ExecutionPolicy Bypass -Command "Import-Module -Name $env:elevation_module -Force; if (Test-Administrator) { exit 0 }; Start-ElevatedProcess -FilePath $env:ComSpec -ArgumentList $env:elevated_command; exit 2"
set "elevation_exit_code=%errorlevel%"
if "%elevation_exit_code%"=="2" (
    rem The elevated copy was started successfully; this non-elevated copy must exit.
    exit /b 0
)
if not "%elevation_exit_code%"=="0" (
    echo Unable to request administrator rights.
    pause
    exit /b 1
)

powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%~dp0PowerShell\Install-TrailingWhitespaceWatcher.ps1"
set "exit_code=%errorlevel%"
if not "%exit_code%"=="0" (
    echo Failed to install or start the trailing whitespace watcher.
) else (
    echo The trailing whitespace watcher is installed and running.
)
pause
exit /b %exit_code%
