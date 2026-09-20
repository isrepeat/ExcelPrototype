@echo off
setlocal
set "watcher_script=%~dp0PowerShell\TrailingWhitespaceWatcher.ps1"
schtasks.exe /Change /TN "UtilityHelpersLib Trailing Whitespace Watcher" /Disable
set "task_exit_code=%errorlevel%"
powershell.exe -NoProfile -ExecutionPolicy Bypass -Command "$ErrorActionPreference = 'Stop'; $watcherScript = [System.IO.Path]::GetFullPath($env:watcher_script); Get-CimInstance Win32_Process -Filter \"Name = 'powershell.exe'\" | Where-Object { $_.CommandLine -like ('*' + $watcherScript + '*') } | ForEach-Object { Invoke-CimMethod -InputObject $_ -MethodName Terminate | Out-Null }"
set "watcher_exit_code=%errorlevel%"
if not "%task_exit_code%"=="0" (
    echo Failed to disable the scheduled task.
)
if not "%watcher_exit_code%"=="0" (
    echo Failed to stop the running watcher.
)
if "%task_exit_code%"=="0" if "%watcher_exit_code%"=="0" (
    echo The trailing whitespace watcher is stopped and disabled.
)
pause
if not "%task_exit_code%"=="0" exit /b %task_exit_code%
exit /b %watcher_exit_code%