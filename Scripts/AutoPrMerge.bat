@ECHO OFF
SETLOCAL

REM The PowerShell implementation and its MessagingModule are kept inside
REM this repository, so AutoPrMerge does not depend on UtilityHelpersLib.
SET "_REPO_ROOT=%~dp0.."
SET "_SCRIPT=%~dp0PowerShell\AutoPrMerge.ps1"

IF NOT EXIST "%_SCRIPT%" (
    ECHO [AutoPrMerge] Script not found: "%_SCRIPT%"
    SET "_RC=1"
    GOTO :EXIT_SCRIPT
)

powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%_SCRIPT%" -BatDir "%_REPO_ROOT%" %*
SET "_RC=%ERRORLEVEL%"

:EXIT_SCRIPT
PAUSE
ENDLOCAL & EXIT /B %_RC%
