[CmdletBinding()]
param(
    [string]$ConfigurationPath
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$elevationModulePath = Join-Path $PSScriptRoot 'Modules\ElevationModule\ElevationModule.psm1'
Import-Module -Name $elevationModulePath -Force
$scriptArguments = @()
if ($PSBoundParameters.ContainsKey('ConfigurationPath')) {
    $scriptArguments += @('-ConfigurationPath', $ConfigurationPath)
}
if (Restart-PowerShellScriptElevated -ScriptPath $PSCommandPath -ScriptArguments $scriptArguments) {
    return
}

$runnerPath = Join-Path $PSScriptRoot 'Run-TrailingWhitespaceWatcher.vbs'
$taskName = 'UtilityHelpersLib Trailing Whitespace Watcher'
$action = New-ScheduledTaskAction -Execute 'wscript.exe' -Argument "`"$runnerPath`""
$trigger = New-ScheduledTaskTrigger -AtLogOn
$settings = New-ScheduledTaskSettingsSet -MultipleInstances IgnoreNew -StartWhenAvailable
$principal = New-ScheduledTaskPrincipal -UserId "$env:USERDOMAIN\$env:USERNAME" -LogonType Interactive -RunLevel Limited
Register-ScheduledTask -TaskName $taskName -Action $action -Trigger $trigger -Settings $settings -Principal $principal -Description 'Keeps configured text files free of terminal whitespace.' -Force
Start-ScheduledTask -TaskName $taskName
Write-Host "Installed and started scheduled task '$taskName'."
