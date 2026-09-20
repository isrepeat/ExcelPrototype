Set-StrictMode -Version Latest

function Test-Administrator {
    [OutputType([bool])]
    param()

    $identity = [Security.Principal.WindowsIdentity]::GetCurrent()
    $principal = [Security.Principal.WindowsPrincipal]::new($identity)
    return $principal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
}

function Start-ElevatedProcess {
    param(
        [Parameter(Mandatory)]
        [string]$FilePath,

        [string[]]$ArgumentList = @()
    )

    Start-Process -FilePath $FilePath -ArgumentList $ArgumentList -Verb RunAs | Out-Null
}

function Restart-PowerShellScriptElevated {
    [OutputType([bool])]
    param(
        [Parameter(Mandatory)]
        [string]$ScriptPath,

        [string[]]$ScriptArguments = @()
    )

    if (Test-Administrator) {
        return $false
    }

    $resolvedScriptPath = (Resolve-Path -LiteralPath $ScriptPath -ErrorAction Stop).Path
    $powerShellArguments = @(
        '-NoProfile',
        '-ExecutionPolicy', 'Bypass',
        '-File', $resolvedScriptPath
    ) + $ScriptArguments
    Start-ElevatedProcess -FilePath 'powershell.exe' -ArgumentList $powerShellArguments
    return $true
}

Export-ModuleMember -Function Test-Administrator, Start-ElevatedProcess, Restart-PowerShellScriptElevated
