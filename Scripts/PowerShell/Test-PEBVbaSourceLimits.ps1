Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'
. (Join-Path $PSScriptRoot 'WorkbookModules.ps1')
$sourceRoot = Join-Path $PSScriptRoot '../../PROJECTS_EXCEL/vba/PersonalEventBuilder'
$fileCount = 0
$maximumLength = 0
foreach ($file in Get-ChildItem -LiteralPath $sourceRoot -Recurse -Filter '*.vba') {
    # Measure the code sent to VBE, after Unicode literals expand into ChrW calls.
    $code = ConvertTo-WorkbookModuleCode ([IO.File]::ReadAllText($file.FullName))
    $lineNumber = 0
    $continuations = 0
    foreach ($line in ($code -split '\r?\n')) {
        $lineNumber++
        $maximumLength = [Math]::Max($maximumLength, $line.Length)
        if ($line.Length -gt 1023) {
            throw "$($file.Name):${lineNumber}: encoded VBA line exceeds 1023 characters ($($line.Length))."
        }
        if ($line -match ' _$') {
            $continuations++
            if ($continuations -gt 24) {
                throw "$($file.Name):${lineNumber}: VBA statement exceeds 24 line continuations."
            }
        }
        else {
            $continuations = 0
        }
    }
    $fileCount++
}
Write-Output "PASS: $fileCount PEB source files; maximum encoded line length $maximumLength."