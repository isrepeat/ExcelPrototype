param(
    [ValidateSet('Update', 'Clear')]
    [string]$Mode = 'Update',
    [string]$SourcePath = (Join-Path $PSScriptRoot '..\..\MacrosExcel\PERSONAL')
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$vbextCtStdModule = 1
$vbextCtClassModule = 2
$vbextCtDocument = 100

function Get-ComponentName {
    param([System.IO.FileInfo]$File)

    if ($File.Name -eq 'ThisWorkbook.vba') {
        return 'ThisWorkbook'
    }
    $attribute = Select-String -LiteralPath $File.FullName -Pattern '^Attribute VB_Name = "([^"]+)"' |
        Select-Object -First 1
    if ($attribute) {
        return $attribute.Matches[0].Groups[1].Value
    }
    return [System.IO.Path]::GetFileNameWithoutExtension($File.Name)
}

function Get-ImportText {
    param([System.IO.FileInfo]$File)

    $lines = Get-Content -LiteralPath $File.FullName
    $lines | Where-Object {
        $_ -notmatch '^VERSION ' -and $_ -notmatch '^Attribute VB_'
    }
}

function Set-ComponentCode {
    param(
        [object]$Component,
        [string[]]$Lines
    )

    $codeModule = $Component.CodeModule
    if ($codeModule.CountOfLines -gt 0) {
        $codeModule.DeleteLines(1, $codeModule.CountOfLines)
    }
    if ($Lines.Count -gt 0) {
        $codeModule.AddFromString(($Lines -join [Environment]::NewLine))
    }
}

function Remove-StandardAndClassModules {
    param([object]$Workbook)

    $componentsToRemove = @($Workbook.VBProject.VBComponents | Where-Object {
        $_.Type -in $vbextCtStdModule, $vbextCtClassModule
    })
    $componentIndex = 0
    foreach ($component in $componentsToRemove) {
        $componentIndex++
        Write-Host "Removing component $componentIndex/$($componentsToRemove.Count): $($component.Name)"
        $Workbook.VBProject.VBComponents.Remove($component)
    }
    return $componentsToRemove
}

$excel = $null
try {
    $excel = [Runtime.InteropServices.Marshal]::GetActiveObject('Excel.Application')
}
catch {
    throw 'Open Excel and load PERSONAL.XLSB before running this script.'
}

$personalWorkbook = @($excel.Workbooks | Where-Object {
    $_.Name -ieq 'PERSONAL.XLSB'
}) | Select-Object -First 1
if ($null -eq $personalWorkbook) {
    throw 'PERSONAL.XLSB is not loaded in the active Excel instance.'
}

if ($Mode -eq 'Clear') {
    $componentsToRemove = @(Remove-StandardAndClassModules $personalWorkbook)
    $personalWorkbook.Save()
    Write-Host "Removed PERSONAL.XLSB modules and classes: $($componentsToRemove.Count)"
    exit 0
}

if (-not (Test-Path -LiteralPath $SourcePath -PathType Container)) {
    throw "PERSONAL source folder was not found: $SourcePath"
}

$sourceFiles = @(Get-ChildItem -LiteralPath $SourcePath -File |
    Where-Object { $_.Extension -in '.vba', '.cls' } |
    Sort-Object Name)
if ($sourceFiles.Count -eq 0) {
    throw "No VBA source files were found: $SourcePath"
}

$removedCount = @(Remove-StandardAndClassModules $personalWorkbook).Count
$updatedCount = 0
$sourceFileCount = $sourceFiles.Count
foreach ($sourceFile in $sourceFiles) {
    Write-Host "Importing component $($updatedCount + 1)/${sourceFileCount}: $($sourceFile.Name)"
    $componentName = Get-ComponentName $sourceFile
    $component = $null
    try {
        $component = $personalWorkbook.VBProject.VBComponents.Item($componentName)
    }
    catch {
        $component = $null
    }
    if ($sourceFile.Name -eq 'ThisWorkbook.vba' -and $null -eq $component) {
        $component = @($personalWorkbook.VBProject.VBComponents | Where-Object {
            $_.Type -eq $vbextCtDocument
        }) | Select-Object -First 1
    }

    $sourceLines = @(Get-ImportText $sourceFile)
    if ($null -ne $component) {
        Set-ComponentCode $component $sourceLines
    }
    else {
        if ($sourceFile.Name -eq 'ThisWorkbook.vba') {
            throw 'PERSONAL.XLSB has no document module for ThisWorkbook.vba.'
        }
        $componentType = if ($sourceFile.Name.StartsWith('obj_')) {
            $vbextCtClassModule
        }
        else {
            $vbextCtStdModule
        }
        $component = $personalWorkbook.VBProject.VBComponents.Add($componentType)
        $component.Name = $componentName
        Set-ComponentCode $component $sourceLines
    }
    $updatedCount++
}

$personalWorkbook.Save()
try {
    $excel.Run("'PERSONAL.XLSB'!ex_Core.fn_ReloadPersonalRuntime")
}
catch {
    throw "PERSONAL.XLSB was updated, but runtime reload failed: $($_.Exception.Message)"
}
Write-Host "Removed PERSONAL.XLSB modules and classes: $removedCount"
Write-Host "Updated PERSONAL.XLSB components: $updatedCount"
Write-Host 'PERSONAL.XLSB runtime reloaded.'