param(
    [Parameter(Mandatory = $true)]
    [string]$WorkbookPath,

    [Parameter(Mandatory = $true)]
    [string]$VbaFolderPath,

    [int]$CloseWaitSeconds = 30,

    [string]$LogPath
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$vbextCtStdModule = 1
$vbextCtClassModule = 2
$vbextCtMsForm = 3
$vbextCtDocument = 100

function Write-ReloadLog {
    param([string]$Message)

    if ([string]::IsNullOrWhiteSpace($LogPath)) {
        return
    }
    $logFolderPath = [System.IO.Path]::GetDirectoryName($LogPath)
    if (-not [string]::IsNullOrWhiteSpace($logFolderPath)) {
        [System.IO.Directory]::CreateDirectory($logFolderPath) | Out-Null
    }
    $line = '{0:yyyy-MM-dd HH:mm:ss} | {1}' -f [DateTime]::Now, $Message
    $encoding = if ([System.IO.Path]::GetFileName($LogPath) -ieq 'diagnostic.log') {
        [System.Text.Encoding]::Unicode
    }
    else {
        [System.Text.Encoding]::UTF8
    }
    [System.IO.File]::AppendAllText($LogPath, $line + [Environment]::NewLine, $encoding)
}

function Get-ComponentName {
    param([System.IO.FileInfo]$File)

    $sourceText = [System.IO.File]::ReadAllText($File.FullName, [System.Text.Encoding]::UTF8)
    $attributeMatch = [regex]::Match($sourceText, '(?m)^Attribute VB_Name = "([^"]+)"')
    if ($attributeMatch.Success) {
        return $attributeMatch.Groups[1].Value
    }
    $componentName = $File.Name
    if ($componentName.EndsWith('.utf8.vba', [System.StringComparison]::OrdinalIgnoreCase)) {
        return $componentName.Substring(0, $componentName.Length - '.utf8.vba'.Length)
    }
    if ($componentName.EndsWith('.cls.vba', [System.StringComparison]::OrdinalIgnoreCase)) {
        return $componentName.Substring(0, $componentName.Length - '.cls.vba'.Length)
    }
    return [System.IO.Path]::GetFileNameWithoutExtension($componentName)
}

function Get-ComponentType {
    param([System.IO.FileInfo]$File)

    if ($File.Name.EndsWith('.cls.vba', [System.StringComparison]::OrdinalIgnoreCase)) {
        return $vbextCtClassModule
    }
    if ($File.Name.EndsWith('.frm.vba', [System.StringComparison]::OrdinalIgnoreCase)) {
        return $vbextCtMsForm
    }
    return $vbextCtStdModule
}

function Get-ImportText {
    param([System.IO.FileInfo]$File)

    $result = New-Object System.Collections.Generic.List[string]
    $insideClassHeader = $false
    foreach ($line in [System.IO.File]::ReadAllLines($File.FullName, [System.Text.Encoding]::UTF8)) {
        $trimmedLine = $line.Trim()
        if ($trimmedLine -eq 'VERSION 1.0 CLASS') {
            $insideClassHeader = $true
            continue
        }
        if ($insideClassHeader -and $trimmedLine -eq 'END') {
            $insideClassHeader = $false
            continue
        }
        if (-not $insideClassHeader -and -not $trimmedLine.StartsWith('Attribute ')) {
            $result.Add($line)
        }
    }
    return $result -join [Environment]::NewLine
}

function Get-ConfigValue {
    param(
        [object]$Workbook,
        [string]$KeyName
    )

    foreach ($worksheet in $Workbook.Worksheets) {
        foreach ($table in $worksheet.ListObjects) {
            if ($table.Name -ine 'tbConfig' -or $null -eq $table.DataBodyRange) {
                continue
            }
            $keyColumn = $table.ListColumns.Item('Key').Index
            if ($keyColumn -ge $table.ListColumns.Count) {
                continue
            }
            foreach ($row in $table.ListRows) {
                if ([string]$row.Range.Cells.Item(1, $keyColumn).Value2 -ieq $KeyName) {
                    return ([string]$row.Range.Cells.Item(1, $keyColumn + 1).Value2).Trim()
                }
            }
        }
    }
    throw "The configuration key '$KeyName' was not found."
}

function Get-ImportFiles {
    param(
        [string]$SourceFolder,
        [string[]]$Patterns
    )

    $filesByPath = @{}
    $sourceFiles = Get-ChildItem -LiteralPath $SourceFolder -Recurse -File | Where-Object {
        $_.Name.EndsWith('.vba', [System.StringComparison]::OrdinalIgnoreCase)
    }
    foreach ($pattern in $Patterns) {
        $normalizedPattern = $pattern.Replace('/', '\')
        foreach ($sourceFile in $sourceFiles) {
            $relativePath = $sourceFile.FullName.Substring($SourceFolder.Length).TrimStart('\')
            if ($relativePath -like $normalizedPattern) {
                $filesByPath[$sourceFile.FullName] = $sourceFile
            }
        }
    }
    return @($filesByPath.Values | Sort-Object FullName)
}

function Test-DocumentModuleSource {
    param([System.IO.FileInfo]$File)

    return $File.Name -ieq 'ThisWorkbook.vba' -or $File.Name.StartsWith('ws_', [System.StringComparison]::OrdinalIgnoreCase)
}

function Set-DocumentModuleCode {
    param(
        [object]$Workbook,
        [System.IO.FileInfo]$File
    )

    if ($File.Name -ieq 'ThisWorkbook.vba') {
        $componentName = $Workbook.CodeName
    }
    else {
        $sheetName = $File.Name.Substring(3, $File.Name.Length - 7)
        $componentName = $Workbook.Worksheets.Item($sheetName).CodeName
    }
    $codeModule = $Workbook.VBProject.VBComponents.Item($componentName).CodeModule
    if ($codeModule.CountOfLines -gt 0) {
        $codeModule.DeleteLines(1, $codeModule.CountOfLines)
    }
    $codeModule.AddFromString((Get-ImportText $File))
}

if (-not (Test-Path -LiteralPath $WorkbookPath -PathType Leaf)) {
    throw "Workbook not found: $WorkbookPath"
}
if (-not (Test-Path -LiteralPath $VbaFolderPath -PathType Container)) {
    throw "VBA source folder not found: $VbaFolderPath"
}

$workbookPath = (Resolve-Path -LiteralPath $WorkbookPath).Path
$vbaFolderPath = (Resolve-Path -LiteralPath $VbaFolderPath).Path
$workbookName = [System.IO.Path]::GetFileName($workbookPath)
$lockPath = Join-Path ([System.IO.Path]::GetDirectoryName($workbookPath)) ('~$' + $workbookName)
Write-ReloadLog "VBA_EXTERNAL_RELOAD_PROCESS_STARTED | Workbook=$workbookName"
$deadline = [DateTime]::UtcNow.AddSeconds($CloseWaitSeconds)
while (Test-Path -LiteralPath $lockPath) {
    if ([DateTime]::UtcNow -ge $deadline) {
        throw "The workbook was not closed within $CloseWaitSeconds seconds: $workbookName"
    }
    Start-Sleep -Milliseconds 250
}

$excel = $null
$workbook = $null
$targetExcel = $null
try {
    try {
        $targetExcel = [Runtime.InteropServices.Marshal]::GetActiveObject('Excel.Application')
    }
    catch {
        $targetExcel = $null
    }
    Write-ReloadLog "VBA_EXTERNAL_RELOAD_STAGE | Name=ImportComponents"
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    $excel.AutomationSecurity = 3
    $workbook = $excel.Workbooks.Open($workbookPath, 0, $false)

    $profileName = Get-ConfigValue $workbook 'ThisWorkbook::id'
    $modulesConfigPath = Join-Path $vbaFolderPath 'modules.json'
    $modulesConfig = Get-Content -LiteralPath $modulesConfigPath -Raw -Encoding UTF8 | ConvertFrom-Json
    $patterns = @($modulesConfig.PSObject.Properties[$profileName].Value)
    if ($patterns.Count -eq 0) {
        throw "The profile '$profileName' has no VBA module patterns."
    }
    $importFiles = @(Get-ImportFiles $vbaFolderPath $patterns)
    if ($importFiles.Count -eq 0) {
        throw "No VBA source files match the profile '$profileName'."
    }

    for ($componentIndex = $workbook.VBProject.VBComponents.Count; $componentIndex -ge 1; $componentIndex--) {
        $component = $workbook.VBProject.VBComponents.Item($componentIndex)
        if ($component.Type -in $vbextCtStdModule, $vbextCtClassModule, $vbextCtMsForm) {
            $workbook.VBProject.VBComponents.Remove($component)
        }
        elseif ($component.Type -eq $vbextCtDocument -and $component.CodeModule.CountOfLines -gt 0) {
            $component.CodeModule.DeleteLines(1, $component.CodeModule.CountOfLines)
        }
    }

    foreach ($importFile in $importFiles) {
        if (Test-DocumentModuleSource $importFile) {
            Set-DocumentModuleCode $workbook $importFile
            continue
        }
        $component = $workbook.VBProject.VBComponents.Add((Get-ComponentType $importFile))
        $component.Name = Get-ComponentName $importFile
        $component.CodeModule.AddFromString((Get-ImportText $importFile))
    }
    $workbook.Save()
    Write-ReloadLog "VBA_EXTERNAL_RELOAD_STAGE_COMPLETED | Name=ImportComponents | ModuleCount=$($importFiles.Count)"
}
catch {
    Write-ReloadLog "VBA_EXTERNAL_RELOAD_ERROR | Description=$($_.Exception.Message)"
    throw
}
finally {
    if ($null -ne $workbook) {
        $workbook.Close($false)
        [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($workbook)
        $workbook = $null
    }
    if ($null -ne $excel) {
        $excel.Quit()
        [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($excel)
        $excel = $null
    }
}

Write-ReloadLog "VBA_EXTERNAL_RELOAD_OPEN_REQUESTED | Workbook=$workbookName"
$openedInActiveExcel = $false
if ($null -ne $targetExcel) {
    try {
        $targetExcel.Workbooks.Open($workbookPath) | Out-Null
        $openedInActiveExcel = $true
    }
    catch {
        Write-ReloadLog "VBA_EXTERNAL_RELOAD_OPEN_ACTIVE_EXCEL_ERROR | Description=$($_.Exception.Message)"
    }
}
if (-not $openedInActiveExcel) {
    [System.GC]::Collect()
    [System.GC]::WaitForPendingFinalizers()
    $excelArgument = '"{0}"' -f $workbookPath.Replace('"', '""')
    Start-Process -FilePath 'excel.exe' -ArgumentList $excelArgument
}