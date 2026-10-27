param(
    [string]$ProjectRoot = (Split-Path (Split-Path $PSScriptRoot -Parent) -Parent)
)

$ErrorActionPreference = 'Stop'
$folder = Join-Path $ProjectRoot 'workbook/PeopleRegistry'
$bookPath = Join-Path $folder 'PeopleRegistry.xlsm'
if (Test-Path -LiteralPath $bookPath) {
    throw 'The generator already exists. Use Ctrl+Alt+R to update it without replacing its data.'
}
[void][IO.Directory]::CreateDirectory($folder)
$configRows = [IO.File]::ReadAllLines((Join-Path $ProjectRoot 'config/PeopleRegistry/wsConfig.txt'), [Text.UTF8Encoding]::new($false, $true))
$config = @{}
foreach ($line in $configRows) {
    $parts = $line.Split("`t")
    if ($parts.Count -ne 3 -or $config.ContainsKey($parts[1])) {
        throw 'Invalid or duplicate configuration row.'
    }
    $config.Add($parts[1], $parts[2])
}

function Get-Config([string]$Key) {
    $fullKey = 'PeopleRegistry::' + $Key
    if (-not $config.ContainsKey($fullKey)) {
        throw "Missing configuration key: $fullKey"
    }
    return $config[$fullKey]
}

$excel = $null
$book = $null
try {
    # Author native tables and import source without compiling or executing VBA.
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    $excel.AutomationSecurity = 3
    $excel.EnableEvents = $false
    $excel.ScreenUpdating = $false
    $book = $excel.Workbooks.Add(-4167)
    $inputSheet = $book.Worksheets.Item(1)
    $inputSheet.Name = Get-Config 'Sheet.Input'
    $service = $book.Worksheets.Add([Type]::Missing, $inputSheet)
    $service.Name = 'wsConfig'
    foreach ($column in 1..3) {
        $service.Columns.Item($column).NumberFormat = '@'
        $service.Cells.Item(1, $column).Value2 = @('Kind', 'Key', 'Value')[$column - 1]
    }
    $row = 2
    foreach ($line in $configRows) {
        $parts = $line.Split("`t")
        foreach ($column in 1..3) {
            $service.Cells.Item($row, $column).Value2 = $parts[$column - 1]
        }
        $row++
    }
    $table = $service.ListObjects.Add(1, $service.Range("A1:C$($row - 1)"), [Type]::Missing, 1)
    $table.Name = 'tbConfig'
    $table.TableStyle = ''
    foreach ($kind in @('History')) {
        $sheet = $book.Worksheets.Add([Type]::Missing, $book.Worksheets.Item($book.Worksheets.Count))
        $sheet.Name = Get-Config 'Sheet.History'
        $anchor = Get-Config 'Layout.HistoryCell'
        $start = $sheet.Range($anchor)
        $headers = (Get-Config "Header.$kind").Split('|')
        foreach ($index in 0..($headers.Count - 1)) {
            $start.Offset(0, $index).Value2 = $headers[$index]
        }
        $start.Offset(-2).Value2 = Get-Config "Title.$kind"
        $table = $sheet.ListObjects.Add(1, $start.Resize(2, $headers.Count), [Type]::Missing, 1)
        $table.Name = Get-Config "Table.$kind"
        $table.TableStyle = ''
        foreach ($index in 1..$headers.Count) {
            # History timestamps are immutable ISO text; no numeric/date format is required.
            $table.ListColumns.Item($index).Range.EntireColumn.NumberFormat = '@'
        }
        if ($kind -ne 'Input') {
            [void]$table.ListRows.Item(1).Delete()
        }
    }
    . (Join-Path $PSScriptRoot 'Set-GeneratorForm.ps1')
    . (Join-Path $PSScriptRoot 'Set-OriginAliases.ps1')
    Set-OriginAliases $book $config
    Set-GeneratorForm $book $config
    . (Join-Path $PSScriptRoot 'Import-Modules.ps1')
    Import-PeopleRegistryModules $book (Split-Path $PSScriptRoot -Parent)
    [void]$inputSheet.Activate()
    $service.Visible = 0
    $excel.ActiveWindow.Zoom = [int](Get-Config 'Layout.Zoom')
    . (Join-Path $PSScriptRoot 'Set-HistoryProtection.ps1')
    Set-HistoryProtection $book $config
    [void]$book.SaveAs($bookPath, 52)
    $book.Close($false)
    $book = $null
    $book = $excel.Workbooks.Open($bookPath, 0, $true)
    if ($book.Worksheets.Count -ne 4) {
        throw 'Unexpected workbook layout.'
    }
    'Generator prepared and reopened without VBA compilation or execution.'
} finally {
    if ($null -ne $book) {
        $book.Close($false)
    }
    if ($null -ne $excel) {
        $excel.ScreenUpdating = $true
        $excel.Quit()
        [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($excel)
    }
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
}