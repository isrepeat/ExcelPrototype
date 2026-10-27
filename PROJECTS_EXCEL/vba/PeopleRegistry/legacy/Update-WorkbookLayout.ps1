param(
    [string]$ProjectRoot = (Split-Path (Split-Path $PSScriptRoot -Parent) -Parent)
)

$ErrorActionPreference = 'Stop'
$bookPath = [IO.Path]::GetFullPath((Join-Path $ProjectRoot 'workbook/PeopleRegistry/PeopleRegistry.xlsm'))
$configRows = [IO.File]::ReadAllLines((Join-Path $ProjectRoot 'config/PeopleRegistry/wsConfig.txt'), [Text.UTF8Encoding]::new($false, $true))
$config = @{}
foreach ($line in $configRows) {
    $fields = $line.Split("`t")
    if ($fields.Count -ne 3 -or $config.ContainsKey($fields[1])) {
        throw 'Invalid or duplicate configuration row.'
    }
    $config.Add($fields[1], $fields[2])
}

function Get-Config([string]$Key) {
    $fullKey = 'PeopleRegistry::' + $Key
    if (-not $config.ContainsKey($fullKey)) {
        throw "Missing configuration key: $fullKey"
    }
    return $config[$fullKey]
}

function Get-Table($Book, [string]$Name) {
    foreach ($sheet in $Book.Worksheets) {
        foreach ($table in $sheet.ListObjects) {
            if ($table.Name -eq $Name) {
                return $table
            }
        }
    }
    throw "Required table missing: $Name"
}

Add-Type -TypeDefinition 'using System; using System.Runtime.InteropServices; public static class PeopleRegistryLayoutRot { [DllImport("oleaut32.dll", PreserveSig=false)] public static extern void GetActiveObject(ref Guid clsid, IntPtr reserved, [MarshalAs(UnmanagedType.IUnknown)] out object value); }'
$excel = $null
$book = $null
$ownsExcel = $false
$ownsBook = $false
$guarded = $false
$historyProtectionReady = $false
try {
    try {
        $classId = [Guid]'00024500-0000-0000-C000-000000000046'
        [PeopleRegistryLayoutRot]::GetActiveObject([ref]$classId, [IntPtr]::Zero, [ref]$excel)
    } catch {
        $excel = $null
    }
    if ($null -ne $excel) {
        foreach ($candidate in $excel.Workbooks) {
            if ($candidate.FullName -ieq $bookPath) {
                $book = $candidate
                break
            }
        }
    }
    if ($null -eq $book) {
        $excel = New-Object -ComObject Excel.Application
        $ownsExcel = $true
        $excel.Visible = $false
        $excel.AutomationSecurity = 3
        $excel.EnableEvents = $false
        $book = $excel.Workbooks.Open($bookPath, 0, $false)
        $ownsBook = $true
    }
    if ($book.ReadOnly -or $book.VBProject.Mode -ne 2) {
        throw 'The generator must be writable and its VBA project must be idle.'
    }
    $oldEvents = $excel.EnableEvents
    $oldUpdating = $excel.ScreenUpdating
    $oldAlerts = $excel.DisplayAlerts
    $guarded = $true
    $excel.EnableEvents = $false
    $excel.ScreenUpdating = $false
    $excel.DisplayAlerts = $false

    $backupFolder = [IO.Path]::GetFullPath((Join-Path $book.Path (Get-Config 'Layout.BackupFolder')))
    [void][IO.Directory]::CreateDirectory($backupFolder)
    $backupPath = Join-Path $backupFolder ('PeopleRegistry.before-source-' + (Get-Date -Format 'yyyyMMdd-HHmmss') + '.xlsm')
    $book.SaveCopyAs($backupPath)
    $counts = @{}
    foreach ($kind in @('History')) {
        $table = Get-Table $book (Get-Config "Table.$kind")
        $counts[$kind] = $table.ListRows.Count
    }

    $service = $book.Worksheets.Item('wsConfig')

    $history = Get-Table $book (Get-Config 'Table.History')
    . (Join-Path $PSScriptRoot 'Set-HistoryProtection.ps1')
    $historyProtectionReady = $true
    Set-HistoryProtection $book $config $false
    $historySheet = $null
    foreach ($sheet in $book.Worksheets) {
        if ($sheet.Name -eq (Get-Config 'Sheet.History')) {
            $historySheet = $sheet
            break
        }
    }
    if ($null -eq $historySheet) {
        $historySheet = $book.Worksheets.Add([Type]::Missing, $book.Worksheets.Item($book.Worksheets.Count))
        $historySheet.Name = Get-Config 'Sheet.History'
    }
    if ($history.Parent.Name -ne $historySheet.Name) {
        $oldTitle = $history.HeaderRowRange.Cells.Item(1, 1).Offset(-2)
        [void]$history.Range.Cut($historySheet.Range((Get-Config 'Layout.HistoryCell')))
        [void]$oldTitle.ClearContents()
        $history = Get-Table $book (Get-Config 'Table.History')
    }
    $historySheet.Visible = -1

    $configTable = $service.ListObjects.Item('tbConfig')
    $preservedForm = $null
    foreach ($row in $configTable.ListRows) {
        if ($row.Range.Cells.Item(1, 2).Value2 -eq 'PeopleRegistry::Form.Cells') {
            $oldAddresses = ([string]$row.Range.Cells.Item(1, 3).Value2).Split('|')
            if ($oldAddresses.Count -notin @(3, 7, 8, 9)) { throw 'Invalid saved form schema.' }
            $preservedForm = [object[]]::new(3)
            $formSheet = $book.Worksheets.Item((Get-Config 'Sheet.Input'))
            foreach ($index in 0..1) { $preservedForm[$index] = $formSheet.Range($oldAddresses[$index]).Value2 }
            if ($oldAddresses.Count -eq 3) {
                $preservedForm[2] = $formSheet.Range($oldAddresses[2]).Value2
            } else {
                $start = 4
                if ($oldAddresses.Count -gt 7) { $start = 5 }
                $preservedForm[2] = (@($start..($start + 2) | ForEach-Object {
                    [string]$formSheet.Range($oldAddresses[$_]).Value2
                }) -join "`n").Trim()
            }
            break
        }
    }
    [void]$configTable.Resize($configTable.HeaderRowRange.Resize($configRows.Count + 1, 3))
    $configTable.DataBodyRange.NumberFormat = '@'
    $matrix = [object[,]]::new($configRows.Count, 3)
    for ($row = 0; $row -lt $configRows.Count; $row++) {
        $fields = $configRows[$row].Split("`t")
        foreach ($column in 0..2) {
            $matrix[$row, $column] = [string]$fields[$column]
        }
    }
    $configTable.DataBodyRange.Value2 = $matrix
    . (Join-Path $PSScriptRoot 'Set-OriginAliases.ps1')
    Set-OriginAliases $book $config
    foreach ($kind in @('History')) {
        $table = Get-Table $book (Get-Config "Table.$kind")
        if ($table.ListRows.Count -ne $counts[$kind]) {
            throw "Migration changed the row count: $kind"
        }
        $headers = (Get-Config "Header.$kind").Split('|')
        if ($headers.Count -ne $table.ListColumns.Count) {
            throw "Unexpected table schema: $kind"
        }
        foreach ($column in 1..$headers.Count) {
            $table.ListColumns.Item($column).Name = $headers[$column - 1]
        }
        $table.HeaderRowRange.Cells.Item(1, 1).Offset(-2).Value2 = Get-Config "Title.$kind"
    }
    function Get-Color([string]$Key) {
        $hex = Get-Config $Key
        if ($hex -notmatch '^#[0-9A-Fa-f]{6}$') {
            throw "Invalid HEX color: $Key"
        }
        return [Convert]::ToInt32($hex.Substring(1, 2), 16) +
            256 * [Convert]::ToInt32($hex.Substring(3, 2), 16) +
            65536 * [Convert]::ToInt32($hex.Substring(5, 2), 16)
    }
    $style = $book.TableStyles.Item('PeopleRegistryDark')
    foreach ($element in @(@(0, 'Theme.Background'), @(1, 'Theme.Header'), @(5, 'Theme.Background'), @(6, 'Theme.Alternate'))) {
        $part = $style.TableStyleElements.Item($element[0])
        $part.Interior.Color = Get-Color $element[1]
        $part.Font.Color = Get-Color 'Theme.Foreground'
        if ($element[0] -eq 5 -or $element[0] -eq 6) { $part.StripeSize = 1 }
    }
    foreach ($sheet in $book.Worksheets) {
        $sheet.Cells.Interior.Color = Get-Color 'Theme.Background'
        $sheet.Cells.Font.Color = Get-Color 'Theme.Foreground'
        $sheet.Cells.Font.Name = Get-Config 'Theme.FontName'
        $sheet.Tab.Color = Get-Color 'Theme.Accent'
        foreach ($table in $sheet.ListObjects) {
            $table.Range.Interior.Pattern = -4142
            $table.TableStyle = 'PeopleRegistryDark'
            $table.ShowTableStyleRowStripes = $true
            $table.ShowTableStyleColumnStripes = $false
            $table.HeaderRowRange.Interior.Color = Get-Color 'Theme.Header'
            $table.HeaderRowRange.Font.Bold = $true
            $table.Range.Borders.Color = Get-Color 'Theme.Border'
            $table.Range.Borders.LineStyle = 1
        }
    }
    . (Join-Path $PSScriptRoot 'Set-GeneratorForm.ps1')
    Set-GeneratorForm $book $config $preservedForm
    $inputSheet = $book.Worksheets.Item((Get-Config 'Sheet.Input'))

    # Import source only. Do not compile, initialize or invoke generation macros.
    . (Join-Path $PSScriptRoot 'Import-Modules.ps1')
    Import-PeopleRegistryModules $book (Split-Path $PSScriptRoot -Parent)
    [void]$inputSheet.Activate()
    Set-HistoryProtection $book $config
    $book.Save()
    'Backup: ' + $backupPath
    'Layout: generator, ID history, origin dictionary and configuration worksheets.'
    foreach ($kind in $counts.Keys) {
        $kind + ' rows preserved: ' + $counts[$kind]
    }
    'Sources imported without VBA compilation or execution; RUH was not opened or modified.'
} finally {
    if ($historyProtectionReady) { Set-HistoryProtection $book $config }
    if ($guarded) {
        $excel.EnableEvents = $oldEvents
        $excel.ScreenUpdating = $oldUpdating
        $excel.DisplayAlerts = $oldAlerts
    }
    if ($ownsBook -and $null -ne $book) {
        $book.Close($false)
    }
    if ($ownsExcel -and $null -ne $excel) {
        $excel.Quit()
        [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($excel)
    }
}