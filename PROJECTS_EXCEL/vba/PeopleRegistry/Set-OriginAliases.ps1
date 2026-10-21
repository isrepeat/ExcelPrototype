function Set-OriginAliases($Book, [hashtable]$Config) {
    function Read-AliasConfig([string]$Key) {
        $name = 'PeopleRegistry::' + $Key
        if (-not $Config.ContainsKey($name)) { throw "Missing configuration key: $name" }
        return $Config[$name]
    }
    $sheet = $null
    foreach ($candidate in $Book.Worksheets) {
        if ($candidate.Name -eq (Read-AliasConfig 'Sheet.Aliases')) { $sheet = $candidate; break }
    }
    if ($null -eq $sheet) {
        $sheet = $Book.Worksheets.Add([Type]::Missing, $Book.Worksheets.Item((Read-AliasConfig 'Sheet.History')))
        $sheet.Name = Read-AliasConfig 'Sheet.Aliases'
    }
    $sheet.Visible = -1
    $name = Read-AliasConfig 'Table.Aliases'
    $table = $null
    foreach ($candidateSheet in $Book.Worksheets) {
        foreach ($candidate in $candidateSheet.ListObjects) {
            if ($candidate.Name -eq $name) { $table = $candidate; break }
        }
        if ($null -ne $table) { break }
    }
    if ($null -ne $table -and $table.Parent.Name -ne $sheet.Name) {
        $oldTitle = $table.HeaderRowRange.Cells.Item(1, 1).Offset(-2)
        $oldHelp = $oldTitle.Offset(1)
        $rowCount = $table.ListRows.Count
        [void]$table.Range.Cut($sheet.Range((Read-AliasConfig 'Layout.AliasesCell')))
        [void]$oldTitle.ClearContents()
        [void]$oldHelp.ClearContents()
        $table = $sheet.ListObjects.Item($name)
        if ($table.ListRows.Count -ne $rowCount) { throw 'Moving the dictionary changed its row count.' }
    }
    $headers = (Read-AliasConfig 'Header.Aliases').Split('|')
    if ($headers.Count -ne 2) { throw 'The origin dictionary requires two columns.' }
    if ($null -eq $table) {
        $start = $sheet.Range((Read-AliasConfig 'Layout.AliasesCell'))
        $start.Value2 = $headers[0]
        $start.Offset(0, 1).Value2 = $headers[1]
        $table = $sheet.ListObjects.Add(1, $start.Resize(2, 2), [Type]::Missing, 1)
        $table.Name = $name
        $table.TableStyle = ''
    }
    if ($table.ListColumns.Count -ne 2) { throw 'Invalid origin dictionary schema.' }
    foreach ($index in 0..1) { $table.ListColumns.Item($index + 1).Name = $headers[$index] }
    $table.Range.NumberFormat = '@'
    $table.Range.WrapText = $true
    $title = $table.HeaderRowRange.Cells.Item(1, 1).Offset(-2)
    $title.Value2 = Read-AliasConfig 'Title.Aliases'
    $title.Font.Bold = $true
    $help = $title.Offset(1)
    $help.Value2 = Read-AliasConfig 'Help.Aliases'
    $help.WrapText = $true
    if ([bool]::Parse((Read-AliasConfig 'Layout.Aliases.AutoFitRows'))) {
        [void]$help.EntireRow.AutoFit()
        [void]$table.Range.EntireRow.AutoFit()
    }
}