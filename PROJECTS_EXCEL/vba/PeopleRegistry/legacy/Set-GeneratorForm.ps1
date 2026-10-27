function Set-GeneratorForm($Book, [hashtable]$Config, [object[]]$InputValues = $null) {
    function Read-FormConfig([string]$Key) {
        $keyName = 'PeopleRegistry::' + $Key
        if (-not $Config.ContainsKey($keyName)) { throw "Missing configuration key: $keyName" }
        return $Config[$keyName]
    }
    function Read-FormColor([string]$Key) {
        $hex = Read-FormConfig $Key
        if ($hex -notmatch '^#[0-9A-Fa-f]{6}$') { throw "Invalid color: $Key" }
        return [Convert]::ToInt32($hex.Substring(1, 2), 16) +
            256 * [Convert]::ToInt32($hex.Substring(3, 2), 16) +
            65536 * [Convert]::ToInt32($hex.Substring(5, 2), 16)
    }
    $sheet = $Book.Worksheets.Item((Read-FormConfig 'Sheet.Input'))
    $values = [object[]]::new(3)
    $addresses = (Read-FormConfig 'Form.Cells').Split('|')
    $labels = (Read-FormConfig 'Form.Labels').Split('|')
    $headers = (Read-FormConfig 'Form.Headers').Split('|')
    $replacementAddresses = (Read-FormConfig 'Form.ReplaceCells').Split('|')
    $replacementValues = @($replacementAddresses | ForEach-Object { [string]$sheet.Range($_).Value2 })
    $legacy = $null
    foreach ($table in $sheet.ListObjects) {
        if ($table.Name -eq 'tbInput') { $legacy = $table }
    }
    if ($null -ne $InputValues) {
        if ($InputValues.Count -ne 3) { throw 'The preserved form must have three values.' }
        $values = $InputValues
    } elseif ($null -ne $legacy) {
        $populated = 0
        foreach ($row in $legacy.ListRows) {
            if ($Book.Application.CountA($row.Range.Cells.Item(1, 1).Resize(1, 5)) -gt 0) {
                $populated++
                foreach ($column in 1..2) { $values[$column - 1] = $row.Range.Cells.Item(1, $column).Value2 }
                $values[2] = (@(6..8 | ForEach-Object { [string]$row.Range.Cells.Item(1, $_).Value2 }) -join "`n").Trim()
            }
        }
        if ($populated -gt 1) { throw 'Keep one input person before migrating to the single-person form. All original rows remain in the backup.' }
        [void]$legacy.Delete()
    } else {
        foreach ($index in 0..2) { $values[$index] = $sheet.Range($addresses[$index]).Value2 }
    }
    [void]$sheet.Cells.UnMerge()
    [void]$sheet.Cells.Validation.Delete()
    [void]$sheet.Cells.Clear()
    $sheet.Cells.Interior.Color = Read-FormColor 'Theme.Background'
    $sheet.Cells.Font.Color = Read-FormColor 'Theme.Foreground'
    $sheet.Cells.Font.Name = Read-FormConfig 'Theme.FontName'
    $sheet.Cells.Font.Size = [double](Read-FormConfig 'Theme.FontSize')
    . (Join-Path $PSScriptRoot 'Set-WorkbookLayout.ps1')
    Set-WorkbookLayout $Book $Config
    foreach ($index in 0..2) {
        $sheet.Range($labels[$index]).Value2 = [string]$headers[$index]
        $sheet.Range($labels[$index]).Font.Bold = $true
        $field = $sheet.Range($addresses[$index])
        $field.NumberFormat = '@'
        $field.Interior.Color = Read-FormColor 'Theme.Header'
        if ($index -le 1) { $field.Interior.Color = Read-FormColor 'Theme.Alternate' }
        $field.Borders.Color = Read-FormColor 'Theme.Border'
        $field.Borders.LineStyle = 1
        $sheet.Range($addresses[$index]).Value2 = [string]$values[$index]
    }
    foreach ($address in $addresses) { $sheet.Range($address).EntireRow.RowHeight = [double](Read-FormConfig 'Layout.FieldHeight') }
    $result = $sheet.Range((Read-FormConfig 'Layout.ResultRange'))
    [void]$result.Merge()
    $result.Interior.Color = Read-FormColor 'Theme.Header'
    $result.Borders.Color = Read-FormColor 'Theme.Border'
    $result.Borders.LineStyle = 1
    $result.VerticalAlignment = -4160
    $infoCells = (Read-FormConfig 'Form.InfoCells').Split('|')
    $infoKeys = @('Title.Input', 'Help.Input', 'Title.Result', 'Help.Result', 'Title.Replace', 'Help.Replace')
    foreach ($index in 0..5) {
        $sheet.Range($infoCells[$index]).Value2 = Read-FormConfig $infoKeys[$index]
        $sheet.Range($infoCells[$index]).Font.Color = Read-FormColor 'Theme.Muted'
    }
    foreach ($index in @(0, 2, 4)) {
        $sheet.Range($infoCells[$index]).Font.Size = [double](Read-FormConfig 'Theme.TitleFontSize')
        $sheet.Range($infoCells[$index]).Font.Bold = $true
        $sheet.Range($infoCells[$index]).Font.Color = Read-FormColor 'Theme.Accent'
    }
    for ($shapeIndex = $sheet.Shapes.Count; $shapeIndex -ge 1; $shapeIndex--) {
        if ($sheet.Shapes.Item($shapeIndex).Name -in @('btnPreview','btnClear')) { [void]$sheet.Shapes.Item($shapeIndex).Delete() }
    }
    $replacementLabels = (Read-FormConfig 'Form.ReplaceLabels').Split('|')
    $replacementHeaders = (Read-FormConfig 'Form.ReplaceHeaders').Split('|')
    foreach ($index in 0..1) {
        $sheet.Range($replacementLabels[$index]).Value2 = [string]$replacementHeaders[$index]
        $sheet.Range($replacementLabels[$index]).Font.Bold = $true
        $field = $sheet.Range($replacementAddresses[$index])
        $field.NumberFormat = '@'
        $field.Value2 = [string]$replacementValues[$index]
        $field.EntireRow.RowHeight = [double](Read-FormConfig 'Layout.FieldHeight')
        $field.Interior.Color = Read-FormColor 'Theme.Alternate'
        $field.Borders.Color = Read-FormColor 'Theme.Border'
        $field.Borders.LineStyle = 1
    }
    $buttonNames = @('btnSave','btnReplace')
    $commands = @('SaveIDs','ReplaceID')
    $captionKeys = @('Button.Save','Button.Replace')
    $anchors = (Read-FormConfig 'Layout.ButtonCells').Split('|')
    foreach ($index in 0..($buttonNames.Count - 1)) {
        $button = $null
        foreach ($candidate in $sheet.Shapes) {
            if ($candidate.Name -eq $buttonNames[$index]) { $button = $candidate; break }
        }
        if ($null -eq $button) {
            $button = $sheet.Shapes.AddShape(5, 0, 0, 1, 1)
            $button.Name = $buttonNames[$index]
        }
        $button.Left = $sheet.Range($anchors[$index]).Left
        $button.Top = $sheet.Range($anchors[$index]).Top
        $button.Width = [double](Read-FormConfig 'Layout.ButtonWidth')
        $button.Height = [double](Read-FormConfig 'Layout.ButtonHeight')
        $button.AutoShapeType = 1
        $button.Fill.ForeColor.RGB = Read-FormColor 'Theme.Alternate'
        $button.Line.ForeColor.RGB = Read-FormColor 'Theme.Border'
        $button.TextFrame2.TextRange.Text = Read-FormConfig $captionKeys[$index]
        $button.TextFrame2.TextRange.Font.Size = [double](Read-FormConfig 'Theme.ButtonFontSize')
        $button.TextFrame2.TextRange.Font.Name = Read-FormConfig 'Theme.FontName'
        $button.TextFrame2.TextRange.Font.Bold = -1
        $button.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = Read-FormColor 'Theme.Accent'
        $button.OnAction = "'" + $Book.Name.Replace("'", "''") + "'!ex_PeopleRegistry." + $commands[$index]
        $button.Placement = 3
    }
    [void]$sheet.Activate()
    $Book.Application.ActiveWindow.DisplayGridlines = $false
    $Book.Application.ActiveWindow.FreezePanes = $false
    $Book.Application.ActiveWindow.Zoom = [int](Read-FormConfig 'Layout.Zoom')
    $scroll = $sheet.Range((Read-FormConfig 'Layout.Input.ScrollCell'))
    $Book.Application.ActiveWindow.ScrollRow = $scroll.Row
    $Book.Application.ActiveWindow.ScrollColumn = $scroll.Column
    [void]$sheet.Range((Read-FormConfig 'Layout.Input.ActiveCell')).Select()
    'Generator form prepared; input fields preserved.'
}