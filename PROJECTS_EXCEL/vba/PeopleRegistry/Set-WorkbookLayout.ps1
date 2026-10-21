function Set-WorkbookLayout($Book, [hashtable]$Config) {
    function Read-LayoutConfig([string]$Key) {
        $name = 'PeopleRegistry::' + $Key
        if (-not $Config.ContainsKey($name)) { throw "Missing configuration key: $name" }
        [string]$Config[$name]
    }
    $roles = @('Input','History','Aliases','Config')
    foreach ($role in $roles) {
        $sheetName = 'wsConfig'
        if ($role -ne 'Config') { $sheetName = Read-LayoutConfig "Sheet.$role" }
        $sheet = $Book.Worksheets.Item($sheetName)
        $sheet.Columns.ColumnWidth = [double](Read-LayoutConfig 'Layout.DefaultColumnWidth')
        $sheet.Rows.RowHeight = [double](Read-LayoutConfig "Layout.$role.DefaultRowHeight")
        $sheet.Cells.Font.Size = [double](Read-LayoutConfig 'Theme.FontSize')
        foreach ($entry in (Read-LayoutConfig "Layout.$role.ColumnWidths").Split('|')) {
            $parts = $entry.Split('=')
            $sheet.Columns.Item($parts[0]).ColumnWidth = [double]$parts[1]
        }
        if ($role -in @('Aliases','Config') -and [bool]::Parse((Read-LayoutConfig "Layout.$role.AutoFitRows"))) {
            foreach ($table in $sheet.ListObjects) { [void]$table.Range.EntireRow.AutoFit() }
            if ($role -eq 'Aliases') { [void]$sheet.Rows.Item(2).AutoFit() }
        }
    }
    $sheet = $Book.Worksheets.Item((Read-LayoutConfig 'Sheet.Input'))
    foreach ($entry in (Read-LayoutConfig 'Layout.Input.RowHeights').Split('|')) {
        $parts = $entry.Split('=')
        $sheet.Rows.Item([int]$parts[0]).RowHeight = [double]$parts[1]
    }
    $addresses = (Read-LayoutConfig 'Form.Cells').Split('|')[0..1] + (Read-LayoutConfig 'Form.ReplaceCells').Split('|')
    foreach ($address in $addresses) { $sheet.Range($address).EntireRow.RowHeight = [double](Read-LayoutConfig 'Layout.FieldHeight') }
    $view = $sheet.Range((Read-LayoutConfig 'Layout.Input.ViewRange'))
    $view.WrapText = $true
    $view.VerticalAlignment = -4108
    $view.HorizontalAlignment = [int](Read-LayoutConfig 'Layout.Input.ContentAlignment')
    $sheet.Range((Read-LayoutConfig 'Layout.ResultRange')).VerticalAlignment = -4160
    foreach ($address in (Read-LayoutConfig 'Form.Labels').Split('|') + (Read-LayoutConfig 'Form.ReplaceLabels').Split('|')) {
        $sheet.Range($address).HorizontalAlignment = [int](Read-LayoutConfig 'Layout.Input.LabelAlignment')
    }
    $info = (Read-LayoutConfig 'Form.InfoCells').Split('|')
    foreach ($i in @(0,2,4)) { $sheet.Range($info[$i]).Font.Size = [double](Read-LayoutConfig 'Theme.TitleFontSize') }
    $sheet.PageSetup.PrintArea = Read-LayoutConfig 'Layout.Input.PrintArea'
    $sheet.PageSetup.PaperSize = [int](Read-LayoutConfig 'Layout.Input.PaperSize')
    $sheet.PageSetup.Orientation = [int](Read-LayoutConfig 'Layout.Input.Orientation')
    $sheet.PageSetup.Zoom = $false
    $sheet.PageSetup.FitToPagesWide = 1
    $sheet.PageSetup.FitToPagesTall = 1
    [void]$sheet.Activate()
    $window = $Book.Application.ActiveWindow
    $window.Zoom = [int](Read-LayoutConfig 'Layout.Zoom')
    $window.DisplayGridlines = $false
    $window.FreezePanes = $false
    $scroll = $sheet.Range((Read-LayoutConfig 'Layout.Input.ScrollCell'))
    $window.ScrollRow = $scroll.Row
    $window.ScrollColumn = $scroll.Column
    [void]$sheet.Range((Read-LayoutConfig 'Layout.Input.ActiveCell')).Select()
}