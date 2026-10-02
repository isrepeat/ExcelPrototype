Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'
$sourceRoot = [IO.Path]::GetFullPath((Join-Path $PSScriptRoot '../vba/common'))
$fixtureRoot = Join-Path ([IO.Path]::GetTempPath()) ('TableQueries-' + [Guid]::NewGuid().ToString('N'))
[IO.Directory]::CreateDirectory($fixtureRoot) | Out-Null
$excel = $null
$runner = $null
$dataBook = $null
try {
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    $excel.EnableEvents = $false
    $excel.AutomationSecurity = 1
    $dataPath = Join-Path $fixtureRoot 'Data.xlsx'
    $dataBook = $excel.Workbooks.Add()
    $sheet = $dataBook.Worksheets.Item(1)
    $sheet.Name = 'Data'
    $headers = @('ID', 'Name', 'Amount', 'When', 'Flag', 'Notes')
    for ($col = 1; $col -le $headers.Count; $col++) { $sheet.Cells.Item(1, $col).Value2 = $headers[$col - 1] }
    for ($row = 2; $row -le 5; $row++) {
        $sheet.Cells.Item($row, 1).Value2 = [double]($row - 1)
        $sheet.Cells.Item($row, 2).Value2 = if ($row -eq 2) { ' Alice ' } elseif ($row -eq 3) { 'alice' } elseif ($row -eq 4) { 'Bob' } else { 'Carol' }
        $sheet.Cells.Item($row, 3).Value2 = [double](@(2, 10, 30, 20)[$row - 2])
        $sheet.Cells.Item($row, 4).Value2 = [double](46000 + $row)
        $sheet.Cells.Item($row, 5).Formula = if ($row % 2 -eq 0) { '=TRUE()' } else { '=FALSE()' }
        if ($row -eq 2) { $sheet.Cells.Item($row, 6).Value2 = 'literal %_[ token' }
    }
    $table = $sheet.ListObjects.Add(1, $sheet.Range('A1:F5'), $null, 1)
    $table.Name = 'tblData'
    $table.ShowTotals = $true
    $dataBook.SaveAs($dataPath, 51)
    $dataBook.Close($false)
    $dataBook = $null
    $runner = $excel.Workbooks.Add()
    $names = @('obj_TableSource', 'obj_TableQuery', 'obj_QueryCondition', 'obj_QueryGroup', 'obj_QueryOrder',
        'obj_DataTable', 'obj_QueryWorkbookSession', 'obj_ExcelQueryExecutor', 'obj_AdoQueryExecutor',
        'obj_TableQueryService', 'obj_ITableQueryExecutor', 'obj_UiRawTable', 'obj_IUiTableSource', 'ex_TableQuery')
    foreach ($file in Get-ChildItem -LiteralPath $sourceRoot -Recurse -Filter '*.vba') {
        $source = [IO.File]::ReadAllText($file.FullName, [Text.Encoding]::UTF8)
        $match = [regex]::Match($source, '(?m)^Attribute VB_Name = "([^"]+)"')
        if (-not $match.Success -or $match.Groups[1].Value -notin $names) { continue }
        $name = $match.Groups[1].Value
        $source = [regex]::Replace($source, '(?s)^VERSION 1\.0 CLASS\s+BEGIN.*?END\s*', '')
        $source = [regex]::Replace($source, '(?m)^Attribute .*\r?\n?', '')
        $component = $runner.VBProject.VBComponents.Add($(if ($file.Name.EndsWith('.cls.vba')) { 2 } else { 1 }))
        $component.Name = $name
        $component.CodeModule.AddFromString($source)
    }
    $probe = $runner.VBProject.VBComponents.Add(1)
    $probe.Name = 'ex_QueryProbe'
    $probe.CodeModule.AddFromString(@'
Option Explicit
Private Sub Assert(ByVal ok As Boolean, ByVal message As String)
    If Not ok Then Err.Raise vbObjectError + 2200, , message
End Sub
Public Function Run(ByVal path As String) As String
    Dim source As New obj_TableSource, query As obj_TableQuery, service As New obj_TableQueryService
    Dim result As obj_DataTable, ui As obj_UiRawTable, diagnostic As String
    Dim backend As Long, book As Workbook, condition As obj_QueryCondition, group As obj_QueryGroup
    Dim executor As obj_ITableQueryExecutor
    Dim excelExecutor As obj_ExcelQueryExecutor
    On Error GoTo EH
    Assert source.Initialize(), "Source initialization"
    Set query = New obj_TableQuery
    Assert Not service.TryExecute(source, query, result, diagnostic), "Uninitialized service accepted query"
    Assert service.Initialize(), "Service initialization"
    source.WorkbookPath = path: source.TableName = "tblData"
    For backend = QueryExcel To QueryAdo
        service.Backend = backend
        Set query = New obj_TableQuery
    Assert query.Initialize(), "Query initialization"
        query.AddColumn "ID": query.AddColumn "Amount"
        query.AddCondition "Amount", QueryGreaterOrEqual, 10, QueryNumber
        query.AddOrderBy "Amount", True, QueryNumber
        query.Limit = 2
        Assert service.TryExecute(source, query, result, diagnostic), diagnostic
        Assert result.RowCount = 2 And result.ColumnCount = 2, "Projection/limit"
        Assert result.ValueAt(1, 1) = 3 And result.ValueAt(2, 1) = 4, "Numeric sorting"
        Assert IsNumeric(result.ValueAt(1, 2)) And VarType(result.ValueAt(1, 2)) <> vbString, "Numeric values preserved"
        Set query = New obj_TableQuery
    Assert query.Initialize(), "Query initialization"
        query.AddCondition "Name", QueryEquals, "ALICE", QueryText, True
        Assert service.TryExecute(source, query, result, diagnostic), diagnostic
        Assert result.RowCount = 2, "Normalized equality"
        Set query = New obj_TableQuery
    Assert query.Initialize(), "Query initialization"
        query.AddCondition "Name", QueryEquals, "Alice", QueryText, True, True
        Assert service.TryExecute(source, query, result, diagnostic), diagnostic
        Assert result.RowCount = 1, "Case-sensitive equality"
        Set query = New obj_TableQuery
    Assert query.Initialize(), "Query initialization"
        query.AddCondition "Notes", QueryContains, "%_["
        Assert service.TryExecute(source, query, result, diagnostic), diagnostic
        Assert result.RowCount = 1, "Literal contains"
        Set query = New obj_TableQuery
    Assert query.Initialize(), "Query initialization"
        query.AddCondition "When", QueryGreaterOrEqual, CDate(46004), QueryDate
        Assert service.TryExecute(source, query, result, diagnostic), diagnostic
        Assert result.RowCount = 2, "Date comparison"
        Set query = New obj_TableQuery
    Assert query.Initialize(), "Query initialization"
        query.AddCondition "Flag", QueryEquals, True, QueryBoolean
        Assert service.TryExecute(source, query, result, diagnostic), diagnostic
        Assert result.RowCount = 2, "Boolean comparison"
        Set query = New obj_TableQuery
    Assert query.Initialize(), "Query initialization"
        query.AddCondition "Notes", QueryIsEmpty
        Assert service.TryExecute(source, query, result, diagnostic), diagnostic
        Assert result.RowCount = 3, "Empty cells"
        Set query = New obj_TableQuery
    Assert query.Initialize(), "Query initialization"
        Set group = New obj_QueryGroup
        Assert group.Initialize(), "Group initialization"
        group.MatchAny = True
        Set condition = New obj_QueryCondition
        Assert condition.Initialize(), "Condition initialization"
        condition.ColumnName = "ID": condition.Operation = QueryEquals: condition.ValueType = QueryNumber: condition.Value = 1
        group.AddCondition condition
        Set condition = New obj_QueryCondition
        Assert condition.Initialize(), "Condition initialization"
        condition.ColumnName = "ID": condition.Operation = QueryEquals: condition.ValueType = QueryNumber: condition.Value = 4
        group.AddCondition condition
        query.Filter.AddGroup group
        Assert service.TryExecute(source, query, result, diagnostic), diagnostic
        Assert result.RowCount = 2, "OR group"
        Set query = New obj_TableQuery
    Assert query.Initialize(), "Query initialization"
        query.AddCondition "ID", QueryEquals, 999, QueryNumber
        Assert service.TryExecute(source, query, result, diagnostic), diagnostic
        Assert result.RowCount = 0 And result.ColumnCount = 6, "Empty result preserves schema"
        Assert result.TryCreateUiTable(ui, diagnostic), diagnostic
        Assert ui.RowCount = 0 And ui.ColumnCount = 6, "Empty UI adapter"
        query.AddColumn "Missing"
        Assert Not service.TryExecute(source, query, result, diagnostic), "Missing column accepted"
        Assert result Is Nothing And Len(diagnostic) > 0, "Missing-column diagnostic"
    Next backend
    Set book = Application.Workbooks.Open(path)
    book.Worksheets("Data").Cells(2, 2).Value2 = "Unsaved"
    service.Backend = QueryAuto
    Set query = New obj_TableQuery
    Assert query.Initialize(), "Query initialization"
    query.AddCondition "Name", QueryEquals, "Unsaved"
    Assert service.TryExecute(source, query, result, diagnostic), diagnostic
    Assert result.RowCount = 1, "Auto did not read unsaved values"
    service.Backend = QueryAdo
    Assert service.TryExecute(source, query, result, diagnostic), diagnostic
    Assert result.RowCount = 0, "ADO did not read saved file"
    Assert Not book.Saved, "Query changed user workbook state"
    book.Close False
    source.TableName = "": source.SheetName = "Data": source.RangeAddress = "A1:A2"
    Set query = New obj_TableQuery
    Assert query.Initialize(), "Query initialization"
    For backend = QueryExcel To QueryAdo
        service.Backend = backend
        Assert service.TryExecute(source, query, result, diagnostic), diagnostic
        Assert result.RowCount = 1 And result.ColumnCount = 1, "Single-cell matrix"
    Next backend
    service.Dispose
    Assert Not service.TryExecute(source, query, result, diagnostic), "Disposed service accepted query"
    Assert Not service.Initialize(), "Disposed service reinitialized"
    service.Dispose
    Set service = New obj_TableQueryService
    Assert service.Initialize(), "New service initialization"
    Assert Not service.Initialize(), "Service initialized twice"
    Assert service.TryExecute(source, query, result, diagnostic), diagnostic
    Set excelExecutor = New obj_ExcelQueryExecutor
    Set executor = excelExecutor
    excelExecutor.Dispose
    Assert Not executor.TryExecute(source, query, result, diagnostic), "Disposed executor accepted query"
    Assert Not excelExecutor.Initialize(), "Disposed executor reinitialized"
    Set excelExecutor = New obj_ExcelQueryExecutor
    Set executor = excelExecutor
    Assert Not executor.TryExecute(source, query, result, diagnostic), "Uninitialized executor accepted query"
    Assert excelExecutor.Initialize(), "Concrete executor initialization"
    Assert Not excelExecutor.Initialize(), "Executor initialized twice"
    Assert executor.TryExecute(source, query, result, diagnostic), diagnostic
    query.Dispose
    Assert Not query.Initialize(), "Disposed query reinitialized"
    query.Dispose
    Set query = New obj_TableQuery
    Assert query.Initialize(), "New query initialization"
    Assert Not query.Initialize(), "Query initialized twice"
    Assert query.Columns.Count = 0 And query.Orders.Count = 0 And query.Limit = 0, "Query reset"
    source.Dispose
    Assert Not source.Initialize(), "Disposed source reinitialized"
    Set source = New obj_TableSource
    Assert source.Initialize(), "New source initialization"
    Assert Not source.Initialize(), "Source initialized twice"
    Assert Len(source.WorkbookPath) = 0 And Len(source.TableName) = 0, "Source reset"
    Run = "PASS: Excel/ADO parity, typed filters, OR, sorting, limits, totals exclusion, empty UI results, missing columns, live/saved state, scalar ranges and disposal."
    Exit Function
EH:
    Run = "FAIL backend=" & CStr(backend) & ": " & Err.Description
End Function
'@)
    $runnerPath = Join-Path $fixtureRoot 'Runner.xlsm'
    $runner.SaveAs($runnerPath, 52)
    try { $result = $excel.Run("'Runner.xlsm'!ex_QueryProbe.Run", $dataPath) }
    catch {
        $pane = $excel.VBE.ActiveCodePane
        $startLine = 0; $startColumn = 0; $endLine = 0; $endColumn = 0
        $pane.GetSelection([ref]$startLine, [ref]$startColumn, [ref]$endLine, [ref]$endColumn)
        Write-Output ($pane.CodeModule.Name + ':' + $startLine + ' ' + $pane.CodeModule.Lines($startLine, 1))
        throw
    }
    if (-not $result.StartsWith('PASS:')) { throw $result }
    Write-Output $result
    Write-Output "Fixtures preserved: $fixtureRoot"
}
finally {
    if ($null -ne $dataBook) { $dataBook.Close($false) }
    if ($null -ne $runner) { $runner.Close($false) }
    if ($null -ne $excel) { $excel.Quit() }
}