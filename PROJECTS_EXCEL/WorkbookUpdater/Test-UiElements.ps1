param()

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'
$projectRoot = [IO.Path]::GetFullPath((Join-Path $PSScriptRoot '..'))
$sourceRoot = Join-Path $projectRoot 'vba'
$manifest = Get-Content -LiteralPath (Join-Path $sourceRoot 'modules.json') -Raw | ConvertFrom-Json
$profileSources = [System.Collections.Generic.List[System.IO.FileInfo]]::new()
foreach ($sourceFile in Get-ChildItem -LiteralPath $sourceRoot -Recurse -Filter '*.vba') {
    $relativePath = [IO.Path]::GetRelativePath($sourceRoot, $sourceFile.FullName).Replace([char]92, [char]47)
    $covered = $false
    foreach ($pattern in $manifest.PersonalEventBuilder) {
        if ($relativePath -like $pattern) { $covered = $true; break }
    }
    if ($covered) { $profileSources.Add($sourceFile) }
}
$fixturePath = Join-Path ([IO.Path]::GetTempPath()) ('UiElements-' + [guid]::NewGuid().ToString('N') + '.xlsm')
$excel = $null
$book = $null

try {
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    $excel.EnableEvents = $false
    $book = $excel.Workbooks.Add()
    $book.Worksheets.Item(1).Name = 'MainPage'
    $configSheet = $book.Worksheets.Add()
    $configSheet.Name = 'wsConfig'
    $configSheet.Cells.Item(1, 1).Value2 = 'Key'
    $configSheet.Cells.Item(1, 2).Value2 = 'Value'
    $configRow = 2
    foreach ($line in Get-Content -LiteralPath (Join-Path $projectRoot 'config/PersonalEventBuilder/wsConfig.txt') -Encoding utf8) {
        $parts = $line -split '\t|\\t'
        if ($parts.Count -eq 3 -and $parts[0] -eq 'value') {
            $configSheet.Cells.Item($configRow, 1).Value2 = $parts[1]
            $configSheet.Cells.Item($configRow, 2).Value2 = $parts[2]
            $configRow++
        }
    }
    $configTable = $configSheet.ListObjects.Add(1, $configSheet.Range('A1:B' + ($configRow - 1)), $null, 1)
    $configTable.Name = 'tbConfig'

    foreach ($file in $profileSources) {
        if ($file.Name -eq 'ThisWorkbook.vba') { continue }
        $source = [IO.File]::ReadAllText($file.FullName)
        $nameMatch = [regex]::Match($source, '(?m)^Attribute VB_Name = "([^"]+)"')
        $name = if ($nameMatch.Success) { $nameMatch.Groups[1].Value } else { $file.BaseName -replace '\.cls$', '' }
        $source = [regex]::Replace($source, '(?s)^VERSION 1.0 CLASS\r?\nBEGIN.*?\r?\nEND\r?\n', '')
        $source = [regex]::Replace($source, '(?m)^Attribute [^\r\n]+\r?\n?', '')
        $type = if ($file.Name.EndsWith('.cls.vba')) { 2 } else { 1 }
        $component = $book.VBProject.VBComponents.Add($type)
        $component.Name = $name
        $component.CodeModule.AddFromString($source)
    }
    $component = $book.VBProject.VBComponents.Add(2)
    $component.Name = 'obj_TestCallbacks'
    $component.CodeModule.AddFromString(@"
Option Explicit
Public Count As Long
Public Function Execute() As Boolean
    Count = Count + 1
    Execute = True
End Function
"@)
    $component = $book.VBProject.VBComponents.Add(1)
    $component.Name = 'ex_UiElementProbe'
    $component.CodeModule.AddFromString(@"
Option Explicit
Public Function Run(ByVal uiFolder As String) As String
    Dim bindingContext As New obj_UiBindingContext
    Dim context As New obj_UiRenderContext
    Dim definition As obj_UiPageDefinition
    Dim callbacks As New obj_TestCallbacks
    Dim command As New obj_UiCommand
    Dim tables As New obj_UiRawTableList
    Dim name As Variant
    Dim diagnostic As String
    Dim snapshot As String
    Dim errors As Collection
    Dim sheet As Worksheet
    Dim shapeCount As Long
    Dim otherBinding As New obj_UiBindingContext
    Dim otherContext As New obj_UiRenderContext
    Dim otherDefinition As New obj_UiPageDefinition
    Dim otherDocument As Object
    Dim otherSheet As Worksheet
    Dim tagFactory As New obj_UiElementFactory
    Dim controlFactory As New obj_UiControlFactory
    Dim lookupProfile As New obj_LookupProfile
    Dim lookupService As New obj_LookupService
    Dim lookupResult As obj_LookupResult
    Dim resolvedSource As String
    Dim resolvedPath As String
    Dim phase As String
    Dim generatedTable As obj_UiRawTable
    Dim generatedValues As Variant
    Dim generatedIndex As Long
    Dim generatedRow As Long

    On Error GoTo EH
    phase = "markup validation"
    tagFactory.Initialize "stackpanel"
    ex_UiElementFactory.fn_Register "flow", tagFactory
    controlFactory.Initialize "form"
    ex_UiControlFactory.fn_Register "DraftForm", controlFactory
    CheckMarkup "<page xmlns='urn:excelprototype:profiles' xmlns:controls='urn:excelprototype:controls'><grid><controls:label name='Bad' text='Title' columnSapn='2'/></grid></page>", False, "columnSapn"
    CheckMarkup "<page xmlns='urn:excelprototype:profiles' xmlns:controls='urn:excelprototype:controls'><grid><stackPanel/></grid></page>", False, "orientation"
    CheckMarkup "<page xmlns='urn:excelprototype:profiles' xmlns:controls='urn:excelprototype:controls'><grid><page xmlns='urn:excelprototype:profiles'/></grid></page>", False, "page"
    CheckMarkup "<page xmlns='urn:excelprototype:profiles' xmlns:controls='urn:excelprototype:controls'><grid><controls:button name='Bad' caption='Title'><grid/></controls:button></grid></page>", False, "Child element"
    CheckMarkup "<page xmlns='urn:excelprototype:profiles' xmlns:controls='urn:excelprototype:controls'><grid><controls:form name='Bad' dataContext='{Binding Path=Form}' orientation='vertical'><grid/></controls:form></grid></page>", False, "Child element"
    CheckMarkup "<page xmlns='urn:excelprototype:profiles' xmlns:controls='urn:excelprototype:controls'><grid><controls:unknown name='Bad'/></grid></page>", False, "Unknown Control"
    CheckMarkup "<page xmlns='urn:excelprototype:profiles' xmlns:controls='urn:excelprototype:controls'><grid><controls:label name='Bad' text='Title' row='1.5'/></grid></page>", False, "row"
    CheckMarkup "<page xmlns='urn:excelprototype:profiles' xmlns:controls='urn:excelprototype:controls'><grid><controls:input name='Bad' value='{Binding Path=Name}' readOnly='yes'/></grid></page>", False, "readOnly"
    CheckMarkup "<page xmlns='urn:excelprototype:profiles' xmlns:controls='urn:excelprototype:controls'><grid/>unexpected text</page>", False, "Text content"
    CheckMarkup "<page xmlns='urn:excelprototype:profiles' xmlns:controls='urn:excelprototype:controls'><grid/><styles/><styles/></page>", False, "Invalid child count"
    CheckMarkup "<page xmlns='urn:excelprototype:profiles' xmlns:controls='urn:excelprototype:controls'><grid><styles/></grid></page>", False, "styles"
    CheckMarkup "<page xmlns='urn:excelprototype:profiles' xmlns:controls='urn:excelprototype:controls'><grid><controls:label name='Bad'/></grid></page>", False, "text|caption"
    CheckMarkup "<page xmlns='urn:excelprototype:profiles' xmlns:controls='urn:excelprototype:controls'><grid><controls:input name='Bad' value='{Binding Path=}'/></grid></page>", False, "value"
    CheckMarkup "<page xmlns='urn:excelprototype:profiles' xmlns:controls='urn:excelprototype:controls'><grid><flow orientation='horizontal'><controls:label name='Good' text='Title'/></flow></grid></page>", True, ""
    CheckMarkup "<page xmlns='urn:excelprototype:profiles'/>", False, "Invalid child count"
    CheckMarkup "<page xmlns='urn:excelprototype:profiles' xmlns:controls='urn:excelprototype:controls'><grid/><grid/></page>", False, "Invalid child count"
    CheckMarkup "<page xmlns='urn:excelprototype:profiles' xmlns:controls='urn:excelprototype:controls'><grid/><stackPanel orientation='vertical'/></page>", False, "Child element"
    CheckMarkup "<page xmlns='urn:excelprototype:profiles' xmlns:controls='urn:excelprototype:controls'><grid/><styles><controlStyle/></styles></page>", False, "name"
    CheckMarkup "<page xmlns='urn:excelprototype:profiles' xmlns:controls='urn:excelprototype:controls'><grid/><styles/></page>", True, ""
    CheckMarkup "<page xmlns='urn:excelprototype:profiles' xmlns:c='urn:excelprototype:controls'><grid><c:label name='Alias' text='OK'/></grid></page>", True, ""
    CheckMarkup "<page xmlns='urn:excelprototype:profiles' xmlns:c='urn:other'><grid><c:label name='Bad' text='Title'/></grid></page>", False, "Unknown Element"
    CheckMarkup "<page xmlns='urn:excelprototype:profiles'><grid><form name='Bad'/></grid></page>", False, "Unknown Element"
    CheckMarkup "<page xmlns='urn:excelprototype:profiles'><grid><control type='Label' name='Bad' text='Title'/></grid></page>", False, "Unknown Element"
    CheckMarkup "<page xmlns='urn:excelprototype:profiles' xmlns:c='urn:excelprototype:controls'><grid><c:form name='Bad' dataContext='{Binding Path=Form}' orientation='vertical'><c:field name='Title' label='Title' type='text'/></c:form></grid></page>", False, "Unknown Control"
    CheckMarkup "<page xmlns='urn:excelprototype:profiles' xmlns:c='urn:excelprototype:controls'><grid source='Draft'/></page>", False, "source"
    CheckMarkup "<page xmlns='urn:excelprototype:profiles' xmlns:c='urn:excelprototype:controls'><grid dataContext='Draft'/></page>", False, "dataContext"
    CheckMarkup "<page xmlns='urn:excelprototype:profiles' xmlns:c='urn:excelprototype:controls'><grid><c:tableList name='Old' source='{Binding Path=Data.Tables}'/></grid></page>", False, "source"
    CheckMarkup "<page xmlns='urn:excelprototype:profiles' xmlns:c='urn:excelprototype:controls'><grid><c:form name='MissingValue' dataContext='{Binding Path=Form}' orientation='vertical'><field name='Notes' label='Notes' type='text'/></c:form></grid></page>", False, "value"
    Set sheet = ThisWorkbook.Worksheets("MainPage")
    phase = "PersonalEventBuilder setup"
    bindingContext.Initialize
    bindingContext.SetValue "Text", "Title", "Test page"
    bindingContext.SetValue "Text", "Reset", "Hello"
    bindingContext.SetValue "Text", "UpdatePage", "Update"
    bindingContext.SetValue "Text", "GenerateTables", "Generate"
    bindingContext.SetValue "Form", "EventName", "Initial"
    bindingContext.SetValue "Form", "Category", "Meeting"
    bindingContext.SetValue "Form", "Notes2", vbNullString
    bindingContext.SetValue "Form", "Notes3", vbNullString
    bindingContext.SetValue "Form", "Notes", "Notes"
    bindingContext.SetValue "Data", "EventTypes", "Meeting,Training,Leave"
    bindingContext.SetValue "Resources", "PrimaryButton", "primaryButton"
    bindingContext.SetValue "Resources", "PageTitle", "pageTitle"
    tables.Initialize
    bindingContext.SetObject "Data", "Tables", tables
    If Not lookupProfile.Initialize("PersonalEventBuilder::lookup.Personnel", diagnostic) Then Err.Raise 5, , diagnostic
    If Not lookupService.Initialize(lookupProfile) Then Err.Raise 5, , "Lookup initialize"
    If Not lookupService.TrySearch("", lookupResult, diagnostic) Then Err.Raise 5, , diagnostic
    bindingContext.SetObject "Data", "PersonnelCandidates", lookupResult
    bindingContext.SetValue "Data", "SelectedPersonnel", ""
    bindingContext.SetValue "Text", "PersonnelStatus", "Search"
    For Each name In Array("PersonNameLabel", "PersonIdLabel", "PersonRankLabel", "PersonPositionLabel", "PersonUnitLabel", "RefreshPersonnel")
        bindingContext.SetValue "Text", CStr(name), CStr(name)
    Next name
    For Each name In Array("PersonName", "PersonId", "PersonRank", "PersonPosition", "PersonUnit")
        bindingContext.SetValue "Form", CStr(name), ""
    Next name
    command.Initialize callbacks, "Execute"
    For Each name In Array("ResetCommand", "UpdatePageCommand", "GenerateTablesCommand", _
        "FormChangedCommand", "SubmitFormCommand", "SearchPersonnel", "SelectPersonnel", "RefreshPersonnel")
        bindingContext.SetObject "Commands", CStr(name), command
    Next name
    If Not ex_UiPageLoader.fn_TryLoad(uiFolder & "\MainPage.xaml", definition) Then Err.Raise 5, , "Load"
    snapshot = definition.Document.xml
    If Not context.Initialize(sheet, definition, uiFolder, bindingContext) Then Err.Raise 5, , "Initialize"
    If Not context.Build(diagnostic) Then Err.Raise 5, , diagnostic
    If definition.Document.xml <> snapshot Then Err.Raise 5, , "Page DOM changed during Build"
    context.Styles.BeginPage sheet, definition.Document, uiFolder
    phase = "PersonalEventBuilder first render"
    If Not context.RenderTree(diagnostic) Then Err.Raise 5, , diagnostic
    If sheet.Shapes("btn_UpdatePage").TopLeftCell.Row <> 2 Then Err.Raise 5, , "Update button position"
    If sheet.Shapes("btn_Reset").TopLeftCell.Row <> 4 Then Err.Raise 5, , "Reset button position"
    If sheet.Shapes("btn_GenerateTables").TopLeftCell.Row <> 8 Then Err.Raise 5, , "Generate button position"
    If sheet.Shapes("btn_UpdatePage").TopLeftCell.Column <= 3 Then Err.Raise 5, , "Buttons must be right of form"

    If sheet.Range("C2").Value2 <> "Initial" Then Err.Raise 5, , "Field layout or binding"
    If sheet.Range("C4").MergeArea.Cells.Count <> 1 Then Err.Raise 5, , "Single-cell field"
    bindingContext.SetValue "Text", "Title", "Changed title"
    If sheet.Range("B1").Value2 <> "Changed title" Then Err.Raise 5, , "Label refresh"
    bindingContext.SetValue "Text", "Reset", "Changed button"
    If sheet.Shapes("btn_Reset").TextFrame2.TextRange.Text <> "Changed button" Then Err.Raise 5, , "Button refresh"
    shapeCount = sheet.Shapes.Count
    phase = "PersonalEventBuilder reflow"
    context.InvalidateMeasure
    If Not context.FlushLayout(diagnostic) Then Err.Raise 5, , diagnostic
    If sheet.Shapes.Count <> shapeCount Then Err.Raise 5, , "Shapes leaked after reflow"
    phase = "PersonalEventBuilder generated Smart tables"
    tables.Clear
    For generatedIndex = 1 To 10
        ReDim generatedValues(1 To 3, 1 To 3)
        For generatedRow = 1 To 3
            generatedValues(generatedRow, 1) = "Candidate " & VBA.CStr(generatedIndex) & "." & VBA.CStr(generatedRow)
            generatedValues(generatedRow, 2) = "Group " & VBA.CStr(generatedIndex)
            generatedValues(generatedRow, 3) = "Ready"
        Next generatedRow
        Set generatedTable = New obj_UiRawTable
        If Not generatedTable.Initialize(generatedValues, VBA.Array("Candidate", "Category", "Status"), _
                "Table " & VBA.CStr(generatedIndex)) Then Err.Raise 5, , "Generated table initialization"
        If Not tables.Add(generatedTable) Then Err.Raise 5, , "Generated table list update"
    Next generatedIndex
    context.InvalidateMeasure
    If Not context.FlushLayout(diagnostic) Then Err.Raise 5, , diagnostic
    If sheet.ListObjects.Count <> 3 Then Err.Raise 5, , "Generated Smart table count"
    If sheet.ListObjects("tbGeneratedEvent2").DataBodyRange.Cells(1, 1).Value2 <> "Candidate 2.1" Then _
        Err.Raise 5, , "Generated Smart table 2"
    If sheet.ListObjects("tbGeneratedEvent5").DataBodyRange.Cells(3, 3).Value2 <> "Ready" Then _
        Err.Raise 5, , "Generated Smart table 5"
    If sheet.ListObjects("tbGeneratedEvent8").ListRows.Count <> 3 Then Err.Raise 5, , "Generated Smart table 8"
    bindingContext.SetValue "Form", "EventName", "Updated"
    If sheet.Range("C2").Value2 <> "Updated" Then Err.Raise 5, , "Reactive update"
    sheet.Range("C2").Value2 = "User"
    If Not context.Router.DispatchCells(sheet.Range("C2")) Then Err.Raise 5, , "Cell dispatch"
    If callbacks.Count <> 1 Then Err.Raise 5, , "Change command"
    If Not context.Router.DispatchShape("btn_Reset") Then Err.Raise 5, , "Shape dispatch"
    If callbacks.Count <> 2 Then Err.Raise 5, , "Button command"
    Set errors = New Collection
    If Not context.ValidateForm("EventDraftForm", errors) Then Err.Raise 5, , "Valid form"
    bindingContext.SetValue "Form", "EventName", vbNullString
    Set errors = New Collection
    If context.ValidateForm("EventDraftForm", errors) Or errors.Count <> 1 Then Err.Raise 5, , "Required field"
    If definition.Document.xml <> snapshot Then Err.Raise 5, , "Page DOM changed during Render"
    Set otherSheet = ThisWorkbook.Worksheets.Add()
    phase = "nested controls"
    otherSheet.Name = "OtherPage"
    Set otherDocument = CreateObject("MSXML2.DOMDocument.6.0")
    If Not otherDocument.LoadXML("<page xmlns='urn:excelprototype:profiles' xmlns:controls='urn:excelprototype:controls'><grid><flow orientation='vertical' dataContext='{Binding Path=Draft.Person}'><controls:draftform name='OtherForm' orientation='horizontal'>" & _
        "<field name='Accepted' value='{Binding Path=Accepted}' label='Accepted' type='checkbox' required='true'/>" & _
        "<field name='Name' value='{Binding Path=Name}' label='Name' type='text' readOnly='true'/></controls:draftform><stackPanel orientation='horizontal'><controls:label name='Tail1' text='Left' columnSpan='2'/><controls:label name='Tail2' text='Right' columnSpan='3'/></stackPanel></flow></grid></page>") Then Err.Raise 5, , "Test XML"
    otherDefinition.Initialize otherDocument, "OtherPage.xaml"
    otherBinding.Initialize
    otherBinding.TrySetPathValue "Draft", "Person.Name", "Nested"
    If Not ex_UiBindingRuntime.fn_TryParseBinding("{Binding Path=Person.Name}", _
        "Draft", resolvedSource, resolvedPath, otherBinding) Then Err.Raise 5, , "Nested parse"
    If resolvedSource <> "Draft" Or resolvedPath <> "Person.Name" Then Err.Raise 5, , "Nested source"
    otherBinding.TrySetPathValue "Draft", "Person.Accepted", False
    otherBinding.TrySetPathValue "Draft", "Person.Name", "Read only"
    otherContext.Initialize otherSheet, otherDefinition, uiFolder, otherBinding
    If Not otherContext.Build(diagnostic) Then Err.Raise 5, , diagnostic
    otherContext.Styles.BeginPage otherSheet, otherDocument, uiFolder
    If Not otherContext.RenderTree(diagnostic) Then Err.Raise 5, , diagnostic
    If otherSheet.Range("A2").Value2 <> "Left" Or otherSheet.Range("C2").Value2 <> "Right" Then Err.Raise 5, , "Nested stack layout"
    If otherSheet.Range("D1").Value2 <> "Read only" Then Err.Raise 5, , "Horizontal layout"
    otherBinding.TrySetPathValue "Draft", "Person.Accepted", True
    If otherSheet.Shapes("chk_1").ControlFormat.Value <> xlOn Then Err.Raise 5, , "Checkbox refresh"
    otherSheet.Shapes("chk_1").ControlFormat.Value = xlOff
    If Not otherContext.Router.DispatchShape("chk_1") Then Err.Raise 5, , "Checkbox dispatch"
    If otherSheet.Range("B1").Value2 <> False Then Err.Raise 5, , "Checkbox reverse binding"
    otherSheet.Range("D1").Value2 = "Attempt"
    If Not otherContext.Router.DispatchCells(otherSheet.Range("D1")) Then Err.Raise 5, , "Readonly dispatch"
    If otherSheet.Range("A2").Value2 <> "Left" Or otherSheet.Range("C2").Value2 <> "Right" Then Err.Raise 5, , "Nested stack layout"
    If otherSheet.Range("D1").Value2 <> "Read only" Then Err.Raise 5, , "Readonly restore"
    Set errors = New Collection
    If otherContext.ValidateForm("OtherForm", errors) Or errors.Count <> 1 Then Err.Raise 5, , "Checkbox validation"
    bindingContext.SetValue "Text", "Title", "First page"
    If sheet.Range("B1").Value2 <> "First page" Then Err.Raise 5, , "Page binding isolation"
    If Not context.InvalidateVisual("Reset", diagnostic) Then Err.Raise 5, , diagnostic
    otherContext.Dispose
    otherContext.Dispose
    context.Dispose
    context.Dispose
    bindingContext.SetValue "Form", "EventName", "After disposal"
    If sheet.Range("C2").Value2 <> vbNullString Then Err.Raise 5, , "Subscription survived disposal"
    phase = "grid controls"
    CheckGrid uiFolder
    phase = "raw tables"
    CheckTables uiFolder
    phase = "smart tables"
    CheckSmartTables uiFolder
    Run = "PASS"
    Exit Function
EH:
    Run = "FAIL: " & phase & " | " & Err.Source & ": " & Err.Description
End Function

Private Sub CheckGrid(ByVal uiFolder As String)
    Dim document As Object
    Dim definition As New obj_UiPageDefinition
    Dim binding As New obj_UiBindingContext
    Dim context As New obj_UiRenderContext
    Dim sheet As Worksheet
    Dim diagnostic As String
    Dim tables As New obj_UiRawTableList
    Dim table As New obj_UiRawTable
    Dim values(1 To 1, 1 To 2) As Variant
    Dim phase As String

    On Error GoTo EH_GRID
    phase = "markup"
    CheckMarkup "<page xmlns='urn:excelprototype:profiles'><grid><grid.rowDefinitions><rowDefinition size='bad'/></grid.rowDefinitions></grid></page>", False, "size"
    CheckMarkup "<page xmlns='urn:excelprototype:profiles'><grid><grid.rowDefinitions/></grid></page>", False, "Invalid child count"
    Set document = CreateObject("MSXML2.DOMDocument.6.0")
    If Not document.LoadXML("<page xmlns='urn:excelprototype:profiles' xmlns:c='urn:excelprototype:controls'><grid><grid.columnDefinitions><columnDefinition size='auto'/><columnDefinition size='1'/><columnDefinition size='auto'/></grid.columnDefinitions><grid.rowDefinitions><rowDefinition size='auto'/><rowDefinition size='1'/><rowDefinition size='auto'/></grid.rowDefinitions><c:form name='GridForm' dataContext='{Binding Path=Draft}' orientation='vertical'><field name='Name' value='{Binding Path=Name}' label='Name' type='text'/><field name='Notes' value='{Binding Path=Notes}' label='Notes' type='text'/></c:form><c:label name='Right' text='Right' column='3'/><c:tableList name='Tables' itemsSource='{Binding Path=Data.Tables}' row='3'/></grid></page>") Then Err.Raise 5, , "Grid XML"
    binding.Initialize
    binding.SetValue "Draft", "Name", "Name value"
    binding.SetValue "Draft", "Notes", "Notes value"
    values(1, 1) = "Row"
    values(1, 2) = "Value"
    table.Initialize values, Array("A", "B"), "Table"
    tables.Initialize
    tables.Add table
    binding.SetObject "Data", "Tables", tables
    Set sheet = ThisWorkbook.Worksheets.Add()
    definition.Initialize document, "Grid.xaml"
    context.Initialize sheet, definition, uiFolder, binding
    If Not context.Build(diagnostic) Then Err.Raise 5, , diagnostic
    context.Styles.BeginPage sheet, document, uiFolder
    phase = "first render"
    If Not context.RenderTree(diagnostic) Then Err.Raise 5, , diagnostic
    If sheet.Range("D1").Value2 <> "Right" Or sheet.Range("A4").Value2 <> "Table" Then Err.Raise 5, , "Auto grid positioning"
    If sheet.Range("D1").MergeArea.Rows.Count <> 2 Then Err.Raise 5, , "Grid slot height"
    phase = "second render"
    If Not context.RenderTree(diagnostic) Then Err.Raise 5, , diagnostic
    If sheet.Range("D1").Value2 <> "Right" Then Err.Raise 5, , "Repeated grid measure"
    document.documentElement.firstChild.lastChild.setAttribute "column", "4"
    definition.Initialize document, "OutOfBounds.xaml"
    Dim invalidContext As New obj_UiRenderContext
    invalidContext.Initialize sheet, definition, uiFolder, binding
    If invalidContext.Build(diagnostic) Then Err.Raise 5, , "Grid accepted invalid track index"
    invalidContext.Dispose
    document.documentElement.firstChild.lastChild.setAttribute "column", "1"
    context.Dispose

    phase = "span render"
    If Not document.LoadXML("<page xmlns='urn:excelprototype:profiles' xmlns:c='urn:excelprototype:controls'><grid><grid.columnDefinitions><columnDefinition size='2'/><columnDefinition size='3'/><columnDefinition size='1'/></grid.columnDefinitions><c:label name='Span' text='Span' columnSpan='2'/><c:label name='After' text='After' column='3'/></grid></page>") Then Err.Raise 5, , "Span XML"
    sheet.Cells.UnMerge
    sheet.Cells.ClearContents
    Set definition = New obj_UiPageDefinition
    definition.Initialize document, "Span.xaml"
    Set context = New obj_UiRenderContext
    context.Initialize sheet, definition, uiFolder, binding
    If Not context.Build(diagnostic) Then Err.Raise 5, , diagnostic
    context.Styles.BeginPage sheet, document, uiFolder
    If Not context.RenderTree(diagnostic) Then Err.Raise 5, , diagnostic
    If sheet.Range("A1").MergeArea.Columns.Count <> 5 Or sheet.Range("F1").Value2 <> "After" Then Err.Raise 5, , "Grid track span"
    context.Dispose

    phase = "star render"
    If Not document.LoadXML("<page xmlns='urn:excelprototype:profiles' xmlns:c='urn:excelprototype:controls'><grid rowSpan='2' columnSpan='9'><grid.columnDefinitions><columnDefinition size='*'/><columnDefinition size='2*'/></grid.columnDefinitions><grid.rowDefinitions><rowDefinition size='*'/></grid.rowDefinitions><c:label name='Star' text='Star' column='2'/></grid></page>") Then Err.Raise 5, , "Star XML"
    sheet.Cells.UnMerge
    sheet.Cells.ClearContents
    Set definition = New obj_UiPageDefinition
    definition.Initialize document, "Star.xaml"
    Set context = New obj_UiRenderContext
    context.Initialize sheet, definition, uiFolder, binding
    If Not context.Build(diagnostic) Then Err.Raise 5, , diagnostic
    context.Styles.BeginPage sheet, document, uiFolder
    If Not context.RenderTree(diagnostic) Then Err.Raise 5, , diagnostic
    If sheet.Range("D1").Value2 <> "Star" Or sheet.Range("D1").MergeArea.Columns.Count <> 6 Then Err.Raise 5, , "Weighted star grid"
    context.Dispose

    document.documentElement.firstChild.removeAttribute "columnSpan"
    Set definition = New obj_UiPageDefinition
    definition.Initialize document, "InvalidStar.xaml"
    Set context = New obj_UiRenderContext
    context.Initialize sheet, definition, uiFolder, binding
    If Not context.Build(diagnostic) Then Err.Raise 5, , diagnostic
    If context.RenderTree(diagnostic) Then Err.Raise 5, , "Unbounded star grid accepted"
    context.Dispose
    binding.Dispose
    Exit Sub
EH_GRID:
    Err.Raise Err.Number, "CheckGrid " & phase, Err.Description
End Sub

Private Sub CheckTables(ByVal uiFolder As String)
    Dim binding As New obj_UiBindingContext
    Dim context As New obj_UiRenderContext
    Dim definition As New obj_UiPageDefinition
    Dim first As New obj_UiRawTable
    Dim second As New obj_UiRawTable
    Dim tables As New obj_UiRawTableList
    Dim document As Object
    Dim sheet As Worksheet
    Dim values(1 To 1, 1 To 2) As Variant
    Dim diagnostic As String
    Dim control As obj_IUiControl
    Dim node As Object
    Dim rows As Long, columns As Long

    values(1, 1) = "One"
    values(1, 2) = 10
    first.Initialize values, Array("Name", "Value"), "First"
    values(1, 1) = "Two"
    second.Initialize values, Array("Name", "Value"), "Second"
    tables.Initialize
    tables.Add first
    tables.Add second
    binding.Initialize
    binding.SetObject "Data", "Single", first
    binding.SetObject "Data", "Many", tables
    Set sheet = ThisWorkbook.Worksheets.Add()
    Set document = CreateObject("MSXML2.DOMDocument.6.0")
    If Not document.LoadXML("<page xmlns='urn:excelprototype:profiles' xmlns:c='urn:excelprototype:controls'><grid><c:table name='Single' itemsSource='{Binding Path=Data.Single}' row='2' column='2'/><c:tableList name='Many' itemsSource='{Binding Path=Data.Many}' row='2' column='5' gapRows='2'/></grid></page>") Then Err.Raise 5, , "Table XML"
    definition.Initialize document, "Tables.xaml"
    context.Initialize sheet, definition, uiFolder, binding
    If Not context.Build(diagnostic) Then Err.Raise 5, , diagnostic
    context.Styles.BeginPage sheet, document, uiFolder
    If Not context.RenderTree(diagnostic) Then Err.Raise 5, , diagnostic
    If sheet.Range("B2").Value2 <> "First" Or sheet.Range("B4").Value2 <> "One" Then Err.Raise 5, , "Single table"
    If sheet.Range("E2").Value2 <> "First" Or sheet.Range("E7").Value2 <> "Second" Or sheet.Range("E9").Value2 <> "Two" Then Err.Raise 5, , "Table list placement"
    If sheet.Range("E5").Value2 <> "" Or sheet.Range("E6").Value2 <> "" Then Err.Raise 5, , "Table list gap"
    Set node = document.documentElement.firstChild.firstChild.cloneNode(False)
    node.setAttribute "itemsSource", "{Binding Path=Data.Many}"
    node.setAttribute "type", "Table"
    Set control = ex_UiControlFactory.fn_Create(node)
    If Not control.Configure(node, context, "", diagnostic) Then Err.Raise 5, , diagnostic
    If control.Measure(rows, columns, diagnostic) Then Err.Raise 5, , "Single table accepted multiple tables"
    control.Dispose
    context.Dispose
    binding.Dispose
End Sub

Private Sub CheckSmartTables(ByVal uiFolder As String)
    Dim sheet As Worksheet
    Dim otherSheet As Worksheet
    Dim tables As obj_UiRawTableList
    Dim context As obj_UiRenderContext
    Dim first As ListObject
    Dim second As ListObject
    Dim diagnostic As String
    Dim rules As String
    Dim document As Object
    Dim policy As obj_UiTablePolicy
    Dim representation As String
    Dim tableName As String
    Dim selector As Variant
    Dim intervals As Collection
    Dim errorNumber As Long
    Dim phase As String

    On Error GoTo EH_SMART
    phase = "index selectors"
    Set policy = New obj_UiTablePolicy
    For Each selector In Array("", "0", "-1", "2-1", "1,", "1.5", "1e2", "1--2", "2147483648")
        On Error Resume Next
        Set intervals = policy.ParseIndices(CStr(selector))
        errorNumber = Err.Number
        Err.Clear
        On Error GoTo 0
        If errorNumber = 0 Then Err.Raise 5, , "Invalid index accepted: " & selector
    Next selector
    Set document = CreateObject("MSXML2.DOMDocument.6.0")
    rules = "<c:rule index='1' representation='Smart' excelTableName='tbProbeFirst'/>" & _
        "<c:rule index='2-3, 8-12' representation='Smart'/><c:rule index='3' representation='Raw'/>"
    If Not document.LoadXML("<c:tableList xmlns:c='urn:excelprototype:controls'><c:tableList.tablePolicy><c:tablePolicy defaultRepresentation='Raw'>" & rules & "</c:tablePolicy></c:tableList.tablePolicy></c:tableList>") Then Err.Raise 5, , "Policy XML"
    If Not policy.Initialize(document.documentElement) Then Err.Raise 5, , "Policy initialization"
    policy.Resolve 1, representation, tableName
    If representation <> "Smart" Or tableName <> "tbProbeFirst" Then Err.Raise 5, , "Explicit name rule"
    policy.Resolve 3, representation, tableName
    If representation <> "Raw" Or tableName <> "" Then Err.Raise 5, , "Last rule precedence"
    policy.Resolve 10, representation, tableName
    If representation <> "Smart" Then Err.Raise 5, , "Index interval"
    policy.Dispose

    phase = "initial mixed render"
    Set sheet = ThisWorkbook.Worksheets.Add()
    Set tables = MakeSmartTables(3, 1)
    Set context = RenderSmartFixture(sheet, tables, rules, uiFolder)
    If sheet.ListObjects.Count <> 2 Then Err.Raise 5, , "Mixed table representations"
    Set first = sheet.ListObjects("tbProbeFirst")
    Set second = sheet.ListObjects(2)
    If first.Range.Address <> sheet.Range("A2:B3").Address Then Err.Raise 5, , "Title included in Smart table"
    If Left(second.Name, 12) <> "tbGenerated_" Then Err.Raise 5, , "Automatic name"
    If first.DataBodyRange.Cells(1, 2).Value2 <> 10 Then Err.Raise 5, , "Smart values"
    phase = "repeated render and promotion"
    If Not context.RenderTree(diagnostic) Then Err.Raise 5, , diagnostic
    If sheet.ListObjects.Count <> 2 Or sheet.ListObjects("tbProbeFirst").Name <> "tbProbeFirst" Then _
        Err.Raise 5, , "ListObject was not retained"
    If Not context.TryEnsureSmart("Many", 3, diagnostic, "tbProbePromoted") Then Err.Raise 5, , diagnostic
    If Not context.TryEnsureSmart("Many", 3, diagnostic) Then Err.Raise 5, , diagnostic
    If sheet.ListObjects.Count <> 3 Then Err.Raise 5, , "Promotion is not idempotent"
    If sheet.ListObjects("tbProbePromoted").DataBodyRange.Cells(1, 2).Value2 <> 30 Then Err.Raise 5, , "Promotion changed data"
    If Not context.RenderTree(diagnostic) Then Err.Raise 5, , diagnostic
    If sheet.ListObjects.Count <> 2 Then Err.Raise 5, , "Policy did not restore Raw"

    phase = "resize"
    Set tables = MakeSmartTables(2, 3)
    context.BindingContext.SetObject "Data", "Many", tables
    If Not context.RenderTree(diagnostic) Then Err.Raise 5, , diagnostic
    If sheet.ListObjects.Count <> 2 Or sheet.ListObjects("tbProbeFirst").Name <> "tbProbeFirst" Then _
        Err.Raise 5, , "Resize removed ListObject"
    If sheet.ListObjects("tbProbeFirst").ListRows.Count <> 3 Then Err.Raise 5, , "Resize row count"
    phase = "context recreation"
    context.Dispose
    Set tables = MakeSmartTables(1, 1)
    Set context = RenderSmartFixture(sheet, tables, rules, uiFolder)
    If sheet.ListObjects.Count <> 1 Then Err.Raise 5, , "Stale table survived page rebuild"
    If sheet.ListObjects("tbProbeFirst").Name <> "tbProbeFirst" Then Err.Raise 5, , "Ownership lost on context disposal"
    If sheet.ListObjects("tbProbeFirst").ListRows.Count <> 1 Then Err.Raise 5, , "Shrink row count"
    context.Dispose

    phase = "collision rejection"
    Set otherSheet = ThisWorkbook.Worksheets.Add()
    otherSheet.Range("A1").Value2 = "Key"
    otherSheet.Range("A2").Value2 = 42
    If ex_UiTables.fn_TryEnsureSmart(otherSheet.Range("A1:A2"), "External|1", diagnostic, "tbProbeFirst") Then Err.Raise 5, , "Duplicate name accepted"
    If otherSheet.ListObjects.Count <> 0 Or otherSheet.Range("A2").Value2 <> 42 Then Err.Raise 5, , "Collision mutated target"
    otherSheet.Range("B1").Value2 = "Key"
    If ex_UiTables.fn_TryEnsureSmart(otherSheet.Range("A1:B2"), "External|1", diagnostic) Then Err.Raise 5, , "Duplicate headers accepted"
    phase = "smart to raw"
    Set context = RenderSmartFixture(sheet, tables, "", uiFolder)
    If sheet.ListObjects.Count <> 0 Then Err.Raise 5, , "Smart to Raw cleanup"
    context.Dispose
    Set tables = MakeSmartTables(1, 0)
    Set context = RenderSmartFixture(sheet, tables, "<c:rule index='1' representation='Smart'/>", uiFolder)
    If sheet.ListObjects.Count <> 1 Then Err.Raise 5, , "Empty Smart table"
    If sheet.Range("A2").Value2 <> "Key" Or sheet.Range("A3").Value2 <> "" Then Err.Raise 5, , "Empty Smart layout"
    context.Dispose
    Exit Sub
EH_SMART:
    Err.Raise Err.Number, "CheckSmartTables " & phase, Err.Description
End Sub

Private Function MakeSmartTables( _
    ByVal count As Long, _
    ByVal firstRows As Long _
) As obj_UiRawTableList
    Dim tables As New obj_UiRawTableList
    Dim table As obj_UiRawTable
    Dim values As Variant
    Dim index As Long
    Dim row As Long
    Dim rows As Long

    If Not tables.Initialize() Then Err.Raise 5, , "Table list initialization"
    For index = 1 To count
        rows = 1
        If index = 1 Then rows = firstRows
        Set table = New obj_UiRawTable
        If rows = 0 Then
            If Not table.InitializeEmpty(Array("Key", "Amount"), "Title") Then Err.Raise 5, , "Empty table initialization"
        Else
            ReDim values(1 To rows, 1 To 2)
            For row = 1 To rows
                values(row, 1) = "Item"
                values(row, 2) = index * 10
            Next row
            If Not table.Initialize(values, Array("Key", "Amount"), "Title") Then Err.Raise 5, , "Table initialization"
        End If
        tables.Add table
    Next index
    Set MakeSmartTables = tables
End Function

Private Function RenderSmartFixture( _
    ByVal sheet As Worksheet, _
    ByVal tables As obj_UiRawTableList, _
    ByVal rules As String, _
    ByVal uiFolder As String _
) As obj_UiRenderContext
    Dim binding As New obj_UiBindingContext
    Dim context As New obj_UiRenderContext
    Dim definition As New obj_UiPageDefinition
    Dim document As Object
    Dim diagnostic As String

    Set document = CreateObject("MSXML2.DOMDocument.6.0")
    If Not document.LoadXML("<page xmlns='urn:excelprototype:profiles' xmlns:c='urn:excelprototype:controls'><grid><c:tableList name='Many' itemsSource='{Binding Path=Data.Many}'><c:tableList.tablePolicy><c:tablePolicy defaultRepresentation='Raw'>" & rules & "</c:tablePolicy></c:tableList.tablePolicy></c:tableList></grid></page>") Then Err.Raise 5, , "Smart XML"
    If Not binding.Initialize() Then Err.Raise 5, , "Binding initialization"
    binding.SetObject "Data", "Many", tables
    If Not definition.Initialize(document, "Smart.xaml") Then Err.Raise 5, , "Definition initialization"
    If Not context.Initialize(sheet, definition, uiFolder, binding) Then Err.Raise 5, , "Context initialization"
    If Not context.Build(diagnostic) Then Err.Raise 5, , diagnostic
    context.Styles.BeginPage sheet, document, uiFolder
    If Not context.RenderTree(diagnostic) Then Err.Raise 5, , diagnostic
    Set RenderSmartFixture = context
End Function

Private Sub CheckMarkup(ByVal markup As String, ByVal expected As Boolean, ByVal member As String)
    Dim document As Object
    Dim validator As New obj_UiMarkupValidator
    Dim errors As New Collection
    Dim item As obj_UiMarkupDiagnostic
    Dim found As Boolean

    Set document = CreateObject("MSXML2.DOMDocument.6.0")
    document.preserveWhiteSpace = True
    If Not document.LoadXML(markup) Then Err.Raise 5, , "Invalid test XML"
    If validator.Validate(document.documentElement, errors) <> expected Then Err.Raise 5, , "Unexpected markup validation: " & markup
    If expected Then Exit Sub
    For Each item In errors
        If InStr(1, item.Describe(), member, vbTextCompare) > 0 And Left(item.Path, 5) = "/page" Then found = True
    Next item
    If Not found Then Err.Raise 5, , "Missing markup diagnostic: " & member
End Sub
"@)
    $book.SaveAs($fixturePath, 52)
    try { $result = $excel.Run("'" + $book.Name + "'!ex_UiElementProbe.Run", (Join-Path $projectRoot 'ui/PersonalEventBuilder')) }
    catch {
        $pane = $excel.VBE.ActiveCodePane
        $startLine = 0; $startColumn = 0; $endLine = 0; $endColumn = 0
        $pane.GetSelection([ref]$startLine, [ref]$startColumn, [ref]$endLine, [ref]$endColumn)
        Write-Output ($pane.CodeModule.Name + ':' + $startLine + ' ' + $pane.CodeModule.Lines($startLine, 1))
        throw
    }
    if ($result -ne 'PASS') { throw $result }
    Write-Output 'PASS: real page tree, unchanged XML, layout, reactive binding, events, validation and disposal.'
    Write-Output "Fixture: $fixturePath"
}
finally {
    if ($null -ne $book) { $book.Close($false) }
    if ($null -ne $excel) {
        $excel.Quit()
        [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($excel)
    }
}