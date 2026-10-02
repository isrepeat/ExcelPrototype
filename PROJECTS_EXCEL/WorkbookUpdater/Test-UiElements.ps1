param()

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'
$projectRoot = [IO.Path]::GetFullPath((Join-Path $PSScriptRoot '..'))
$sourceRoot = Join-Path $projectRoot 'vba'
$manifest = Get-Content -LiteralPath (Join-Path $sourceRoot 'modules.json') -Raw | ConvertFrom-Json
foreach ($sourceFile in Get-ChildItem -LiteralPath $sourceRoot -Recurse -Filter '*.vba') {
    $relativePath = [IO.Path]::GetRelativePath($sourceRoot, $sourceFile.FullName).Replace([char]92, [char]47)
    $covered = $false
    foreach ($pattern in $manifest.PersonalEventBuilder) {
        if ($relativePath -like $pattern) { $covered = $true; break }
    }
    if (-not $covered) { throw "Source is missing from modules.json: $relativePath" }
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
    foreach ($file in Get-ChildItem -LiteralPath (Join-Path $projectRoot 'vba') -Recurse -Filter '*.vba') {
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
Public Function Run(ByVal uiFolder As String) As Boolean
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
    Dim resolvedSource As String
    Dim resolvedPath As String

    tagFactory.Initialize "stackpanel"
    ex_UiElementFactory.fn_Register "flow", tagFactory
    controlFactory.Initialize "form"
    ex_UiControlFactory.fn_Register "DraftForm", controlFactory
    CheckMarkup "<page><grid><control type='Label' name='Bad' text='Title' columnSapn='2'/></grid></page>", False, "columnSapn"
    CheckMarkup "<page><grid><stackPanel/></grid></page>", False, "orientation"
    CheckMarkup "<page><grid><page/></grid></page>", False, "page"
    CheckMarkup "<page><grid><control type='Button' name='Bad' caption='Title'><grid/></control></grid></page>", False, "Child element"
    CheckMarkup "<page><grid><control type='Form' name='Bad' source='Form' orientation='vertical'><grid/></control></grid></page>", False, "Child element"
    CheckMarkup "<page><grid><control type='Unknown' name='Bad'/></grid></page>", False, "Unknown Control"
    CheckMarkup "<page><grid><control type='Label' name='Bad' text='Title' row='1.5'/></grid></page>", False, "row"
    CheckMarkup "<page><grid><control type='Input' name='Bad' value='{Binding Path=Name}' readOnly='yes'/></grid></page>", False, "readOnly"
    CheckMarkup "<page><grid/>unexpected text</page>", False, "Text content"
    CheckMarkup "<page><grid/><styles/><styles/></page>", False, "Invalid child count"
    CheckMarkup "<page><grid><styles/></grid></page>", False, "styles"
    CheckMarkup "<page><grid><control type='Label' name='Bad'/></grid></page>", False, "text|caption"
    CheckMarkup "<page><grid><control type='Input' name='Bad' value='{Binding Path=}'/></grid></page>", False, "value"
    CheckMarkup "<page><grid><flow orientation='horizontal'><control type='Label' name='Good' text='Title'/></flow></grid></page>", True, ""
    CheckMarkup "<page/>", False, "Invalid child count"
    CheckMarkup "<page><grid/><grid/></page>", False, "Invalid child count"
    CheckMarkup "<page><grid/><stackPanel orientation='vertical'/></page>", False, "Child element"
    CheckMarkup "<page><grid/><styles><controlStyle/></styles></page>", False, "name"
    CheckMarkup "<page><grid/><styles/></page>", True, ""
    Set sheet = ThisWorkbook.Worksheets("MainPage")
    bindingContext.Initialize
    bindingContext.SetValue "Text", "Title", "Test page"
    bindingContext.SetValue "Text", "HelloWorld", "Hello"
    bindingContext.SetValue "Text", "UpdatePage", "Update"
    bindingContext.SetValue "Text", "GenerateTables", "Generate"
    bindingContext.SetValue "Form", "EventName", "Initial"
    bindingContext.SetValue "Form", "Category", "Meeting"
    bindingContext.SetValue "Form", "Notes", "Notes"
    bindingContext.SetValue "Data", "EventTypes", "Meeting,Training,Leave"
    bindingContext.SetValue "Resources", "PrimaryButton", "primaryButton"
    bindingContext.SetValue "Resources", "PageTitle", "pageTitle"
    tables.Initialize
    bindingContext.SetObject "Data", "Tables", tables
    command.Initialize callbacks, "Execute"
    For Each name In Array("HelloWorldCommand", "UpdatePageCommand", "GenerateTablesCommand", _
        "FormChangedCommand", "SubmitFormCommand")
        bindingContext.SetObject "Commands", CStr(name), command
    Next name
    If Not ex_UiPageLoader.fn_TryLoad(uiFolder & "\MainPage.xaml", definition) Then Err.Raise 5, , "Load"
    snapshot = definition.Document.xml
    If Not context.Initialize(sheet, definition, uiFolder, bindingContext) Then Err.Raise 5, , "Initialize"
    If Not context.Build(diagnostic) Then Err.Raise 5, , diagnostic
    If definition.Document.xml <> snapshot Then Err.Raise 5, , "Page DOM changed during Build"
    context.Styles.BeginPage sheet, definition.Document, uiFolder
    If Not context.RenderTree(diagnostic) Then Err.Raise 5, , diagnostic
    If sheet.Range("D7").Value2 <> "Initial" Then Err.Raise 5, , "Field layout or binding"
    If sheet.Range("D9").MergeArea.Rows.Count <> 3 Then Err.Raise 5, , "Field height"
    bindingContext.SetValue "Text", "Title", "Changed title"
    If sheet.Range("A1").Value2 <> "Changed title" Then Err.Raise 5, , "Label refresh"
    bindingContext.SetValue "Text", "HelloWorld", "Changed button"
    If sheet.Shapes("btn_HelloWorld").TextFrame2.TextRange.Text <> "Changed button" Then Err.Raise 5, , "Button refresh"
    shapeCount = sheet.Shapes.Count
    context.InvalidateMeasure
    If Not context.FlushLayout(diagnostic) Then Err.Raise 5, , diagnostic
    If sheet.Shapes.Count <> shapeCount Then Err.Raise 5, , "Shapes leaked after reflow"
    bindingContext.SetValue "Form", "EventName", "Updated"
    If sheet.Range("D7").Value2 <> "Updated" Then Err.Raise 5, , "Reactive update"
    sheet.Range("D7").Value2 = "User"
    If Not context.Router.DispatchCells(sheet.Range("D7")) Then Err.Raise 5, , "Cell dispatch"
    If callbacks.Count <> 1 Then Err.Raise 5, , "Change command"
    If Not context.Router.DispatchShape("btn_HelloWorld") Then Err.Raise 5, , "Shape dispatch"
    If callbacks.Count <> 2 Then Err.Raise 5, , "Button command"
    Set errors = New Collection
    If Not context.ValidateForm("EventDraftForm", errors) Then Err.Raise 5, , "Valid form"
    bindingContext.SetValue "Form", "EventName", vbNullString
    Set errors = New Collection
    If context.ValidateForm("EventDraftForm", errors) Or errors.Count <> 1 Then Err.Raise 5, , "Required field"
    If definition.Document.xml <> snapshot Then Err.Raise 5, , "Page DOM changed during Render"
    Set otherSheet = ThisWorkbook.Worksheets.Add()
    otherSheet.Name = "OtherPage"
    Set otherDocument = CreateObject("MSXML2.DOMDocument.6.0")
    If Not otherDocument.LoadXML("<page><grid><flow orientation='vertical'><control type='DraftForm' name='OtherForm' source='Draft' orientation='horizontal'>" & _
        "<field name='Accepted' label='Accepted' type='checkbox' required='true'/>" & _
        "<field name='Name' label='Name' type='text' readOnly='true'/></control><stackPanel orientation='horizontal'><control type='Label' name='Tail1' text='Left' columnSpan='2'/><control type='Label' name='Tail2' text='Right' columnSpan='3'/></stackPanel></flow></grid></page>") Then Err.Raise 5, , "Test XML"
    otherDefinition.Initialize otherDocument, "OtherPage.xaml"
    otherBinding.Initialize
    otherBinding.TrySetPathValue "Draft", "Person.Name", "Nested"
    If Not ex_UiBindingRuntime.fn_TryParseBinding("{Binding Path=Person.Name}", _
        "Draft", resolvedSource, resolvedPath, otherBinding) Then Err.Raise 5, , "Nested parse"
    If resolvedSource <> "Draft" Or resolvedPath <> "Person.Name" Then Err.Raise 5, , "Nested source"
    otherBinding.SetValue "Draft", "Accepted", False
    otherBinding.SetValue "Draft", "Name", "Read only"
    otherContext.Initialize otherSheet, otherDefinition, uiFolder, otherBinding
    If Not otherContext.Build(diagnostic) Then Err.Raise 5, , diagnostic
    otherContext.Styles.BeginPage otherSheet, otherDocument, uiFolder
    If Not otherContext.RenderTree(diagnostic) Then Err.Raise 5, , diagnostic
    If otherSheet.Range("A2").Value2 <> "Left" Or otherSheet.Range("C2").Value2 <> "Right" Then Err.Raise 5, , "Nested stack layout"
    If otherSheet.Range("I1").Value2 <> "Read only" Then Err.Raise 5, , "Horizontal layout"
    otherBinding.SetValue "Draft", "Accepted", True
    If otherSheet.Shapes("chk_1").ControlFormat.Value <> xlOn Then Err.Raise 5, , "Checkbox refresh"
    otherSheet.Shapes("chk_1").ControlFormat.Value = xlOff
    If Not otherContext.Router.DispatchShape("chk_1") Then Err.Raise 5, , "Checkbox dispatch"
    If otherSheet.Range("C1").Value2 <> False Then Err.Raise 5, , "Checkbox reverse binding"
    otherSheet.Range("I1").Value2 = "Attempt"
    If Not otherContext.Router.DispatchCells(otherSheet.Range("I1")) Then Err.Raise 5, , "Readonly dispatch"
    If otherSheet.Range("A2").Value2 <> "Left" Or otherSheet.Range("C2").Value2 <> "Right" Then Err.Raise 5, , "Nested stack layout"
    If otherSheet.Range("I1").Value2 <> "Read only" Then Err.Raise 5, , "Readonly restore"
    Set errors = New Collection
    If otherContext.ValidateForm("OtherForm", errors) Or errors.Count <> 1 Then Err.Raise 5, , "Checkbox validation"
    bindingContext.SetValue "Text", "Title", "First page"
    If sheet.Range("A1").Value2 <> "First page" Then Err.Raise 5, , "Page binding isolation"
    If Not context.InvalidateVisual("HelloWorld", diagnostic) Then Err.Raise 5, , diagnostic
    otherContext.Dispose
    otherContext.Dispose
    context.Dispose
    context.Dispose
    bindingContext.SetValue "Form", "EventName", "After disposal"
    If sheet.Range("D7").Value2 <> vbNullString Then Err.Raise 5, , "Subscription survived disposal"
    Run = True
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
    if (-not $excel.Run("'" + $book.Name + "'!ex_UiElementProbe.Run", (Join-Path $projectRoot 'ui/PersonalEventBuilder'))) {
        throw 'UI element contract test failed.'
    }
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