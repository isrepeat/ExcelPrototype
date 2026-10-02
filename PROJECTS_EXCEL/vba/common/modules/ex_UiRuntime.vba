Attribute VB_Name = "ex_UiRuntime"
Option Explicit

Private Const BUTTON_SHAPE_PREFIX As String = "btn_"
Private Const SELECT_SHAPE_PREFIX As String = "sel_"

' --------------------------------------
' namespace Lifecycle {
' --------------------------------------
Public Sub fn_Module_Dispose()
End Sub
' --------------------------------------
' } // namespace Lifecycle
' --------------------------------------

' --------------------------------------
' namespace API {
' --------------------------------------
Public Sub fn_RenderPages( _
    ByVal uiFolderRelativePath As String, _
    ByVal uiBindingContext As obj_UiBindingContext _
)
    Dim targetWorksheet As Worksheet
    Dim controlRange As Range
    Dim startedAt As Double

    startedAt = VBA.Timer
    ex_Core.fn_Diagnostic_WriteLog "UI_RENDER_STARTED | Workbook=" & ThisWorkbook.Name
    ex_UiBindings.fn_Reset

    For Each targetWorksheet In ThisWorkbook.Worksheets
        If Not private_RenderPage(targetWorksheet, False, uiFolderRelativePath, uiBindingContext) Then Exit Sub
    Next targetWorksheet
    ex_Core.fn_Diagnostic_WriteLog "UI_RENDER_COMPLETED | Workbook=" & ThisWorkbook.Name
    ex_Core.fn_Diagnostic_WritePerf "RenderPages", startedAt
End Sub

Public Sub fn_RenderActivePage( _
    ByVal uiFolderRelativePath As String, _
    ByVal uiBindingContext As obj_UiBindingContext _
)
    Dim targetWorksheet As Worksheet
    Dim controlRange As Range

    If Not (TypeOf Application.ActiveSheet Is Worksheet) Then Exit Sub
    Set targetWorksheet = Application.ActiveSheet
    If Not (targetWorksheet.Parent Is ThisWorkbook) Then Exit Sub
    private_RenderPage targetWorksheet, True, uiFolderRelativePath, uiBindingContext
End Sub
' --------------------------------------
' } // namespace API
' --------------------------------------

Private Function private_RenderPage( _
    ByVal targetWorksheet As Worksheet, _
    ByVal notifyWhenMissing As Boolean, _
    ByVal uiFolderRelativePath As String, _
    ByVal uiBindingContext As obj_UiBindingContext _
) As Boolean
    Dim uiPageDefinition As obj_UiPageDefinition
    Dim uiRenderContext As obj_UiRenderContext
    Dim xamlPath As String
    Dim uiFolderPath As String
    Dim uiRootPath As String
    Dim fileSystem As Object
    Dim startedAt As Double

    startedAt = VBA.Timer
    If Not ex_RuntimePaths.fn_TryGetUiFolder(uiRootPath) Then Exit Function
    uiFolderPath = uiRootPath & "\" & uiFolderRelativePath
    xamlPath = uiFolderPath & "\" & targetWorksheet.Name & ".xaml"
    Set fileSystem = VBA.CreateObject("Scripting.FileSystemObject")
    If Not fileSystem.FileExists(xamlPath) Then
        ex_Core.fn_Diagnostic_WriteLog "UI_PAGE_SKIPPED_NO_XAML | Sheet=" & _
            targetWorksheet.Name & " | Path=" & xamlPath
        If notifyWhenMissing Then
            VBA.MsgBox "No XAML page was found for: " & targetWorksheet.Name, _
                VBA.vbExclamation, "PersonalEventBuilder"
        End If
        private_RenderPage = True
        Exit Function
    End If

    ex_Core.fn_Diagnostic_WriteLog "UI_PAGE_RENDER_STARTED | Sheet=" & _
        targetWorksheet.Name & " | Path=" & xamlPath
    If Not ex_UiPageLoader.fn_TryLoad(xamlPath, uiPageDefinition) Then Exit Function
    ex_UiBindings.fn_ClearSelectControls targetWorksheet.Name
    ex_UiBindings.fn_ClearCellBindings targetWorksheet.Name
    ex_Core.fn_Diagnostic_WritePerf "Page.LoadXaml | Sheet=" & targetWorksheet.Name, startedAt

    Set uiRenderContext = New obj_UiRenderContext
    If Not uiRenderContext.Initialize( _
            targetWorksheet, uiPageDefinition, uiFolderPath, uiBindingContext) Then
        VBA.MsgBox "The UI render context cannot be initialized.", _
            VBA.vbExclamation, "PersonalEventBuilder"
        Exit Function
    End If

    ex_StylePipeline.fn_BeginPage targetWorksheet, _
        uiPageDefinition.Document, uiFolderPath
    ex_Core.fn_Diagnostic_WritePerf "Page.BeginStyles | Sheet=" & targetWorksheet.Name, startedAt
    private_ClearUi targetWorksheet
    ex_StylePipeline.fn_ApplyPagePipeline targetWorksheet
    ex_Core.fn_Diagnostic_WritePerf "Page.ApplyStyles | Sheet=" & targetWorksheet.Name, startedAt
    private_LogUiScopeVisibility targetWorksheet, "after-pipeline"
    private_RestoreUiScopeVisibility targetWorksheet
    private_LogUiScopeVisibility targetWorksheet, "after-visibility-restore"
    If Not private_RenderControls(uiRenderContext) Then Exit Function
    ex_Core.fn_Diagnostic_WritePerf "Page.RenderControls | Sheet=" & targetWorksheet.Name, startedAt
    private_LogUiScopeVisibility targetWorksheet, "after-controls"

    ex_Core.fn_Diagnostic_WriteLog "UI_PAGE_RENDER_COMPLETED | Sheet=" & _
        targetWorksheet.Name
    uiRenderContext.Dispose
    uiPageDefinition.Dispose
    private_RenderPage = True
    ex_Core.fn_Diagnostic_WritePerf "Page.Render | Sheet=" & targetWorksheet.Name, startedAt
End Function

Private Function private_RenderControls(ByVal uiRenderContext As obj_UiRenderContext) As Boolean
    Dim controlNode As Object
    Dim uiControl As obj_IUiControl
    Dim targetWorksheet As Worksheet
    Dim controlRange As Range
    Dim startedAt As Double
    Dim controlName As String
    Dim controlType As String
    Dim renderStage As String
    Dim errorNumber As Long
    Dim errorDescription As String

    Set targetWorksheet = uiRenderContext.TargetWorksheet
    On Error GoTo EH
    renderStage = "ApplyFormLayouts"
    If Not private_ApplyFormLayouts(uiRenderContext.PageDefinition.Document) Then Exit Function
    For Each controlNode In uiRenderContext.PageDefinition.Document.SelectNodes( _
            "//*[local-name()='control']")
        controlName = private_ReadAttribute(controlNode, "name")
        controlType = private_ReadAttribute(controlNode, "type")
        renderStage = "Create"
        ex_Core.fn_Diagnostic_WriteLog "UI_CONTROL_RENDER_STARTED | Sheet=" & _
            targetWorksheet.Name & " | Name=" & controlName & " | Type=" & controlType
        startedAt = VBA.Timer
        Set uiControl = ex_UiControlFactory.fn_Create(controlNode)
        If uiControl Is Nothing Then
            ex_Core.fn_Diagnostic_WriteLog "UI_CONTROL_RENDER_FAILED | Sheet=" & _
                targetWorksheet.Name & " | Name=" & controlName & " | Type=" & _
                controlType & " | Stage=" & renderStage & " | Reason=FactoryReturnedNothing"
            Exit Function
        End If
        renderStage = "Configure"
        If Not uiControl.Configure(controlNode) Then
            ex_Core.fn_Diagnostic_WriteLog "UI_CONTROL_RENDER_FAILED | Sheet=" & _
                targetWorksheet.Name & " | Name=" & controlName & " | Type=" & _
                controlType & " | Stage=" & renderStage & " | Reason=ConfigureReturnedFalse"
            Exit Function
        End If
        renderStage = "Measure"
        Set controlRange = uiControl.Measure(uiRenderContext)
        If controlRange Is Nothing Then
            ex_Core.fn_Diagnostic_WriteLog "UI_CONTROL_RENDER_FAILED | Sheet=" & _
                targetWorksheet.Name & " | Name=" & controlName & " | Type=" & _
                controlType & " | Stage=" & renderStage & " | Reason=MeasureReturnedNothing"
            Exit Function
        End If
        renderStage = "Render"
        If Not uiControl.Render(uiRenderContext) Then
            ex_Core.fn_Diagnostic_WriteLog "UI_CONTROL_RENDER_FAILED | Sheet=" & _
                targetWorksheet.Name & " | Name=" & controlName & " | Type=" & _
                controlType & " | Stage=" & renderStage & " | Reason=RenderReturnedFalse"
            Exit Function
        End If
        ex_Core.fn_Diagnostic_WritePerf "Control.Render | Type=" & VBA.TypeName(uiControl), startedAt
        uiRenderContext.AddControl uiControl
    Next controlNode
    private_RenderControls = True
    Exit Function

EH:
    errorNumber = VBA.Err.Number
    errorDescription = VBA.Err.Description
    ex_Core.fn_Diagnostic_WriteLog "UI_CONTROL_RENDER_ERROR | Sheet=" & _
        targetWorksheet.Name & " | Name=" & controlName & " | Type=" & controlType & _
        " | Stage=" & renderStage & " | Number=" & VBA.CStr(errorNumber) & _
        " | Description=" & errorDescription
    ' Persist buffered diagnostics before a modal dialog or a possible Excel crash.
    If Not ex_Core.fn_Diagnostic_Flush() Then
        errorDescription = errorDescription & VBA.vbCrLf & "Diagnostic log flush failed."
    End If
    VBA.MsgBox "Cannot render a UI control: " & errorDescription & VBA.vbCrLf & _
        "Control=" & controlName & " | Type=" & controlType & _
        " | Stage=" & renderStage & " | Error=" & VBA.CStr(errorNumber), _
        VBA.vbExclamation, "PersonalEventBuilder"
End Function

Private Function private_ApplyFormLayouts(ByVal pageDocument As Object) As Boolean
    Dim formNode As Object
    Dim parentNode As Object
    Dim hasFormAncestor As Boolean
    Dim usedRows As Long
    Dim usedColumns As Long
    Dim rowStart As Long
    Dim columnStart As Long
    Dim formName As String

    If pageDocument Is Nothing Then Exit Function
    For Each formNode In pageDocument.SelectNodes("//*[local-name()='form']")
        hasFormAncestor = False
        Set parentNode = formNode.parentNode
        Do While Not parentNode Is Nothing
            If VBA.LCase$(VBA.CStr(parentNode.baseName)) = "form" Then
                hasFormAncestor = True
                Exit Do
            End If
            Set parentNode = parentNode.parentNode
        Loop
        If Not hasFormAncestor Then
            formName = VBA.Trim$(private_ReadAttribute(formNode, "name"))
            If VBA.Len(formName) = 0 Then
                VBA.MsgBox "A form container requires a name.", _
                    VBA.vbExclamation, "PersonalEventBuilder"
                Exit Function
            End If
            rowStart = private_ReadLayoutLong(formNode, "row", 1)
            columnStart = private_ReadLayoutLong(formNode, "column", 1)
            If rowStart <= 0 Or columnStart <= 0 Then Exit Function
            If Not private_LayoutContainer(formNode, rowStart, columnStart, usedRows, usedColumns) Then Exit Function
        End If
    Next formNode
    private_ApplyFormLayouts = True
End Function

Private Function private_LayoutContainer( _
    ByVal containerNode As Object, _
    ByVal rowStart As Long, _
    ByVal columnStart As Long, _
    ByRef outRows As Long, _
    ByRef outColumns As Long _
) As Boolean
    Dim childNode As Object
    Dim childKind As String
    Dim orientation As String
    Dim rowCursor As Long
    Dim columnCursor As Long
    Dim childRow As Long
    Dim childColumn As Long
    Dim childRows As Long
    Dim childColumns As Long

    orientation = VBA.LCase$(private_ReadAttribute(containerNode, "orientation"))
    If orientation <> "horizontal" Then orientation = "vertical"
    outRows = 0
    outColumns = 0
    rowCursor = rowStart
    columnCursor = columnStart
    For Each childNode In containerNode.ChildNodes
        If childNode.NodeType <> 1 Then GoTo ContinueChild
        childKind = VBA.LCase$(VBA.CStr(childNode.baseName))
        If childKind <> "control" And childKind <> "stackpanel" And childKind <> "form" Then GoTo ContinueChild
        If orientation = "horizontal" Then
            childRow = rowStart
            childColumn = columnCursor
        Else
            childRow = rowCursor
            childColumn = columnStart
        End If
        If childKind = "control" Then
            childRows = private_ReadLayoutLong(childNode, "rowSpan", 1)
            childColumns = private_ReadLayoutLong(childNode, "columnSpan", 1)
            If childRows <= 0 Or childColumns <= 0 Then Exit Function
            childNode.setAttribute "row", VBA.CStr(childRow)
            childNode.setAttribute "column", VBA.CStr(childColumn)
        Else
            If Not private_LayoutContainer(childNode, childRow, childColumn, childRows, childColumns) Then Exit Function
        End If
        If orientation = "horizontal" Then
            columnCursor = columnCursor + childColumns
            If childRows > outRows Then outRows = childRows
        Else
            rowCursor = rowCursor + childRows
            If childColumns > outColumns Then outColumns = childColumns
        End If
ContinueChild:
    Next childNode
    If orientation = "horizontal" Then
        outColumns = columnCursor - columnStart
        If outRows = 0 Then outRows = 1
    Else
        outRows = rowCursor - rowStart
        If outColumns = 0 Then outColumns = 1
    End If
    private_LayoutContainer = True
End Function

Private Function private_ReadLayoutLong( _
    ByVal node As Object, _
    ByVal attributeName As String, _
    ByVal defaultValue As Long _
) As Long
    Dim valueText As String

    valueText = private_ReadAttribute(node, attributeName)
    If VBA.IsNumeric(valueText) Then
        private_ReadLayoutLong = VBA.CLng(valueText)
    Else
        private_ReadLayoutLong = defaultValue
    End If
End Function

Private Function private_ReadAttribute(ByVal node As Object, ByVal attributeName As String) As String
    Dim value As Variant

    If node Is Nothing Then Exit Function
    value = node.getAttribute(attributeName)
    If VBA.IsNull(value) Or VBA.IsEmpty(value) Then Exit Function
    private_ReadAttribute = VBA.CStr(value)
End Function

Private Sub private_ClearUi(ByVal targetWorksheet As Worksheet)
    Dim currentShape As Shape
    Dim uiScope As Range
    Dim shapeIndex As Long

    For shapeIndex = targetWorksheet.Shapes.Count To 1 Step -1
        Set currentShape = targetWorksheet.Shapes(shapeIndex)
        If VBA.Left$(currentShape.Name, VBA.Len(BUTTON_SHAPE_PREFIX)) = _
           BUTTON_SHAPE_PREFIX Or VBA.Left$(currentShape.Name, VBA.Len(SELECT_SHAPE_PREFIX)) = _
           SELECT_SHAPE_PREFIX Then _
            currentShape.Delete
    Next shapeIndex
    Set uiScope = targetWorksheet.Range("A1:AN100")
    private_LogUiScopeVisibility targetWorksheet, "before-clear"
    private_RestoreUiScopeVisibility targetWorksheet
    uiScope.UnMerge
    uiScope.Clear
    private_LogUiScopeVisibility targetWorksheet, "after-clear"
    private_RestoreUiScopeVisibility targetWorksheet
    private_LogUiScopeVisibility targetWorksheet, "after-clear-visibility-restore"
End Sub

Private Sub private_RestoreUiScopeVisibility(ByVal targetWorksheet As Worksheet)
    Dim uiScope As Range

    Set uiScope = targetWorksheet.Range("A1:AN100")
    uiScope.EntireRow.Hidden = False
    uiScope.EntireColumn.Hidden = False
End Sub

Private Sub private_LogUiScopeVisibility( _
    ByVal targetWorksheet As Worksheet, _
    ByVal stageName As String _
)
    Dim uiScope As Range

    Set uiScope = targetWorksheet.Range("A1:AN100")
    ex_Core.fn_Diagnostic_WriteLog "UI_PAGE_VISIBILITY | Sheet=" & _
        targetWorksheet.Name & " | Stage=" & stageName & _
        " | HiddenRows=" & VBA.CStr(private_CountHiddenRows(uiScope)) & _
        " | ZeroHeightRows=" & VBA.CStr(private_CountZeroHeightRows(uiScope))
End Sub

Private Function private_CountHiddenRows(ByVal targetRange As Range) As Long
    Dim currentRow As Range

    For Each currentRow In targetRange.Rows
        If currentRow.EntireRow.Hidden Then _
            private_CountHiddenRows = private_CountHiddenRows + 1
    Next currentRow
End Function

Private Function private_CountZeroHeightRows(ByVal targetRange As Range) As Long
    Dim currentRow As Range

    For Each currentRow In targetRange.Rows
        If currentRow.EntireRow.RowHeight = 0 Then _
            private_CountZeroHeightRows = private_CountZeroHeightRows + 1
    Next currentRow
End Function