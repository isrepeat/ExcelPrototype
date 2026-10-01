Attribute VB_Name = "ex_UiRuntime"
Option Explicit

Private Const BUTTON_SHAPE_PREFIX As String = "btn_"

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
Public Sub fn_RenderPages(ByVal uiFolderRelativePath As String, ByVal uiBindingContext As obj_UiBindingContext)
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

Public Sub fn_RenderActivePage(ByVal uiFolderRelativePath As String, ByVal uiBindingContext As obj_UiBindingContext)
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

    Set targetWorksheet = uiRenderContext.TargetWorksheet
    On Error GoTo EH
    For Each controlNode In uiRenderContext.PageDefinition.Document.SelectNodes( _
            "//*[local-name()='control']")
        startedAt = VBA.Timer
        Set uiControl = ex_UiControlFactory.fn_Create(controlNode)
        If uiControl Is Nothing Then Exit Function
        If Not uiControl.Configure(controlNode) Then Exit Function
        Set controlRange = uiControl.Measure(uiRenderContext)
        If controlRange Is Nothing Then Exit Function
        If Not uiControl.Render(uiRenderContext) Then Exit Function
        ex_Core.fn_Diagnostic_WritePerf "Control.Render | Type=" & VBA.TypeName(uiControl), startedAt
        uiRenderContext.AddControl uiControl
    Next controlNode
    private_RenderControls = True
    Exit Function

EH:
    ex_Core.fn_Diagnostic_WriteLog "UI_CONTROL_RENDER_ERROR | Sheet=" & _
        targetWorksheet.Name & " | Number=" & VBA.CStr(VBA.Err.Number) & _
        " | Description=" & VBA.Err.Description
    VBA.MsgBox "Cannot render a UI control: " & VBA.Err.Description, _
        VBA.vbExclamation, "PersonalEventBuilder"
End Function

Private Sub private_ClearUi(ByVal targetWorksheet As Worksheet)
    Dim currentShape As Shape
    Dim uiScope As Range

    For Each currentShape In targetWorksheet.Shapes
        If VBA.Left$(currentShape.Name, VBA.Len(BUTTON_SHAPE_PREFIX)) = _
           BUTTON_SHAPE_PREFIX Then currentShape.Delete
    Next currentShape
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