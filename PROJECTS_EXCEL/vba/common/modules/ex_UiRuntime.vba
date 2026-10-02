Attribute VB_Name = "ex_UiRuntime"
Option Explicit

Private m_pages As Object
Private Const BUTTON_SHAPE_PREFIX As String = "btn_"
Private Const SELECT_SHAPE_PREFIX As String = "sel_"

' --------------------------------------
' namespace Lifecycle {
' --------------------------------------
Public Sub fn_Module_Dispose()
    Dim key As Variant
    Dim context As obj_UiRenderContext

    If Not m_pages Is Nothing Then
        For Each key In m_pages.Keys
            Set context = m_pages(key)
            context.Dispose
        Next key
    End If
    Set m_pages = Nothing
End Sub
' --------------------------------------
' } // namespace Lifecycle
' --------------------------------------

' --------------------------------------
' namespace API {
' --------------------------------------
Public Function fn_TryGetContext( _
    ByVal worksheet As Worksheet, _
    ByRef context As obj_UiRenderContext _
) As Boolean
    Set context = Nothing
    If m_pages Is Nothing Then Exit Function
    If Not worksheet.Parent Is ThisWorkbook Then Exit Function
    If Not m_pages.Exists(worksheet.Name) Then Exit Function
    Set context = m_pages(worksheet.Name)
    fn_TryGetContext = True
End Function

Public Function fn_RenderPages( _
    ByVal uiFolderRelativePath As String, _
    ByVal uiBindingContext As obj_UiBindingContext _
) As Boolean
    Dim targetWorksheet As Worksheet
    Dim controlRange As Range
    Dim startedAt As Double

    startedAt = VBA.Timer
    ex_Core.fn_Diagnostic_WriteLog "UI_RENDER_STARTED | Workbook=" & ThisWorkbook.Name
    fn_Module_Dispose

    For Each targetWorksheet In ThisWorkbook.Worksheets
        If Not private_RenderPage(targetWorksheet, False, uiFolderRelativePath, uiBindingContext) Then Exit Function
    Next targetWorksheet
    fn_RenderPages = True
    ex_Core.fn_Diagnostic_WriteLog "UI_RENDER_COMPLETED | Workbook=" & ThisWorkbook.Name
    ex_Core.fn_Diagnostic_WritePerf "RenderPages", startedAt
End Function

Public Function fn_RenderActivePage( _
    ByVal uiFolderRelativePath As String, _
    ByVal uiBindingContext As obj_UiBindingContext _
) As Boolean
    Dim targetWorksheet As Worksheet
    Dim controlRange As Range

    If Not (TypeOf Application.ActiveSheet Is Worksheet) Then Exit Function
    Set targetWorksheet = Application.ActiveSheet
    If Not (targetWorksheet.Parent Is ThisWorkbook) Then Exit Function
    fn_RenderActivePage = private_RenderPage(targetWorksheet, True, uiFolderRelativePath, uiBindingContext)
End Function
' --------------------------------------
' } // namespace API
' --------------------------------------

' --------------------------------------
' namespace Private {
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
            ex_WindowsUi.fn_ShowMessage "No XAML page was found for: " & targetWorksheet.Name, _
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
        ex_WindowsUi.fn_ShowMessage "The UI render context cannot be initialized.", _
            VBA.vbExclamation, "PersonalEventBuilder"
        Exit Function
    End If

    If m_pages Is Nothing Then
        Set m_pages = VBA.CreateObject("Scripting.Dictionary")
        m_pages.CompareMode = VBA.vbTextCompare
    End If
    Dim diagnostic As String
    Dim previousContext As obj_UiRenderContext

    If Not uiRenderContext.Build(diagnostic) Then
        uiRenderContext.Dispose
        ex_WindowsUi.fn_ShowMessage diagnostic, VBA.vbExclamation, "UI configuration"
        Exit Function
    End If
    If m_pages.Exists(targetWorksheet.Name) Then
        Set previousContext = m_pages(targetWorksheet.Name)
        previousContext.Dispose
        m_pages.Remove targetWorksheet.Name
    End If
    Set m_pages(targetWorksheet.Name) = uiRenderContext
    uiRenderContext.Styles.BeginPage targetWorksheet, _
        uiPageDefinition.Document, uiFolderPath
    ex_Core.fn_Diagnostic_WritePerf "Page.BeginStyles | Sheet=" & targetWorksheet.Name, startedAt
    private_ClearUi targetWorksheet
    uiRenderContext.Styles.ApplyPagePipeline targetWorksheet
    ex_Core.fn_Diagnostic_WritePerf "Page.ApplyStyles | Sheet=" & targetWorksheet.Name, startedAt
    private_LogUiScopeVisibility targetWorksheet, "after-pipeline"
    private_RestoreUiScopeVisibility targetWorksheet
    private_LogUiScopeVisibility targetWorksheet, "after-visibility-restore"
    If Not uiRenderContext.RenderTree(diagnostic) Then
        uiRenderContext.Dispose
        m_pages.Remove targetWorksheet.Name
        ex_WindowsUi.fn_ShowMessage diagnostic, VBA.vbExclamation, "UI render"
        Exit Function
    End If
    ex_Core.fn_Diagnostic_WritePerf "Page.RenderControls | Sheet=" & targetWorksheet.Name, startedAt
    private_LogUiScopeVisibility targetWorksheet, "after-controls"

    ex_Core.fn_Diagnostic_WriteLog "UI_PAGE_RENDER_COMPLETED | Sheet=" & _
        targetWorksheet.Name

    private_RenderPage = True
    ex_Core.fn_Diagnostic_WritePerf "Page.Render | Sheet=" & targetWorksheet.Name, startedAt
End Function

Private Sub private_ClearUi(ByVal targetWorksheet As Worksheet)
    Dim currentShape As Shape
    Dim uiScope As Range
    Dim shapeIndex As Long

    For shapeIndex = targetWorksheet.Shapes.Count To 1 Step -1
        Set currentShape = targetWorksheet.Shapes(shapeIndex)
        If VBA.Left$(currentShape.Name, VBA.Len(BUTTON_SHAPE_PREFIX)) = _
           BUTTON_SHAPE_PREFIX Or VBA.Left$(currentShape.Name, VBA.Len(SELECT_SHAPE_PREFIX)) = _
           SELECT_SHAPE_PREFIX Or VBA.Left$(currentShape.Name, 4) = "chk_" Then _
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
' --------------------------------------
' } // namespace Private
' --------------------------------------