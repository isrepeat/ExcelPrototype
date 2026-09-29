Attribute VB_Name = "ex_UiRuntime"
Option Explicit

Private Const UI_FOLDER_NAME As String = "ui"
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
Public Sub fn_RenderPages()
    Dim targetWorksheet As Worksheet
    Dim uiPageDefinition As obj_UiPageDefinition
    Dim uiRenderContext As obj_UiRenderContext
    Dim xamlPath As String
    Dim uiFolderPath As String
    Dim fileSystem As Object

    ex_Core.fn_Diagnostic_WriteLog "UI_RENDER_STARTED | Workbook=" & ThisWorkbook.Name
    uiFolderPath = ThisWorkbook.Path & "\" & UI_FOLDER_NAME
    Set fileSystem = VBA.CreateObject("Scripting.FileSystemObject")
    ex_UiBindings.fn_Reset

    For Each targetWorksheet In ThisWorkbook.Worksheets
        xamlPath = uiFolderPath & "\" & targetWorksheet.Name & ".xaml"
        If fileSystem.FileExists(xamlPath) Then
            ex_Core.fn_Diagnostic_WriteLog "UI_PAGE_RENDER_STARTED | Sheet=" & _
                targetWorksheet.Name & " | Path=" & xamlPath
            If Not ex_UiPageLoader.fn_TryLoad(xamlPath, uiPageDefinition) Then Exit Sub

            Set uiRenderContext = New obj_UiRenderContext
            If Not uiRenderContext.fn_Initialize( _
                    targetWorksheet, uiPageDefinition, uiFolderPath) Then
                VBA.MsgBox "The UI render context cannot be initialized.", _
                    VBA.vbExclamation, "PersonalEventBuilder"
                Exit Sub
            End If

            ex_StylePipeline.fn_BeginPage targetWorksheet, _
                uiPageDefinition.fn_Document, uiFolderPath
            private_ClearUi targetWorksheet
            ex_StylePipeline.fn_ApplyPagePipeline targetWorksheet
            If Not private_RenderControls(uiRenderContext) Then Exit Sub

            ex_Core.fn_Diagnostic_WriteLog "UI_PAGE_RENDER_COMPLETED | Sheet=" & _
                targetWorksheet.Name
            uiRenderContext.fn_Dispose
            uiPageDefinition.fn_Dispose
            Set uiRenderContext = Nothing
            Set uiPageDefinition = Nothing
        Else
            ex_Core.fn_Diagnostic_WriteLog "UI_PAGE_SKIPPED_NO_XAML | Sheet=" & _
                targetWorksheet.Name & " | Path=" & xamlPath
        End If
    Next targetWorksheet
    ex_Core.fn_Diagnostic_WriteLog "UI_RENDER_COMPLETED | Workbook=" & ThisWorkbook.Name
End Sub
' --------------------------------------
' } // namespace API
' --------------------------------------

Private Function private_RenderControls(ByVal uiRenderContext As obj_UiRenderContext) As Boolean
    Dim controlNode As Object
    Dim uiControl As Object
    Dim targetWorksheet As Worksheet

    Set targetWorksheet = uiRenderContext.fn_TargetWorksheet
    On Error GoTo EH
    For Each controlNode In uiRenderContext.fn_PageDefinition.fn_Document.SelectNodes( _
            "//*[local-name()='control']")
        Set uiControl = ex_UiControlFactory.fn_Create(controlNode)
        If uiControl Is Nothing Then Exit Function
        If Not uiControl.fn_Configure(controlNode) Then Exit Function
        If Not uiControl.fn_Render(uiRenderContext) Then Exit Function
        uiRenderContext.fn_AddControl uiControl
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

    For Each currentShape In targetWorksheet.Shapes
        If VBA.Left$(currentShape.Name, VBA.Len(BUTTON_SHAPE_PREFIX)) = _
           BUTTON_SHAPE_PREFIX Then currentShape.Delete
    Next currentShape
    targetWorksheet.Cells.UnMerge
    targetWorksheet.Cells.Clear
End Sub