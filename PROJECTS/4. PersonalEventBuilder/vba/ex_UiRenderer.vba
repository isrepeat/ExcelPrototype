Option Explicit

Private Const UI_FOLDER_NAME As String = "ui"

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
    Dim page As Object
    Dim worksheet As Worksheet
    Dim xamlPath As String
    Dim fileSystem As Object

    ex_Core.fn_Diagnostic_WriteLog "UI_RENDER_STARTED | Workbook=" & ThisWorkbook.Name
    Set fileSystem = VBA.CreateObject("Scripting.FileSystemObject")
    ex_UiBindings.fn_Reset
    For Each worksheet In ThisWorkbook.Worksheets
        xamlPath = ThisWorkbook.Path & "\" & UI_FOLDER_NAME & "\" & _
            worksheet.Name & ".xaml"
        If fileSystem.FileExists(xamlPath) Then
            ex_Core.fn_Diagnostic_WriteLog "UI_PAGE_RENDER_STARTED | Sheet=" & _
                worksheet.Name & " | Path=" & xamlPath
            If Not ex_UiParser.fn_TryLoadPage(xamlPath, page) Then Exit Sub
            ex_StylePipeline.fn_BeginPage worksheet, page, ThisWorkbook.Path & "\" & UI_FOLDER_NAME
            private_ClearUi worksheet
            ex_StylePipeline.fn_ApplyPagePipeline worksheet
            private_RenderControls worksheet, page
            ex_Core.fn_Diagnostic_WriteLog "UI_PAGE_RENDER_COMPLETED | Sheet=" & _
                worksheet.Name
        Else
            ex_Core.fn_Diagnostic_WriteLog "UI_PAGE_SKIPPED_NO_XAML | Sheet=" & _
                worksheet.Name & " | Path=" & xamlPath
        End If
    Next worksheet
    ex_Core.fn_Diagnostic_WriteLog "UI_RENDER_COMPLETED | Workbook=" & ThisWorkbook.Name
End Sub
' --------------------------------------
' } // namespace API
' --------------------------------------

Private Sub private_RenderControls(ByVal worksheet As Worksheet, ByVal page As Object)
    Dim controlNode As Object
    Dim controlType As String
    Dim rowNo As Long
    Dim columnNo As Long
    Dim rowSpan As Long
    Dim columnSpan As Long
    Dim targetRange As Range
    Dim buttonShape As Shape
    Dim caption As String
    Dim callbackName As String
    Dim controlName As String

    On Error GoTo EH

    For Each controlNode In page.SelectNodes("//*[local-name()='control']")
        controlType = VBA.LCase$(VBA.Trim$(private_ReadText(controlNode, "type")))
        controlName = private_ReadText(controlNode, "name")
        rowNo = private_ReadLong(controlNode, "row", 1)
        columnNo = private_ReadLong(controlNode, "column", 1)
        rowSpan = private_ReadLong(controlNode, "rowSpan", 1)
        columnSpan = private_ReadLong(controlNode, "columnSpan", 1)
        Set targetRange = worksheet.Range(worksheet.Cells(rowNo, columnNo), _
            worksheet.Cells(rowNo + rowSpan - 1, columnNo + columnSpan - 1))
        caption = private_ReadText(controlNode, "text")
        If VBA.Len(caption) = 0 Then caption = private_ReadText(controlNode, "caption")
        Select Case controlType
            Case "label"
                targetRange.Merge
                targetRange.Value2 = caption
                ex_StylePipeline.fn_ApplyControlStyle targetRange, Nothing, controlNode
            Case "button"
                Set buttonShape = worksheet.Shapes.AddShape( _
                    msoShapeRoundedRectangle, targetRange.Left, targetRange.Top, _
                    targetRange.Width, targetRange.Height)
                buttonShape.Name = "btn_" & controlName
                buttonShape.TextFrame2.TextRange.Text = caption
                ex_StylePipeline.fn_ApplyControlStyle targetRange, buttonShape, controlNode
                buttonShape.OnAction = "ex_UiBridge.fn_OnShapeClick"
                callbackName = private_ReadCallback(controlNode)
                If VBA.Len(callbackName) = 0 Then
                    VBA.MsgBox "Button binding is required: " & buttonShape.Name, _
                        VBA.vbExclamation, "PersonalEventBuilder"
                Else
                    ex_UiBindings.fn_Register buttonShape.Name, callbackName
                End If
        End Select
    Next controlNode
    Exit Sub

EH:
    ex_Core.fn_Diagnostic_WriteLog "UI_CONTROL_RENDER_ERROR | Sheet=" & worksheet.Name & _
        " | Control=" & controlName & _
        " | Type=" & controlType & _
        " | Number=" & VBA.CStr(VBA.Err.Number) & _
        " | Description=" & VBA.Err.Description
    VBA.Err.Raise VBA.Err.Number, "private_RenderControls", VBA.Err.Description
End Sub

Private Function private_ReadCallback(ByVal controlNode As Object) As String
    Dim bindingText As String
    Dim moduleName As String
    Dim methodName As String

    bindingText = private_ReadText(controlNode, "onClick")
    moduleName = private_ReadBindingValue(bindingText, "Module=")
    methodName = private_ReadBindingValue(bindingText, "Method=")
    If VBA.Len(moduleName) > 0 And VBA.Len(methodName) > 0 Then _
        private_ReadCallback = moduleName & "." & methodName
End Function

Private Function private_ReadBindingValue(ByVal bindingText As String, ByVal keyName As String) As String
    Dim valueStart As Long
    Dim valueEnd As Long

    valueStart = VBA.InStr(1, bindingText, keyName, VBA.vbTextCompare)
    If valueStart = 0 Then Exit Function
    valueStart = valueStart + VBA.Len(keyName)
    valueEnd = VBA.InStr(valueStart, bindingText, ";")
    If valueEnd = 0 Then valueEnd = VBA.InStr(valueStart, bindingText, "}")
    private_ReadBindingValue = VBA.Trim$(VBA.Mid$(bindingText, valueStart, valueEnd - valueStart))
End Function

Private Function private_ReadLong(ByVal node As Object, ByVal attributeName As String, ByVal defaultValue As Long) As Long
    Dim valueText As String

    valueText = private_ReadText(node, attributeName)
    If VBA.IsNumeric(valueText) Then
        private_ReadLong = VBA.CLng(valueText)
    Else
        private_ReadLong = defaultValue
    End If
End Function

Private Function private_ReadText(ByVal node As Object, ByVal attributeName As String) As String
    Dim attributeValue As Variant

    attributeValue = node.getAttribute(attributeName)
    If VBA.IsNull(attributeValue) Or VBA.IsEmpty(attributeValue) Then Exit Function
    private_ReadText = VBA.CStr(attributeValue)
End Function

Private Sub private_ClearUi(ByVal worksheet As Worksheet)
    Dim currentShape As Shape

    For Each currentShape In worksheet.Shapes
        If VBA.Left$(currentShape.Name, 4) = "btn_" Then currentShape.Delete
    Next currentShape
    worksheet.Cells.UnMerge
    worksheet.Cells.Clear
End Sub