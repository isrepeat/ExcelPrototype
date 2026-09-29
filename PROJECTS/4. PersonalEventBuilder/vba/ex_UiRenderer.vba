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

    ex_UiBindings.fn_Reset
    For Each worksheet In ThisWorkbook.Worksheets
        xamlPath = ThisWorkbook.Path & "\" & UI_FOLDER_NAME & "\" & _
            worksheet.Name & ".xaml"
        If VBA.Len(VBA.Dir$(xamlPath)) > 0 Then
            If Not ex_UiParser.fn_TryLoadPage(xamlPath, page) Then Exit Sub
            private_ClearUi worksheet
            private_RenderControls worksheet, page
        End If
    Next worksheet
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

    For Each controlNode In page.SelectNodes("//*[local-name()='control']")
        controlType = VBA.LCase$(VBA.Trim$(controlNode.getAttribute("type")))
        rowNo = private_ReadLong(controlNode, "row", 1)
        columnNo = private_ReadLong(controlNode, "column", 1)
        rowSpan = private_ReadLong(controlNode, "rowSpan", 1)
        columnSpan = private_ReadLong(controlNode, "columnSpan", 1)
        Set targetRange = worksheet.Range(worksheet.Cells(rowNo, columnNo), _
            worksheet.Cells(rowNo + rowSpan - 1, columnNo + columnSpan - 1))
        caption = VBA.CStr(controlNode.getAttribute("text"))
        If VBA.Len(caption) = 0 Then caption = VBA.CStr(controlNode.getAttribute("caption"))
        Select Case controlType
            Case "label"
                targetRange.Merge
                targetRange.Value2 = caption
                private_ApplyStyle targetRange, controlNode
            Case "button"
                Set buttonShape = worksheet.Shapes.AddShape( _
                    msoShapeRoundedRectangle, targetRange.Left, targetRange.Top, _
                    targetRange.Width, targetRange.Height)
                buttonShape.Name = "btn_" & VBA.CStr(controlNode.getAttribute("name"))
                buttonShape.TextFrame2.TextRange.Text = caption
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
End Sub

Private Function private_ReadCallback(ByVal controlNode As Object) As String
    Dim bindingText As String
    Dim moduleName As String
    Dim methodName As String

    bindingText = VBA.CStr(controlNode.getAttribute("onClick"))
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
    If VBA.IsNumeric(node.getAttribute(attributeName)) Then
        private_ReadLong = VBA.CLng(node.getAttribute(attributeName))
    Else
        private_ReadLong = defaultValue
    End If
End Function

Private Sub private_ApplyStyle(ByVal targetRange As Range, ByVal controlNode As Object)
    targetRange.Font.Name = "Calibri"
    targetRange.Font.Size = private_ReadLong(controlNode, "fontSize", 12)
    targetRange.HorizontalAlignment = xlCenter
    targetRange.VerticalAlignment = xlCenter
End Sub

Private Sub private_ClearUi(ByVal worksheet As Worksheet)
    Dim currentShape As Shape

    For Each currentShape In worksheet.Shapes
        If VBA.Left$(currentShape.Name, 4) = "btn_" Then currentShape.Delete
    Next currentShape
    worksheet.Cells.Clear
End Sub