VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "obj_LabelControlVM"
Option Explicit
Implements obj_IControl

Private m_ControlName As String
Private m_TextRaw As String
Private m_TextResolved As String
Private m_StyleName As String
Private m_LayoutSheet As String
Private m_RowStart As Long
Private m_ColStart As Long
Private m_RowEnd As Long
Private m_ColEnd As Long
Private m_IsConfigured As Boolean

Private Sub obj_IControl_Configure(ByVal controlNode As Object)
    m_IsConfigured = False

    If controlNode Is Nothing Then
        MsgBox "Label: control node is not specified.", vbExclamation
        Exit Sub
    End If

    m_ControlName = Trim$(ex_XmlCore.m_NodeAttrText(controlNode, "name"))
    If Len(m_ControlName) = 0 Then m_ControlName = "label"

    m_TextRaw = CStr(ex_XmlCore.m_NodeAttrText(controlNode, "text"))
    If Len(Trim$(m_TextRaw)) = 0 Then
        m_TextRaw = CStr(ex_XmlCore.m_NodeAttrText(controlNode, "caption"))
    End If
    m_StyleName = LCase$(Trim$(ex_XmlCore.m_NodeAttrText(controlNode, "style")))

    If Not ex_BindingRuntime.m_TryResolveTextBinding(m_TextRaw, Me, m_TextResolved) Then Exit Sub

    m_LayoutSheet = Trim$(ex_XmlCore.m_NodeAttrText(controlNode, "__layoutSheet"))
    If Len(m_LayoutSheet) = 0 Then
        MsgBox "Label: runtime layout sheet is missing for control '" & m_ControlName & "'.", vbExclamation
        Exit Sub
    End If

    If Not mp_TryReadLayoutLongAttr(controlNode, "__layoutRowStart", m_RowStart) Then Exit Sub
    If Not mp_TryReadLayoutLongAttr(controlNode, "__layoutColStart", m_ColStart) Then Exit Sub
    If Not mp_TryReadLayoutLongAttr(controlNode, "__layoutRowEnd", m_RowEnd) Then Exit Sub
    If Not mp_TryReadLayoutLongAttr(controlNode, "__layoutColEnd", m_ColEnd) Then Exit Sub

    If m_RowStart <= 0 Or m_ColStart <= 0 Then
        MsgBox "Label: invalid row/column start for control '" & m_ControlName & "'.", vbExclamation
        Exit Sub
    End If

    If m_RowEnd < m_RowStart Then
        MsgBox "Label: control '" & m_ControlName & "' has invalid spanRows range.", vbExclamation
        Exit Sub
    End If

    If m_ColEnd < m_ColStart Then
        MsgBox "Label: control '" & m_ControlName & "' has invalid spanCells range.", vbExclamation
        Exit Sub
    End If

    m_IsConfigured = True
End Sub

Private Sub obj_IControl_Render(ByVal wb As Workbook)
    Dim ws As Worksheet
    Dim targetRange As Range

    If Not m_IsConfigured Then
        MsgBox "Label: control '" & m_ControlName & "' is not configured.", vbExclamation
        Exit Sub
    End If

    If wb Is Nothing Then Set wb = ThisWorkbook
    If wb Is Nothing Then
        MsgBox "Label: workbook is not specified for control '" & m_ControlName & "'.", vbExclamation
        Exit Sub
    End If

    Set ws = mp_GetWorksheetByName(wb, m_LayoutSheet)
    If ws Is Nothing Then
        MsgBox "Label: sheet '" & m_LayoutSheet & "' was not found for control '" & m_ControlName & "'.", vbExclamation
        Exit Sub
    End If

    On Error GoTo EH_RANGE
    Set targetRange = ws.Range(ws.Cells(m_RowStart, m_ColStart), ws.Cells(m_RowEnd, m_ColEnd))
    On Error GoTo 0

    targetRange.Value2 = m_TextResolved
    targetRange.HorizontalAlignment = xlHAlignLeft
    targetRange.VerticalAlignment = xlVAlignCenter
    targetRange.WrapText = False
    If Not mp_ApplyPresetStyle(targetRange, m_StyleName) Then Exit Sub
    Exit Sub

EH_RANGE:
    MsgBox "Label: failed to resolve target range for control '" & m_ControlName & "'.", vbExclamation
End Sub

Private Function obj_IControl_SupportsAttribute(ByVal attrName As String) As Boolean
    Select Case LCase$(Trim$(attrName))
        Case "text", "caption"
            obj_IControl_SupportsAttribute = True
    End Select
End Function

Private Function mp_ApplyPresetStyle(ByVal targetRange As Range, ByVal styleName As String) As Boolean
    If targetRange Is Nothing Then Exit Function

    Select Case LCase$(Trim$(styleName))
        Case vbNullString
            ' no-op

        Case "tablesection"
            targetRange.Interior.Color = RGB(23, 58, 94)
            targetRange.Font.Color = RGB(234, 246, 255)
            targetRange.Font.Bold = True
            targetRange.Borders.LineStyle = xlContinuous
            targetRange.Borders.Color = RGB(14, 34, 57)
            targetRange.Borders.Weight = xlThin

        Case "tableheadercell"
            targetRange.Interior.Color = RGB(43, 74, 107)
            targetRange.Font.Color = RGB(221, 238, 255)
            targetRange.Font.Bold = True
            targetRange.Borders.LineStyle = xlContinuous
            targetRange.Borders.Color = RGB(31, 54, 80)
            targetRange.Borders.Weight = xlThin

        Case "tabledatacell"
            targetRange.Interior.Color = RGB(58, 58, 58)
            targetRange.Font.Color = RGB(240, 240, 240)
            targetRange.Font.Bold = False
            targetRange.Borders.LineStyle = xlContinuous
            targetRange.Borders.Color = RGB(42, 42, 42)
            targetRange.Borders.Weight = xlThin

        Case "tablespacer"
            targetRange.Interior.Color = RGB(31, 31, 31)
            targetRange.Font.Color = RGB(31, 31, 31)
            targetRange.Font.Bold = False
            targetRange.Borders.LineStyle = xlContinuous
            targetRange.Borders.Color = RGB(31, 31, 31)
            targetRange.Borders.Weight = xlHairline

        Case Else
            MsgBox "Label: unsupported style '" & styleName & "' for control '" & m_ControlName & "'.", vbExclamation
            Exit Function
    End Select

    mp_ApplyPresetStyle = True
End Function

Private Function mp_TryReadLayoutLongAttr( _
    ByVal controlNode As Object, _
    ByVal attrName As String, _
    ByRef outValue As Long _
) As Boolean
    Dim rawText As String

    rawText = Trim$(ex_XmlCore.m_NodeAttrText(controlNode, attrName))
    If Len(rawText) = 0 Then
        MsgBox "Label: runtime layout attribute '" & attrName & "' is missing for control '" & m_ControlName & "'.", vbExclamation
        Exit Function
    End If

    If Not IsNumeric(rawText) Then
        MsgBox "Label: runtime layout attribute '" & attrName & "' must be numeric for control '" & m_ControlName & "'.", vbExclamation
        Exit Function
    End If

    outValue = CLng(rawText)
    mp_TryReadLayoutLongAttr = True
End Function

Private Function mp_GetWorksheetByName(ByVal wb As Workbook, ByVal sheetName As String) As Worksheet
    On Error Resume Next
    Set mp_GetWorksheetByName = wb.Worksheets(sheetName)
    On Error GoTo 0
End Function
