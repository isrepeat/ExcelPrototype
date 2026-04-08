VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "obj_ButtonControlVM"
Option Explicit
Implements obj_IControl

Private Const DEFAULT_CAPTION As String = "Update Code"

Private m_ControlName As String
Private m_CaptionRaw As String
Private m_OnClickRaw As String
Private m_StyleName As String
Private m_CaptionText As String
Private m_OnClickMacroRef As String
Private m_LayoutSheet As String
Private m_RowStart As Long
Private m_ColStart As Long
Private m_RowEnd As Long
Private m_ColEnd As Long
Private m_IsConfigured As Boolean

Private Sub obj_IControl_Configure(ByVal controlNode As Object)
    m_IsConfigured = False

    If controlNode Is Nothing Then
        MsgBox "Button: control node is not specified.", vbExclamation
        Exit Sub
    End If

    m_ControlName = Trim$(ex_XmlCore.m_NodeAttrText(controlNode, "name"))
    If Len(m_ControlName) = 0 Then m_ControlName = "button"

    m_CaptionRaw = CStr(ex_XmlCore.m_NodeAttrText(controlNode, "caption"))
    If Len(m_CaptionRaw) = 0 Then m_CaptionRaw = DEFAULT_CAPTION

    m_StyleName = Trim$(ex_XmlCore.m_NodeAttrText(controlNode, "style"))

    m_OnClickRaw = CStr(ex_XmlCore.m_NodeAttrText(controlNode, "onClick"))
    If Len(Trim$(m_OnClickRaw)) = 0 Then
        MsgBox "Button: onClick is required for control '" & m_ControlName & "'.", vbExclamation
        Exit Sub
    End If

    If Not ex_BindingRuntime.m_TryResolveTextBinding(m_CaptionRaw, Me, m_CaptionText) Then Exit Sub
    If Not ex_BindingRuntime.m_TryResolveMacroBinding(m_OnClickRaw, Me, m_OnClickMacroRef) Then Exit Sub

    m_LayoutSheet = Trim$(ex_XmlCore.m_NodeAttrText(controlNode, "__layoutSheet"))
    If Len(m_LayoutSheet) = 0 Then
        MsgBox "Button: runtime layout sheet is missing for control '" & m_ControlName & "'.", vbExclamation
        Exit Sub
    End If

    If Not mp_TryReadLayoutLongAttr(controlNode, "__layoutRowStart", m_RowStart) Then Exit Sub
    If Not mp_TryReadLayoutLongAttr(controlNode, "__layoutColStart", m_ColStart) Then Exit Sub
    If Not mp_TryReadLayoutLongAttr(controlNode, "__layoutRowEnd", m_RowEnd) Then Exit Sub
    If Not mp_TryReadLayoutLongAttr(controlNode, "__layoutColEnd", m_ColEnd) Then Exit Sub

    If m_RowStart <= 0 Or m_ColStart <= 0 Then
        MsgBox "Button: invalid row/column start for control '" & m_ControlName & "'.", vbExclamation
        Exit Sub
    End If

    If m_RowEnd < m_RowStart Then
        MsgBox "Button: control '" & m_ControlName & "' has invalid spanRows range.", vbExclamation
        Exit Sub
    End If

    If m_ColEnd < m_ColStart Then
        MsgBox "Button: control '" & m_ControlName & "' has invalid spanCells range.", vbExclamation
        Exit Sub
    End If

    m_IsConfigured = True
End Sub

Private Sub obj_IControl_Render(ByVal wb As Workbook)
    Dim ws As Worksheet
    Dim shp As Shape
    Dim buttonName As String
    Dim targetRange As Range

    If Not m_IsConfigured Then
        MsgBox "Button: control '" & m_ControlName & "' is not configured.", vbExclamation
        Exit Sub
    End If

    If wb Is Nothing Then Set wb = ThisWorkbook
    If wb Is Nothing Then
        MsgBox "Button: workbook is not specified for control '" & m_ControlName & "'.", vbExclamation
        Exit Sub
    End If

    Set ws = mp_GetWorksheetByName(wb, m_LayoutSheet)
    If ws Is Nothing Then
        MsgBox "Button: sheet '" & m_LayoutSheet & "' was not found for control '" & m_ControlName & "'.", vbExclamation
        Exit Sub
    End If

    On Error GoTo EH_RANGE
    Set targetRange = ws.Range(ws.Cells(m_RowStart, m_ColStart), ws.Cells(m_RowEnd, m_ColEnd))
    On Error GoTo 0

    If targetRange Is Nothing Then
        MsgBox "Button: failed to resolve target range for control '" & m_ControlName & "'.", vbExclamation
        Exit Sub
    End If
    If targetRange.Width <= 0# Or targetRange.Height <= 0# Then
        MsgBox "Button: target range has non-positive width/height for control '" & m_ControlName & "'.", vbExclamation
        Exit Sub
    End If

    buttonName = "btn_" & m_ControlName

    On Error Resume Next
    ws.Shapes(buttonName).Delete
    On Error GoTo EH_BUTTON

    Set shp = ws.Shapes.AddShape(msoShapeRoundedRectangle, targetRange.Left, targetRange.Top, targetRange.Width, targetRange.Height)
    shp.Name = buttonName
    shp.Placement = xlMoveAndSize
    shp.OnAction = m_OnClickMacroRef

    On Error Resume Next
    shp.TextFrame2.TextRange.Text = m_CaptionText
    shp.TextFrame2.VerticalAnchor = msoAnchorMiddle
    shp.TextFrame2.TextRange.ParagraphFormat.Alignment = msoAlignCenter
    shp.TextFrame.Characters.Text = m_CaptionText
    shp.TextFrame.HorizontalAlignment = xlHAlignCenter
    shp.TextFrame.VerticalAlignment = xlVAlignCenter
    shp.Tags.Add "pn.control", m_ControlName
    shp.Tags.Add "pn.style", m_StyleName
    On Error GoTo EH_BUTTON

    Exit Sub

EH_BUTTON:
    MsgBox "Button: failed to render control '" & m_ControlName & "': " & Err.Description, vbExclamation
    Exit Sub

EH_RANGE:
    MsgBox "Button: failed to resolve target range for control '" & m_ControlName & "': " & Err.Description, vbExclamation
End Sub

Private Function obj_IControl_SupportsAttribute(ByVal attrName As String) As Boolean
    Select Case LCase$(Trim$(attrName))
        Case "caption", "onclick"
            obj_IControl_SupportsAttribute = True
    End Select
End Function

Private Function mp_TryReadLayoutLongAttr( _
    ByVal controlNode As Object, _
    ByVal attrName As String, _
    ByRef outValue As Long _
) As Boolean
    Dim rawText As String

    rawText = Trim$(ex_XmlCore.m_NodeAttrText(controlNode, attrName))
    If Len(rawText) = 0 Then
        MsgBox "Button: runtime layout attribute '" & attrName & "' is missing for control '" & m_ControlName & "'.", vbExclamation
        Exit Function
    End If

    If Not IsNumeric(rawText) Then
        MsgBox "Button: runtime layout attribute '" & attrName & "' must be numeric for control '" & m_ControlName & "'.", vbExclamation
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

Public Property Get DefaultCaption() As String
    DefaultCaption = DEFAULT_CAPTION
End Property
