VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "obj_ButtonControlViewModel"
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
Private m_LayoutLeft As Double
Private m_LayoutTop As Double
Private m_LayoutWidth As Double
Private m_LayoutHeight As Double
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

    If Not mp_TryReadLayoutDoubleAttr(controlNode, "__layoutLeft", m_LayoutLeft) Then Exit Sub
    If Not mp_TryReadLayoutDoubleAttr(controlNode, "__layoutTop", m_LayoutTop) Then Exit Sub
    If Not mp_TryReadLayoutDoubleAttr(controlNode, "__layoutWidth", m_LayoutWidth) Then Exit Sub
    If Not mp_TryReadLayoutDoubleAttr(controlNode, "__layoutHeight", m_LayoutHeight) Then Exit Sub

    If m_LayoutWidth <= 0# Or m_LayoutHeight <= 0# Then
        MsgBox "Button: runtime layout width/height must be greater than zero for control '" & m_ControlName & "'.", vbExclamation
        Exit Sub
    End If

    m_IsConfigured = True
End Sub

Private Sub obj_IControl_Render(ByVal wb As Workbook)
    Dim ws As Worksheet
    Dim shp As Shape
    Dim buttonName As String

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

    buttonName = "btn_" & m_ControlName

    On Error Resume Next
    ws.Shapes(buttonName).Delete
    On Error GoTo EH_BUTTON

    Set shp = ws.Shapes.AddShape(msoShapeRoundedRectangle, m_LayoutLeft, m_LayoutTop, m_LayoutWidth, m_LayoutHeight)
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
End Sub

Private Function obj_IControl_SupportsAttribute(ByVal attrName As String) As Boolean
    Select Case LCase$(Trim$(attrName))
        Case "caption", "onclick"
            obj_IControl_SupportsAttribute = True
    End Select
End Function

Private Function mp_TryReadLayoutDoubleAttr( _
    ByVal controlNode As Object, _
    ByVal attrName As String, _
    ByRef outValue As Double _
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

    outValue = CDbl(rawText)
    mp_TryReadLayoutDoubleAttr = True
End Function

Private Function mp_GetWorksheetByName(ByVal wb As Workbook, ByVal sheetName As String) As Worksheet
    On Error Resume Next
    Set mp_GetWorksheetByName = wb.Worksheets(sheetName)
    On Error GoTo 0
End Function

Public Property Get DefaultCaption() As String
    DefaultCaption = DEFAULT_CAPTION
End Property