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
Private m_Layout As obj_ControlLayout
Private m_CaptionText As String
Private m_OnClickMacroRef As String
Private m_RuntimeControlKey As String
Private m_IsConfigured As Boolean

Private Sub obj_IControl_Configure(ByVal controlNode As Object)
    m_IsConfigured = False
    Set m_Layout = Nothing
    m_RuntimeControlKey = vbNullString

    If controlNode Is Nothing Then
        MsgBox "Button: control node is not specified.", vbExclamation
        Exit Sub
    End If

    m_ControlName = Trim$(ex_XmlCore.m_NodeAttrText(controlNode, "name"))
    If Len(m_ControlName) = 0 Then m_ControlName = "button"

    m_CaptionRaw = CStr(ex_XmlCore.m_NodeAttrText(controlNode, "caption"))
    If Len(m_CaptionRaw) = 0 Then m_CaptionRaw = DEFAULT_CAPTION

    m_OnClickRaw = CStr(ex_XmlCore.m_NodeAttrText(controlNode, "onClick"))
    If Len(Trim$(m_OnClickRaw)) = 0 Then
        MsgBox "Button: onClick is required for control '" & m_ControlName & "'.", vbExclamation
        Exit Sub
    End If

    If Not ex_BindingRuntime.m_TryResolveTextBinding(m_CaptionRaw, Me, m_CaptionText) Then Exit Sub
    If Not ex_BindingRuntime.m_TryResolveMacroBinding(m_OnClickRaw, Me, m_OnClickMacroRef) Then Exit Sub

    Set m_Layout = New obj_ControlLayout
    If Not m_Layout.m_TryReadFromNode(controlNode, "Button", m_ControlName, "style") Then Exit Sub
    m_RuntimeControlKey = "button|" & LCase$(Trim$(m_Layout.LayoutSheet & "|" & m_ControlName))

    m_IsConfigured = True
End Sub

Private Sub obj_IControl_Render(ByVal wb As Workbook)
    Dim ws As Worksheet
    Dim shp As Shape
    Dim buttonName As String
    Dim targetRange As Range
    Dim metaMap As Object
    Dim callbackMacroRef As String

    If Not m_IsConfigured Then
        MsgBox "Button: control '" & m_ControlName & "' is not configured.", vbExclamation
        Exit Sub
    End If

    If wb Is Nothing Then Set wb = ThisWorkbook
    If wb Is Nothing Then
        MsgBox "Button: workbook is not specified for control '" & m_ControlName & "'.", vbExclamation
        Exit Sub
    End If

    Set ws = mp_GetWorksheetByName(wb, m_Layout.LayoutSheet)
    If ws Is Nothing Then
        MsgBox "Button: sheet '" & m_Layout.LayoutSheet & "' was not found for control '" & m_ControlName & "'.", vbExclamation
        Exit Sub
    End If

    On Error GoTo EH_RANGE
    Set targetRange = ws.Range(ws.Cells(m_Layout.RowStart, m_Layout.ColStart), ws.Cells(m_Layout.RowEnd, m_Layout.ColEnd))
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
    callbackMacroRef = mp_GetRuntimeCallbackMacroRef()
    If Len(callbackMacroRef) = 0 Then Exit Sub

    On Error Resume Next
    ws.Shapes(buttonName).Delete
    On Error GoTo EH_BUTTON

    Set shp = ws.Shapes.AddShape(msoShapeRoundedRectangle, targetRange.Left, targetRange.Top, targetRange.Width, targetRange.Height)
    shp.Name = buttonName
    shp.Placement = xlMoveAndSize
    shp.OnAction = callbackMacroRef

    If Not ex_ShapeClickDispatcher.m_RegisterControl(m_RuntimeControlKey, Me) Then Exit Sub
    If Not ex_ShapeClickDispatcher.m_RegisterShapeRoute(shp.Name, m_RuntimeControlKey, "m_RuntimeHandleClick", False) Then Exit Sub

    On Error Resume Next
    shp.TextFrame2.TextRange.Text = m_CaptionText
    shp.TextFrame2.VerticalAnchor = msoAnchorMiddle
    shp.TextFrame2.TextRange.ParagraphFormat.Alignment = msoAlignCenter
    shp.TextFrame.Characters.Text = m_CaptionText
    shp.TextFrame.HorizontalAlignment = xlHAlignCenter
    shp.TextFrame.VerticalAlignment = xlVAlignCenter

    Set metaMap = CreateObject("Scripting.Dictionary")
    metaMap.CompareMode = 1
    metaMap("pn.control") = m_ControlName
    If Len(Trim$(m_Layout.StyleName)) > 0 Then
        metaMap("pn.style") = m_Layout.StyleName
    Else
        metaMap("pn.style") = vbNullString
    End If
    If Not ex_ShapeMetaRuntime.m_TrySetShapeMetaValues(shp, metaMap) Then Exit Sub
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

Public Function m_RuntimeHandleClick() As Boolean
    On Error GoTo EH_CLICK
    Application.Run m_OnClickMacroRef
    m_RuntimeHandleClick = True
    Exit Function

EH_CLICK:
    MsgBox "Button: failed to execute onClick for control '" & m_ControlName & "': " & Err.Description, vbExclamation
End Function

Private Function mp_GetRuntimeCallbackMacroRef() As String
    mp_GetRuntimeCallbackMacroRef = mp_QualifyMacroName("ex_ShapeClickDispatcher.m_OnShapeClick")
End Function

Private Function mp_QualifyMacroName(ByVal macroName As String) As String
    Dim wbName As String

    macroName = Trim$(macroName)
    If Len(macroName) = 0 Then Exit Function
    If InStr(1, macroName, "!", vbBinaryCompare) > 0 Then
        mp_QualifyMacroName = macroName
        Exit Function
    End If

    wbName = ThisWorkbook.Name
    wbName = Replace$(wbName, "'", "''")
    mp_QualifyMacroName = "'" & wbName & "'!" & macroName
End Function

Private Function mp_GetWorksheetByName(ByVal wb As Workbook, ByVal sheetName As String) As Worksheet
    On Error Resume Next
    Set mp_GetWorksheetByName = wb.Worksheets(sheetName)
    On Error GoTo 0
End Function

Public Property Get DefaultCaption() As String
    DefaultCaption = DEFAULT_CAPTION
End Property
