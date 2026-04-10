VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "obj_ControlLayout"
Option Explicit

Private m_StyleName As String
Private m_LayoutSheet As String
Private m_RowStart As Long
Private m_ColStart As Long
Private m_RowEnd As Long
Private m_ColEnd As Long

' //
' // API
' //
Public Function m_TryReadFromNode( _
    ByVal controlNode As Object, _
    ByVal controlTypeLabel As String, _
    ByVal controlName As String, _
    Optional ByVal styleAttrName As String = "style" _
) As Boolean
    If controlNode Is Nothing Then
        MsgBox controlTypeLabel & ": control node is not specified.", vbExclamation
        Exit Function
    End If

    m_StyleName = Trim$(CStr(ex_XmlCore.m_NodeAttrText(controlNode, styleAttrName)))

    m_LayoutSheet = Trim$(CStr(ex_XmlCore.m_NodeAttrText(controlNode, "__layoutSheet")))
    If Len(m_LayoutSheet) = 0 Then
        MsgBox controlTypeLabel & ": runtime layout sheet is missing for control '" & controlName & "'.", vbExclamation
        Exit Function
    End If

    If Not mp_TryReadLayoutLongAttr(controlNode, "__layoutRowStart", controlTypeLabel, controlName, m_RowStart) Then Exit Function
    If Not mp_TryReadLayoutLongAttr(controlNode, "__layoutColStart", controlTypeLabel, controlName, m_ColStart) Then Exit Function
    If Not mp_TryReadLayoutLongAttr(controlNode, "__layoutRowEnd", controlTypeLabel, controlName, m_RowEnd) Then Exit Function
    If Not mp_TryReadLayoutLongAttr(controlNode, "__layoutColEnd", controlTypeLabel, controlName, m_ColEnd) Then Exit Function

    If m_RowStart <= 0 Or m_ColStart <= 0 Then
        MsgBox controlTypeLabel & ": invalid row/column start for control '" & controlName & "'.", vbExclamation
        Exit Function
    End If
    If m_RowEnd < m_RowStart Then
        MsgBox controlTypeLabel & ": invalid spanRows range for control '" & controlName & "'.", vbExclamation
        Exit Function
    End If
    If m_ColEnd < m_ColStart Then
        MsgBox controlTypeLabel & ": invalid spanCells range for control '" & controlName & "'.", vbExclamation
        Exit Function
    End If

    m_TryReadFromNode = True
End Function

Public Property Get StyleName() As String
    StyleName = m_StyleName
End Property

Public Property Get LayoutSheet() As String
    LayoutSheet = m_LayoutSheet
End Property

Public Property Get RowStart() As Long
    RowStart = m_RowStart
End Property

Public Property Get ColStart() As Long
    ColStart = m_ColStart
End Property

Public Property Get RowEnd() As Long
    RowEnd = m_RowEnd
End Property

Public Property Get ColEnd() As Long
    ColEnd = m_ColEnd
End Property

' //
' // Internal
' //
Private Function mp_TryReadLayoutLongAttr( _
    ByVal controlNode As Object, _
    ByVal attrName As String, _
    ByVal controlTypeLabel As String, _
    ByVal controlName As String, _
    ByRef outValue As Long _
) As Boolean
    Dim rawText As String

    rawText = Trim$(CStr(ex_XmlCore.m_NodeAttrText(controlNode, attrName)))
    If Len(rawText) = 0 Then
        MsgBox controlTypeLabel & ": runtime layout attribute '" & attrName & "' is missing for control '" & controlName & "'.", vbExclamation
        Exit Function
    End If
    If Not IsNumeric(rawText) Then
        MsgBox controlTypeLabel & ": runtime layout attribute '" & attrName & "' must be numeric for control '" & controlName & "'.", vbExclamation
        Exit Function
    End If

    outValue = CLng(rawText)
    mp_TryReadLayoutLongAttr = True
End Function

