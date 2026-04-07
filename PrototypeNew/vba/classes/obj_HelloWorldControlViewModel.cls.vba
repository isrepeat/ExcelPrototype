VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "obj_HelloWorldControlViewModel"
Option Explicit
Implements obj_IControl

Private Const DEFAULT_SHEET_NAME As String = "Dev"
Private Const DEFAULT_TARGET_CELL As String = "A1"
Private Const DEFAULT_TEXT As String = "Hello, World!"

Private m_ControlName As String
Private m_SheetName As String
Private m_TargetCell As String
Private m_Text As String

Private Sub obj_IControl_Configure(ByVal controlNode As Object)
    If controlNode Is Nothing Then
        MsgBox "HelloWorld: control node is not specified.", vbExclamation
        Exit Sub
    End If

    m_ControlName = Trim$(ex_XmlCore.m_NodeAttrText(controlNode, "name"))
    If Len(m_ControlName) = 0 Then
        m_ControlName = "helloWorld"
    End If

    m_SheetName = Trim$(ex_XmlCore.m_NodeAttrText(controlNode, "sheet"))
    If Len(m_SheetName) = 0 Then m_SheetName = DEFAULT_SHEET_NAME

    m_TargetCell = Trim$(ex_XmlCore.m_NodeAttrText(controlNode, "targetCell"))
    If Len(m_TargetCell) = 0 Then m_TargetCell = DEFAULT_TARGET_CELL

    m_Text = CStr(ex_XmlCore.m_NodeAttrText(controlNode, "text"))
    If Len(m_Text) = 0 Then m_Text = DEFAULT_TEXT
End Sub

Private Sub obj_IControl_Render(ByVal wb As Workbook)
    Dim ws As Worksheet

    If wb Is Nothing Then Set wb = ThisWorkbook
    If wb Is Nothing Then
        MsgBox "HelloWorld: workbook is not specified for control '" & m_ControlName & "'.", vbExclamation
        Exit Sub
    End If

    Set ws = mp_GetWorksheetByName(wb, m_SheetName)
    If ws Is Nothing Then
        MsgBox "HelloWorld: sheet '" & m_SheetName & "' was not found for control '" & m_ControlName & "'.", vbExclamation
        Exit Sub
    End If

    On Error GoTo EH_CELL
    ws.Range(m_TargetCell).Value2 = m_Text
    Exit Sub
EH_CELL:
    MsgBox "HelloWorld: invalid targetCell '" & m_TargetCell & "' for control '" & m_ControlName & "'.", vbExclamation
End Sub

Private Function mp_GetWorksheetByName(ByVal wb As Workbook, ByVal sheetName As String) As Worksheet
    On Error Resume Next
    Set mp_GetWorksheetByName = wb.Worksheets(sheetName)
    On Error GoTo 0
End Function
