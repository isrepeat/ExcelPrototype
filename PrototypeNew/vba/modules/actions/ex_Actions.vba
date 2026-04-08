Attribute VB_Name = "ex_Actions"
Option Explicit

Public Sub m_LoadPrototypeNewUi()
    Dim activeSheetObj As Object
    Dim ws As Worksheet

    Set activeSheetObj = ThisWorkbook.ActiveSheet
    If activeSheetObj Is Nothing Then
        MsgBox "PrototypeNew: active sheet is not specified.", vbExclamation
        Exit Sub
    End If

    If Not TypeOf activeSheetObj Is Worksheet Then
        MsgBox "PrototypeNew: active sheet is not a worksheet.", vbExclamation
        Exit Sub
    End If

    Set ws = activeSheetObj
    ex_SheetRenderer.m_RenderWorksheet ws
End Sub

Public Sub m_SetStateSimpleTestDefault()
    ex_State.m_SetText ex_State.STATE_ACTIVE_MODE, "SimpleTest"
    ex_State.m_SetText ex_State.STATE_ACTIVE_PROFILE, "Default"
End Sub
