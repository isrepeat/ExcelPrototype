Attribute VB_Name = "ex_UiLoader"
Option Explicit

Public Sub m_LoadPrototypeNewUi(Optional ByVal wb As Workbook)
    If wb Is Nothing Then Set wb = ThisWorkbook
    If wb Is Nothing Then
        MsgBox "PrototypeNew: workbook is not specified.", vbExclamation
        Exit Sub
    End If

    ex_ControlRuntime.m_RenderDevLayout wb
End Sub
