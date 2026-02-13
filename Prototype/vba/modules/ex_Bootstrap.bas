Attribute VB_Name = "ex_Bootstrap"
Option Explicit

' Safe bootstrap that can be called from ThisWorkbook.Workbook_Open.
Public Sub Bootstrap_Open()
    On Error Resume Next
    Application.Run "ex_Presets.m_EnsureProfileDropdown", ws_Dev
    On Error GoTo 0
End Sub

