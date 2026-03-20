Attribute VB_Name = "ex_Navigation"
Option Explicit

Public Sub m_ReturnToDevPage()
    On Error GoTo EH
    ws_Dev.Activate
    Exit Sub
EH:
    Err.Raise vbObjectError + 1725, "ex_Navigation", "Dev sheet is not available for navigation: " & Err.Description
End Sub

Public Sub m_ScrollToPageTop()
    Dim ws As Worksheet

    On Error GoTo EH

    Set ws = ActiveSheet
    If ws Is Nothing Then
        Err.Raise vbObjectError + 1726, "ex_Navigation", "Active sheet is not available for top scroll."
    End If

    ws.Activate
    Application.Goto ws.Cells(1, 1), True
    Exit Sub
EH:
    Err.Raise vbObjectError + 1727, "ex_Navigation", "Unable to scroll to page top: " & Err.Description
End Sub
