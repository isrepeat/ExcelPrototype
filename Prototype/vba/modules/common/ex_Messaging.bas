Attribute VB_Name = "ex_Messaging"
Option Explicit

' =============================================================================
' Status bar notification
' =============================================================================

Public Sub m_ShowNotice(ByVal msg As String, Optional ByVal seconds As Double = 2)
    Application.StatusBar = msg
    Application.OnTime Now + TimeSerial(0, 0, seconds), "ex_Messaging.m_ClearStatusBar"
End Sub

Public Sub m_ClearStatusBar()
    ' Очищает статус бар
    Application.StatusBar = False
End Sub
