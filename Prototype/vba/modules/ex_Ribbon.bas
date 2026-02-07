Attribute VB_Name = "ex_Ribbon"
Option Explicit

Public g_Ribbon As Object

' Вызывается Excel при загрузке Ribbon
Public Sub m_Ribbon_OnLoad(ByVal ribbon As Object)
    Set g_Ribbon = ribbon
End Sub

' Тестовая кнопка
Public Sub m_OnPing(ByVal control As Object)
    MsgBox "Ribbon works!"
End Sub