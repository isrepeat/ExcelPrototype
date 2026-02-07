Attribute VB_Name = "ex_Ribbon"
Option Explicit

Public g_Ribbon As Object

' Вызывается Excel при загрузке Ribbon
Public Sub Ribbon_OnLoad(ByVal ribbon As Object)
    Set g_Ribbon = ribbon
End Sub

' Тестовая кнопка
Public Sub OnPing(ByVal control As Object)
    MsgBox "Ribbon works!"
End Sub