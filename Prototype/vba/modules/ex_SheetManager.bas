Attribute VB_Name = "ex_SheetManager"
Option Explicit

' =============================================================================
' ex_SheetManager
' =============================================================================
' Назначение:
'   Управление жизненным циклом временных рабочих листов (g_Old, g_New, g_Result)
'
' Обязанности:
'   - Удалять временные листы
'   - Предоставлять утилиты для операций с листами
' =============================================================================

Public Sub DeleteResultSheets()
    Dim ws As Worksheet
    Dim i As Long
    
    Application.DisplayAlerts = False
    On Error Resume Next
    
    ' Проходим по листам в обратном порядке, чтобы не пропустить при удалении
    For i = ThisWorkbook.Worksheets.Count To 1 Step -1
        Set ws = ThisWorkbook.Worksheets(i)
        If Left(ws.Name, 2) = "g_" Then
            ws.Delete
        End If
    Next i
    
    On Error GoTo 0
    Application.DisplayAlerts = True
End Sub

Public Sub ex_ApplyDefaultSheetView(ws As Worksheet)

    ws.Activate
    ActiveWindow.Zoom = 115

End Sub