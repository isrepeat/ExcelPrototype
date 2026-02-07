Attribute VB_Name = "ex_TestData"
Option Explicit

Public Sub GenerateTestTables()
    Dim wsOld As Worksheet
    Dim wsNew As Worksheet
    
    Set wsOld = GetOrCreateWorksheet("Old")
    Set wsNew = GetOrCreateWorksheet("New")
    
    wsOld.Cells.Clear
    wsNew.Cells.Clear
    
    WriteHeaders wsOld
    WriteHeaders wsNew
    
    ' -----------------------------
    ' СТАРЫЕ ДАННЫЕ (Old)
    ' -----------------------------
    ' Id | Name   | Value
    wsOld.Cells(2, 1).Value = "1"
    wsOld.Cells(2, 2).Value = "Alpha"
    wsOld.Cells(2, 3).Value = "10"
    
    wsOld.Cells(3, 1).Value = "2"
    wsOld.Cells(3, 2).Value = "Beta"
    wsOld.Cells(3, 3).Value = "20"
    
    wsOld.Cells(4, 1).Value = "3"
    wsOld.Cells(4, 2).Value = "Gamma"
    wsOld.Cells(4, 3).Value = "30"
    
    wsOld.Cells(5, 1).Value = "4"
    wsOld.Cells(5, 2).Value = "Delta"
    wsOld.Cells(5, 3).Value = "40"
    
    ' -----------------------------
    ' НОВЫЕ ДАННЫЕ (New)
    ' -----------------------------
    ' 1 остается без изменений
    wsNew.Cells(2, 1).Value = "1"
    wsNew.Cells(2, 2).Value = "Alpha"
    wsNew.Cells(2, 3).Value = "10"
    
    ' 2 изменено значение
    wsNew.Cells(3, 1).Value = "2"
    wsNew.Cells(3, 2).Value = "Beta"
    wsNew.Cells(3, 3).Value = "25"
    
    ' 3 удалено (отсутствует в New)
    
    ' 4 изменено имя
    wsNew.Cells(4, 1).Value = "4"
    wsNew.Cells(4, 2).Value = "DeltaX"
    wsNew.Cells(4, 3).Value = "40"
    
    ' 5 добавлено
    wsNew.Cells(5, 1).Value = "5"
    wsNew.Cells(5, 2).Value = "Epsilon"
    wsNew.Cells(5, 3).Value = "50"

     ' 6 added
    wsNew.Cells(6, 1).Value = "6"
    wsNew.Cells(6, 2).Value = "Zetta"
    wsNew.Cells(6, 3).Value = "60"

    ' 7 added
    wsNew.Cells(7, 1).Value = "7"
    wsNew.Cells(7, 2).Value = "Tetta"
    wsNew.Cells(7, 3).Value = "70"
    
    wsOld.Columns.AutoFit
    wsNew.Columns.AutoFit

    ex_SheetTheme.ApplyDarkThemeToSheet wsOld
    ex_SheetTheme.ApplyDarkThemeToSheet wsNew
End Sub

Private Sub WriteHeaders(ByVal ws As Worksheet)
    ws.Cells(1, 1).Value = "Id"
    ws.Cells(1, 2).Value = "Name"
    ws.Cells(1, 3).Value = "Value"
End Sub

Private Function GetOrCreateWorksheet(ByVal sheetName As String) As Worksheet
    Dim ws As Worksheet
    Dim fullName As String
    
    ' Add g_ prefix to sheet names
    fullName = "g_" & sheetName
    
    For Each ws In ThisWorkbook.Worksheets
        If StrComp(ws.Name, fullName, vbTextCompare) = 0 Then
            Set GetOrCreateWorksheet = ws
            Exit Function
        End If
    Next ws
    
    Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
    ws.Name = fullName
    Call ex_ApplyDefaultSheetView(ws)
    
    Set GetOrCreateWorksheet = ws
End Function
