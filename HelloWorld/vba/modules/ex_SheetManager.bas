Attribute VB_Name = "ex_SheetManager"
Option Explicit

' =============================================================================
' ex_SheetManager
' =============================================================================
' Purpose:
'   Manage lifecycle of temporary working sheets (g_Old, g_New, g_Result)
'
' Responsibilities:
'   - Delete temporary worksheets
'   - Provide utility functions for sheet operations
' =============================================================================

Public Sub DeleteResultSheets()
    Dim sheetNames As Variant
    Dim i As Long
    
    sheetNames = Array("g_Old", "g_New", "g_Result")
    
    On Error Resume Next
    For i = LBound(sheetNames) To UBound(sheetNames)
        Application.DisplayAlerts = False
        ThisWorkbook.Worksheets(sheetNames(i)).Delete
        Application.DisplayAlerts = True
    Next i
    On Error GoTo 0
End Sub
