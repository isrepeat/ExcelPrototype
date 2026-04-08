Option Explicit

Private Sub Workbook_Open()
    Dim activeSheetObj As Object
    Dim ws As Worksheet

    On Error GoTo EH

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
    Exit Sub
EH:
    MsgBox "PrototypeNew: Workbook_Open failed: " & Err.Description, vbExclamation
End Sub
