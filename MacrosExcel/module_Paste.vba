Option Explicit

'=========================
' Undo storage
'=========================
Private UndoSheetName As String
Private UndoAddresses() As String
Private UndoValues() As Variant
Private UndoCount As Long

'=========================
' Windows Clipboard API
'=========================
#If VBA7 Then
Private Declare PtrSafe Function private_OpenClipboard Lib "user32" Alias "OpenClipboard" (ByVal hwnd As LongPtr) As Long
Private Declare PtrSafe Function private_CloseClipboard Lib "user32" Alias "CloseClipboard" () As Long
Private Declare PtrSafe Function private_GetClipboardData Lib "user32" Alias "GetClipboardData" (ByVal wFormat As Long) As LongPtr
Private Declare PtrSafe Function private_GlobalLock Lib "kernel32" Alias "GlobalLock" (ByVal hMem As LongPtr) As LongPtr
Private Declare PtrSafe Function private_GlobalUnlock Lib "kernel32" Alias "GlobalUnlock" (ByVal hMem As LongPtr) As Long
Private Declare PtrSafe Function private_lstrlenW Lib "kernel32" Alias "lstrlenW" (ByVal lpString As LongPtr) As Long
Private Declare PtrSafe Sub private_CopyMemory Lib "kernel32" Alias "RtlMoveMemory" ( _
    ByVal Destination As LongPtr, _
    ByVal Source As LongPtr, _
    ByVal Length As LongPtr)
#Else
Private Declare Function private_OpenClipboard Lib "user32" Alias "OpenClipboard" (ByVal hwnd As Long) As Long
Private Declare Function private_CloseClipboard Lib "user32" Alias "CloseClipboard" () As Long
Private Declare Function private_GetClipboardData Lib "user32" Alias "GetClipboardData" (ByVal wFormat As Long) As Long
Private Declare Function private_GlobalLock Lib "kernel32" Alias "GlobalLock" (ByVal hMem As Long) As Long
Private Declare Function private_GlobalUnlock Lib "kernel32" Alias "GlobalUnlock" (ByVal hMem As Long) As Long
Private Declare Function private_lstrlenW Lib "kernel32" Alias "lstrlenW" (ByVal lpString As Long) As Long
Private Declare Sub private_CopyMemory Lib "kernel32" Alias "RtlMoveMemory" ( _
    ByVal Destination As Long, _
    ByVal Source As Long, _
    ByVal Length As Long)
#End If
Private Const CF_UNICODETEXT As Long = 13

'==========================================================
' MAIN
'==========================================================
Public Sub fn_PasteClipboardTextToVisibleCells()
    Dim txt As String
    Dim arr As Variant
    Dim visibleCells As Range
    Dim cell As Range
    Dim i As Long
    Dim clipboardCount As Long
    Dim targetCount As Long
    Dim pasteCount As Long
    Dim answer As VbMsgBoxResult

    txt = private_GetClipboardUnicodeText()
    If Len(txt) = 0 Then
        MsgBox "Clipboard is empty.", vbExclamation
        Exit Sub
    End If

    arr = private_ParseExcelClipboardFirstColumn(txt)

    On Error Resume Next
    If Selection.Cells.CountLarge = 1 Then
        Set visibleCells = private_GetVisibleTargetFromSingleCell(Selection.Cells(1))
    Else
        Set visibleCells = Selection.SpecialCells(xlCellTypeVisible)
    End If
    On Error GoTo 0

    If visibleCells Is Nothing Then
        MsgBox "No visible target cells selected.", vbExclamation
        Exit Sub
    End If

    clipboardCount = CLng(UBound(arr) - LBound(arr) + 1)
    targetCount = CLng(visibleCells.Cells.CountLarge)
    pasteCount = WorksheetFunction.Min(clipboardCount, targetCount)

    If clipboardCount <> targetCount Then
        answer = MsgBox( _
            "Clipboard values: " & clipboardCount & vbCrLf & _
            "Visible target cells: " & targetCount & vbCrLf & vbCrLf & _
            "Paste first " & pasteCount & " values?", _
            vbQuestion + vbYesNo, _
            "Count mismatch")
        If answer = vbNo Then Exit Sub
    End If

    ' Save old values for Undo
    UndoSheetName = ActiveSheet.Name
    UndoCount = pasteCount
    ReDim UndoAddresses(1 To UndoCount)
    ReDim UndoValues(1 To UndoCount)

    i = 1
    For Each cell In visibleCells.Cells
        If i > UndoCount Then Exit For
        UndoAddresses(i) = cell.Address(False, False)
        UndoValues(i) = cell.Value
        i = i + 1
    Next cell

    ' Paste
    i = LBound(arr)
    For Each cell In visibleCells.Cells
        If i > UBound(arr) Then Exit For
        cell.Value = arr(i)
        i = i + 1
    Next cell

    Application.OnUndo _
        "Undo visible paste", _
        "fn_UndoPasteVisibleCells"
End Sub

'==========================================================
' Undo
'==========================================================
Public Sub fn_UndoPasteVisibleCells()
    Dim ws As Worksheet
    Dim i As Long
    If UndoCount = 0 Then Exit Sub
    Set ws = Worksheets(UndoSheetName)
    For i = 1 To UndoCount
        ws.Range(UndoAddresses(i)).Value = UndoValues(i)
    Next i
    UndoCount = 0
End Sub

'==========================================================
' Single active cell -> all visible cells below
'==========================================================
Private Function private_GetVisibleTargetFromSingleCell(ByVal startCell As Range) As Range
    Dim lo As ListObject
    Dim colIndex As Long
    Dim firstRow As Long
    Dim rowCount As Long
    Dim rng As Range

    On Error Resume Next
    Set lo = startCell.ListObject
    On Error GoTo 0

    If Not lo Is Nothing Then
        colIndex = startCell.Column - lo.DataBodyRange.Columns(1).Column + 1
        firstRow = startCell.Row - lo.DataBodyRange.Row + 1
        rowCount = lo.DataBodyRange.Rows.Count - firstRow + 1

        Set rng = _
            lo.DataBodyRange.Columns(colIndex) _
            .Cells(firstRow, 1) _
            .Resize(rowCount, 1)

        Set private_GetVisibleTargetFromSingleCell = _
            rng.SpecialCells(xlCellTypeVisible)
    Else
        Set rng = Range( _
            startCell, _
            Cells(Rows.Count, startCell.Column).End(xlUp))
        Set private_GetVisibleTargetFromSingleCell = _
            rng.SpecialCells(xlCellTypeVisible)
    End If
End Function

'==========================================================
' Read Unicode text from Windows Clipboard
'==========================================================
Private Function private_GetClipboardUnicodeText() As String
    Dim hData As LongPtr
    Dim pData As LongPtr
    Dim length As Long
    Dim result As String

    If private_OpenClipboard(0) = 0 Then Exit Function

    hData = private_GetClipboardData(CF_UNICODETEXT)
    If hData <> 0 Then
        pData = private_GlobalLock(hData)
        If pData <> 0 Then
            length = private_lstrlenW(pData)
            If length > 0 Then
                result = String$(length, vbNullChar)
                private_CopyMemory _
                    StrPtr(result), _
                    pData, _
                    length * 2
            End If
            private_GlobalUnlock hData
        End If
    End If

    private_CloseClipboard
    private_GetClipboardUnicodeText = result
End Function

'==========================================================
' Parse Excel clipboard first column
'==========================================================
Private Function private_ParseExcelClipboardFirstColumn(ByVal txt As String) As Variant
    Dim result() As String
    Dim current As String
    Dim i As Long
    Dim n As Long
    Dim ch As String
    Dim inQuotes As Boolean

    Do While Len(txt) > 0 _
        And (Right$(txt, 1) = vbCr _
        Or Right$(txt, 1) = vbLf)
        txt = Left$(txt, Len(txt) - 1)
    Loop

    ReDim result(0)

    For i = 1 To Len(txt)
        ch = Mid$(txt, i, 1)

        If ch = """" Then
            If inQuotes _
                And i < Len(txt) _
                And Mid$(txt, i + 1, 1) = """" Then
                current = current & """"
                i = i + 1
            Else
                inQuotes = Not inQuotes
            End If

        ElseIf (ch = vbCr Or ch = vbLf) _
            And Not inQuotes Then
            result(n) = current
            current = ""
            n = n + 1
            ReDim Preserve result(n)

            If ch = vbCr _
                And i < Len(txt) _
                And Mid$(txt, i + 1, 1) = vbLf Then
                i = i + 1
            End If

        ElseIf ch = vbTab _
            And Not inQuotes Then
            Do While i < Len(txt)
                i = i + 1
                ch = Mid$(txt, i, 1)
                If (ch = vbCr Or ch = vbLf) _
                    And Not inQuotes Then
                    i = i - 1
                    Exit Do
                End If
            Loop
        Else
            current = current & ch
        End If
    Next i

    result(n) = current
    private_ParseExcelClipboardFirstColumn = result
End Function
