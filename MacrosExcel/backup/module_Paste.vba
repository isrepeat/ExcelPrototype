Option Explicit

'==========================================================
' Undo storage: paste to visible cells
'==========================================================
Private UndoSheetName As String
Private UndoAddresses() As String
Private UndoValues() As Variant
Private UndoCount As Long

'==========================================================
' Undo storage: reinterpret selected cells
'==========================================================
Private ReinterpretUndoSheetName As String
Private ReinterpretUndoAddresses() As String
Private ReinterpretUndoValues() As Variant
Private ReinterpretUndoFormats() As String
Private ReinterpretUndoCount As Long

'==========================================================
' Windows Clipboard API
'==========================================================
#If VBA7 Then

Private Declare PtrSafe Function private_OpenClipboard _
    Lib "user32" Alias "OpenClipboard" ( _
    ByVal hwnd As LongPtr) As Long

Private Declare PtrSafe Function private_CloseClipboard _
    Lib "user32" Alias "CloseClipboard" () As Long

Private Declare PtrSafe Function private_GetClipboardData _
    Lib "user32" Alias "GetClipboardData" ( _
    ByVal wFormat As Long) As LongPtr

Private Declare PtrSafe Function private_GlobalLock _
    Lib "kernel32" Alias "GlobalLock" ( _
    ByVal hMem As LongPtr) As LongPtr

Private Declare PtrSafe Function private_GlobalUnlock _
    Lib "kernel32" Alias "GlobalUnlock" ( _
    ByVal hMem As LongPtr) As Long

Private Declare PtrSafe Function private_lstrlenW _
    Lib "kernel32" Alias "lstrlenW" ( _
    ByVal lpString As LongPtr) As Long

Private Declare PtrSafe Sub private_CopyMemory _
    Lib "kernel32" Alias "RtlMoveMemory" ( _
    ByVal Destination As LongPtr, _
    ByVal Source As LongPtr, _
    ByVal Length As LongPtr)

#Else

Private Declare Function private_OpenClipboard _
    Lib "user32" Alias "OpenClipboard" ( _
    ByVal hwnd As Long) As Long

Private Declare Function private_CloseClipboard _
    Lib "user32" Alias "CloseClipboard" () As Long

Private Declare Function private_GetClipboardData _
    Lib "user32" Alias "GetClipboardData" ( _
    ByVal wFormat As Long) As Long

Private Declare Function private_GlobalLock _
    Lib "kernel32" Alias "GlobalLock" ( _
    ByVal hMem As Long) As Long

Private Declare Function private_GlobalUnlock _
    Lib "kernel32" Alias "GlobalUnlock" ( _
    ByVal hMem As Long) As Long

Private Declare Function private_lstrlenW _
    Lib "kernel32" Alias "lstrlenW" ( _
    ByVal lpString As Long) As Long

Private Declare Sub private_CopyMemory _
    Lib "kernel32" Alias "RtlMoveMemory" ( _
    ByVal Destination As Long, _
    ByVal Source As Long, _
    ByVal Length As Long)

#End If

Private Const CF_UNICODETEXT As Long = 13

'==========================================================
' MAIN: Paste clipboard first column to visible cells
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
        MsgBox "Буфер обмена пуст.", vbExclamation
        Exit Sub
    End If

    arr = private_ParseExcelClipboardFirstColumn(txt)

    On Error Resume Next

    If Selection.Cells.CountLarge = 1 Then
        Set visibleCells = _
            private_GetVisibleTargetFromSingleCell( _
                Selection.Cells(1))
    Else
        Set visibleCells = _
            Selection.SpecialCells(xlCellTypeVisible)
    End If

    On Error GoTo 0

    If visibleCells Is Nothing Then
        MsgBox _
            "Не выбраны видимые целевые ячейки.", _
            vbExclamation
        Exit Sub
    End If

    clipboardCount = _
        CLng(UBound(arr) - LBound(arr) + 1)

    targetCount = _
        CLng(visibleCells.Cells.CountLarge)

    pasteCount = _
        WorksheetFunction.Min( _
            clipboardCount, _
            targetCount)

    If clipboardCount <> targetCount Then

        answer = MsgBox( _
            "Значений в буфере: " & _
            clipboardCount & vbCrLf & _
            "Видимых целевых ячеек: " & _
            targetCount & vbCrLf & vbCrLf & _
            "Вставить первые " & _
            pasteCount & " значений?", _
            vbQuestion + vbYesNo, _
            "Количество не совпадает")

        If answer = vbNo Then Exit Sub
    End If

    '------------------------------------------------------
    ' Save old values for Undo
    '------------------------------------------------------
    UndoSheetName = ActiveSheet.Name
    UndoCount = pasteCount

    ReDim UndoAddresses(1 To UndoCount)
    ReDim UndoValues(1 To UndoCount)

    i = 1

    For Each cell In visibleCells.Cells

        If i > UndoCount Then Exit For

        UndoAddresses(i) = _
            cell.Address(False, False)

        UndoValues(i) = _
            cell.Value

        i = i + 1
    Next cell

    '------------------------------------------------------
    ' Paste
    '------------------------------------------------------
    i = LBound(arr)

    For Each cell In visibleCells.Cells

        If i > UBound(arr) Then Exit For

        cell.Value = arr(i)

        i = i + 1
    Next cell

    Application.OnUndo _
        "Отменить вставку в видимые ячейки", _
        "fn_UndoPasteVisibleCells"

End Sub

'==========================================================
' Undo paste to visible cells
'==========================================================
Public Sub fn_UndoPasteVisibleCells()

    Dim ws As Worksheet
    Dim i As Long

    If UndoCount = 0 Then Exit Sub

    On Error Resume Next
    Set ws = _
        ThisWorkbook.Worksheets(UndoSheetName)
    On Error GoTo 0

    If ws Is Nothing Then
        MsgBox _
            "Лист для отмены операции не найден.", _
            vbExclamation
        Exit Sub
    End If

    For i = 1 To UndoCount

        ws.Range(UndoAddresses(i)).Value = _
            UndoValues(i)

    Next i

    UndoCount = 0

End Sub

'==========================================================
' MAIN: Reinterpret selected cells as real text
'
' Решает ситуацию:
' - внутри ячейки хранится число 3448101133;
' - ячейке установлен текстовый формат;
' - но Excel продолжает показывать 3,4481E+09.
'
' Макрос:
' 1. берет внутреннее значение;
' 2. преобразует его в строку;
' 3. устанавливает текстовый формат;
' 4. записывает строку обратно.
'
' Формулы, ошибки и пустые ячейки пропускаются.
' Переносы строк в существующем тексте сохраняются.
'==========================================================
Public Sub fn_ReinterpretSelectedCellsAsText()

    Dim target As Range
    Dim cell As Range

    Dim valueAsText As String

    Dim processedCount As Long
    Dim skippedFormulaCount As Long
    Dim skippedErrorCount As Long
    Dim skippedEmptyCount As Long

    Dim totalCellCount As Double
    Dim answer As VbMsgBoxResult

    Dim oldCalculation As XlCalculation
    Dim oldScreenUpdating As Boolean
    Dim oldEnableEvents As Boolean

    If TypeName(Selection) <> "Range" Then
        MsgBox _
            "Сначала выделите нужные ячейки.", _
            vbExclamation
        Exit Sub
    End If

    Set target = Selection

    totalCellCount = target.Cells.CountLarge

    If totalCellCount > 100000 Then

        answer = MsgBox( _
            "Выделено ячеек: " & _
            Format$(totalCellCount, "#,##0") & vbCrLf & _
            "Обработка большого диапазона может занять время." & _
            vbCrLf & vbCrLf & _
            "Продолжить?", _
            vbQuestion + vbYesNo, _
            "Большой диапазон")

        If answer <> vbYes Then Exit Sub
    End If

    '------------------------------------------------------
    ' Remember current Excel state
    '------------------------------------------------------
    oldCalculation = Application.Calculation
    oldScreenUpdating = Application.ScreenUpdating
    oldEnableEvents = Application.EnableEvents

    On Error GoTo ErrorHandler

    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual
    Application.StatusBar = _
        "Подготовка переинтерпретации ячеек..."

    '------------------------------------------------------
    ' Prepare Undo storage
    '------------------------------------------------------
    ReinterpretUndoSheetName = ActiveSheet.Name
    ReinterpretUndoCount = 0

    ReDim ReinterpretUndoAddresses( _
        1 To CLng(totalCellCount))

    ReDim ReinterpretUndoValues( _
        1 To CLng(totalCellCount))

    ReDim ReinterpretUndoFormats( _
        1 To CLng(totalCellCount))

    '------------------------------------------------------
    ' Process selected cells
    '------------------------------------------------------
    For Each cell In target.Cells

        If IsEmpty(cell.Value2) Then

            skippedEmptyCount = _
                skippedEmptyCount + 1

        ElseIf cell.HasFormula Then

            skippedFormulaCount = _
                skippedFormulaCount + 1

        ElseIf IsError(cell.Value2) Then

            skippedErrorCount = _
                skippedErrorCount + 1

        Else

            ReinterpretUndoCount = _
                ReinterpretUndoCount + 1

            '----------------------------------------------
            ' Save current value and number format
            '----------------------------------------------
            ReinterpretUndoAddresses( _
                ReinterpretUndoCount) = _
                cell.Address(False, False)

            ReinterpretUndoValues( _
                ReinterpretUndoCount) = _
                cell.Value2

            ReinterpretUndoFormats( _
                ReinterpretUndoCount) = _
                cell.NumberFormat

            '----------------------------------------------
            ' Convert internal value to real string
            '
            ' Не используем cell.Text:
            ' cell.Text зависит от ширины столбца и может
            ' вернуть 3,4481E+09 либо #######.
            '----------------------------------------------
            valueAsText = _
                private_ValueToText(cell.Value2)

            '----------------------------------------------
            ' Set text format before writing the value
            '----------------------------------------------
            cell.NumberFormat = "@"
            cell.Value2 = valueAsText

            processedCount = _
                processedCount + 1

        End If

        If processedCount > 0 Then

            If processedCount Mod 500 = 0 Then

                Application.StatusBar = _
                    "Переинтерпретировано ячеек: " & _
                    Format$(processedCount, "#,##0")

            End If

        End If

    Next cell

    '------------------------------------------------------
    ' Shrink Undo arrays to actual number of processed cells
    '------------------------------------------------------
    If ReinterpretUndoCount > 0 Then

        ReDim Preserve ReinterpretUndoAddresses( _
            1 To ReinterpretUndoCount)

        ReDim Preserve ReinterpretUndoValues( _
            1 To ReinterpretUndoCount)

        ReDim Preserve ReinterpretUndoFormats( _
            1 To ReinterpretUndoCount)

        Application.OnUndo _
            "Отменить переинтерпретацию ячеек", _
            "fn_UndoReinterpretSelectedCells"

    End If

CleanExit:

    Application.StatusBar = False
    Application.Calculation = oldCalculation
    Application.EnableEvents = oldEnableEvents
    Application.ScreenUpdating = oldScreenUpdating

    If Err.Number = 0 Then

        MsgBox _
            "Готово." & vbCrLf & vbCrLf & _
            "Преобразовано в текст: " & _
            processedCount & vbCrLf & _
            "Пропущено пустых: " & _
            skippedEmptyCount & vbCrLf & _
            "Пропущено формул: " & _
            skippedFormulaCount & vbCrLf & _
            "Пропущено ошибок: " & _
            skippedErrorCount, _
            vbInformation, _
            "Переинтерпретация"

    End If

    Exit Sub

ErrorHandler:

    MsgBox _
        "Ошибка " & Err.Number & ":" & vbCrLf & _
        Err.Description, _
        vbCritical, _
        "Переинтерпретация"

    Resume CleanExit

End Sub

'==========================================================
' Undo reinterpretation
'==========================================================
Public Sub fn_UndoReinterpretSelectedCells()

    Dim ws As Worksheet
    Dim i As Long

    Dim oldCalculation As XlCalculation
    Dim oldScreenUpdating As Boolean
    Dim oldEnableEvents As Boolean

    If ReinterpretUndoCount = 0 Then Exit Sub

    On Error Resume Next

    Set ws = _
        ThisWorkbook.Worksheets( _
            ReinterpretUndoSheetName)

    On Error GoTo ErrorHandler

    If ws Is Nothing Then

        MsgBox _
            "Лист для отмены операции не найден.", _
            vbExclamation

        Exit Sub
    End If

    oldCalculation = Application.Calculation
    oldScreenUpdating = Application.ScreenUpdating
    oldEnableEvents = Application.EnableEvents

    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual
    Application.StatusBar = _
        "Отмена переинтерпретации..."

    For i = 1 To ReinterpretUndoCount

        With ws.Range( _
            ReinterpretUndoAddresses(i))

            ' Сначала возвращаем исходный формат
            .NumberFormat = _
                ReinterpretUndoFormats(i)

            ' Затем возвращаем исходное значение
            .Value2 = _
                ReinterpretUndoValues(i)

        End With

    Next i

    ReinterpretUndoCount = 0

CleanExit:

    Application.StatusBar = False
    Application.Calculation = oldCalculation
    Application.EnableEvents = oldEnableEvents
    Application.ScreenUpdating = oldScreenUpdating

    Exit Sub

ErrorHandler:

    MsgBox _
        "Ошибка отмены " & _
        Err.Number & ":" & vbCrLf & _
        Err.Description, _
        vbCritical, _
        "Отмена переинтерпретации"

    Resume CleanExit

End Sub

'==========================================================
' Convert internal cell value to text
'==========================================================
Private Function private_ValueToText( _
    ByVal sourceValue As Variant) As String

    Select Case VarType(sourceValue)

        Case vbString

            ' Уже существующий текст возвращается без изменений.
            ' Внутренние переносы строк сохраняются.
            private_ValueToText = sourceValue

        Case vbBoolean

            private_ValueToText = _
                CStr(sourceValue)

        Case vbByte, _
             vbInteger, _
             vbLong, _
             vbSingle, _
             vbDouble, _
             vbCurrency, _
             vbDecimal, _
             vbDate

            private_ValueToText = _
                CStr(sourceValue)

        Case Else

            private_ValueToText = _
                CStr(sourceValue)

    End Select

End Function

'==========================================================
' Single active cell -> all visible cells below
'==========================================================
Private Function private_GetVisibleTargetFromSingleCell( _
    ByVal startCell As Range) As Range

    Dim lo As ListObject
    Dim colIndex As Long
    Dim firstRow As Long
    Dim rowCount As Long
    Dim rng As Range

    On Error Resume Next
    Set lo = startCell.ListObject
    On Error GoTo 0

    If Not lo Is Nothing Then

        colIndex = _
            startCell.Column - _
            lo.DataBodyRange.Columns(1).Column + 1

        firstRow = _
            startCell.Row - _
            lo.DataBodyRange.Row + 1

        rowCount = _
            lo.DataBodyRange.Rows.Count - _
            firstRow + 1

        Set rng = _
            lo.DataBodyRange _
              .Columns(colIndex) _
              .Cells(firstRow, 1) _
              .Resize(rowCount, 1)

        On Error Resume Next

        Set private_GetVisibleTargetFromSingleCell = _
            rng.SpecialCells(xlCellTypeVisible)

        On Error GoTo 0

    Else

        Set rng = Range( _
            startCell, _
            Cells( _
                Rows.Count, _
                startCell.Column) _
            .End(xlUp))

        On Error Resume Next

        Set private_GetVisibleTargetFromSingleCell = _
            rng.SpecialCells(xlCellTypeVisible)

        On Error GoTo 0

    End If

End Function

'==========================================================
' Read Unicode text from Windows Clipboard
'==========================================================
Private Function private_GetClipboardUnicodeText() As String

#If VBA7 Then

    Dim hData As LongPtr
    Dim pData As LongPtr

#Else

    Dim hData As Long
    Dim pData As Long

#End If

    Dim length As Long
    Dim result As String

    If private_OpenClipboard(0) = 0 Then
        Exit Function
    End If

    On Error GoTo CleanExit

    hData = _
        private_GetClipboardData( _
            CF_UNICODETEXT)

    If hData <> 0 Then

        pData = private_GlobalLock(hData)

        If pData <> 0 Then

            length = private_lstrlenW(pData)

            If length > 0 Then

                result = _
                    String$(length, vbNullChar)

                private_CopyMemory _
                    StrPtr(result), _
                    pData, _
                    length * 2

            End If

            private_GlobalUnlock hData

        End If

    End If

CleanExit:

    private_CloseClipboard
    private_GetClipboardUnicodeText = result

End Function

'==========================================================
' Parse Excel clipboard first column
'
' Поддерживает:
' - строки;
' - табуляции;
' - значения в кавычках;
' - переносы строк внутри значения в кавычках.
'==========================================================
Private Function private_ParseExcelClipboardFirstColumn( _
    ByVal txt As String) As Variant

    Dim result() As String
    Dim current As String

    Dim i As Long
    Dim n As Long

    Dim ch As String
    Dim inQuotes As Boolean

    ' Удаляем завершающие переводы строк
    Do While Len(txt) > 0 _
        And ( _
            Right$(txt, 1) = vbCr _
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

            ' Игнорируем остальные столбцы текущей строки
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