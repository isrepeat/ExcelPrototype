Private Sub Workbook_Open()
    BindKeys
End Sub

Private Sub Workbook_BeforeClose(Cancel As Boolean)
    Application.OnKey "^%r"
End Sub

Public Sub BindKeys()
    On Error Resume Next

    ' Application.OnKey действует глобально для всего экземпляра Excel.
    ' Явно указываем книгу с глобальными макросами, чтобы Excel не выбрал
    ' одноимённую процедуру из активной книги.
    Application.OnKey "+%{UP}", private_GlobalMacroRef("fn_MoveTableRowsUp")
    Application.OnKey "+%{DOWN}", private_GlobalMacroRef("fn_MoveTableRowsDown")

    Application.OnKey "^q", private_GlobalMacroRef("FilterContainsCurrentColumn")
    Application.OnKey "^r", private_GlobalMacroRef("fn_RecalculateActiveSheet")
    Application.OnKey "^%r", private_GlobalMacroRef("fn_ReloadActiveWorkbookVba")
    'Application.OnKey "^d", "PasteClipboardRowToVisibleCellsSkipTabs"

    Application.OnKey "%{PGUP}", private_GlobalMacroRef("fn_DatePlusOne")
    Application.OnKey "%{PGDN}", private_GlobalMacroRef("fn_DateMinusOne")

    ' EN: Ctrl + `
    Err.Clear
    Application.OnKey "^`", private_GlobalMacroRef("fn_ToggleFirstTwoRows")

    ' RU/UKR fallback: Ctrl + '
    If Err.Number <> 0 Then
        Err.Clear
        Application.OnKey "^'", private_GlobalMacroRef("fn_ToggleFirstTwoRows")
    End If

    ' RU fallback: Ctrl + ¸
    If Err.Number <> 0 Then
        Err.Clear
        Application.OnKey "^¸", private_GlobalMacroRef("fn_ToggleFirstTwoRows")
    End If

    On Error GoTo 0
End Sub

Private Function private_GlobalMacroRef(ByVal macroName As String) As String
    private_GlobalMacroRef = "'" & Replace$(ThisWorkbook.Name, "'", "''") & _
        "'!" & macroName
End Function
