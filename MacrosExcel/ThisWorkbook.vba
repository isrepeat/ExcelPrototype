Private Sub Workbook_Open()
    BindKeys
    
    ' Application.OnKey "^v", "SmartPasteFromClipboard"

End Sub

Private Sub Workbook_BeforeClose(Cancel As Boolean)

    ' Application.OnKey "^v"

End Sub

Public Sub BindKeys()
    On Error Resume Next

    Application.OnKey "%{UP}", "fn_MoveTableRowsUp"
    Application.OnKey "%{DOWN}", "fn_MoveTableRowsDown"

    Application.OnKey "^q", "FilterContainsCurrentColumn"
    Application.OnKey "^e", "fn_RecalculateActiveSheet"
    Application.OnKey "^d", "PasteClipboardRowToVisibleCellsSkipTabs"

    Application.OnKey "%{PGUP}", "fn_DatePlusOne"
    Application.OnKey "%{PGDN}", "fn_DateMinusOne"

    ' EN: Ctrl + `
    Err.Clear
    Application.OnKey "^`", "fn_ToggleFirstTwoRows"

    ' RU/UKR fallback: Ctrl + '
    If Err.Number <> 0 Then
        Err.Clear
        Application.OnKey "^'", "fn_ToggleFirstTwoRows"
    End If

    ' RU fallback: Ctrl + ¸
    If Err.Number <> 0 Then
        Err.Clear
        Application.OnKey "^¸", "fn_ToggleFirstTwoRows"
    End If

    On Error GoTo 0
End Sub
