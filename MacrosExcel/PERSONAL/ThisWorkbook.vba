Option Explicit

Private m_filterInputEvents As obj_FilterInputEvents


Private Sub Workbook_Open()
    private_InitializeFilterInputEvents
    BindKeys
End Sub

Private Sub Workbook_BeforeClose(Cancel As Boolean)
    Application.OnKey "^%r"
End Sub

Public Sub BindKeys()
    On Error Resume Next
    private_InitializeFilterInputEvents

    ' Application.OnKey applies to the entire Excel instance.
    ' Explicitly use the workbook that owns the global macros so that Excel
    ' does not resolve an identically named procedure in the active workbook.
    Application.OnKey "^q", private_GlobalMacroRef("ex_ShortcutsHandlers.fn_FilterContainsCurrentColumn")
    Application.OnKey "^r", private_GlobalMacroRef("ex_ShortcutsHandlers.fn_RecalculateActiveSheet")
    Application.OnKey "^%r", private_GlobalMacroRef("ex_Core.fn_ReloadActiveWorkbookVba")
    'Application.OnKey "^d", "PasteClipboardRowToVisibleCellsSkipTabs"

    Application.OnKey "%{PGUP}", private_GlobalMacroRef("ex_ShortcutsHandlers.fn_DatePlusOne")
    Application.OnKey "%{PGDN}", private_GlobalMacroRef("ex_ShortcutsHandlers.fn_DateMinusOne")

    ' EN: Ctrl + `
    Err.Clear
    Application.OnKey "^`", private_GlobalMacroRef("ex_ShortcutsHandlers.fn_ToggleFirstTwoRows")

    ' RU/UKR fallback: Ctrl + '
    If Err.Number <> 0 Then
        Err.Clear
        Application.OnKey "^'", private_GlobalMacroRef("ex_ShortcutsHandlers.fn_ToggleFirstTwoRows")
    End If

    ' RU fallback: Ctrl + ¸
    If Err.Number <> 0 Then
        Err.Clear
        Application.OnKey "^¸", private_GlobalMacroRef("ex_ShortcutsHandlers.fn_ToggleFirstTwoRows")
    End If

    On Error GoTo 0
End Sub

' --------------------------------------
' namespace Runtime {
' --------------------------------------
' Recreates runtime event handlers and global keyboard bindings after a code update.
Public Sub fn_ReloadRuntime()
    Set m_filterInputEvents = Nothing
    private_InitializeFilterInputEvents
    BindKeys
End Sub
' --------------------------------------
' } // namespace Runtime
' --------------------------------------


Private Sub private_InitializeFilterInputEvents()
    If m_filterInputEvents Is Nothing Then
        Set m_filterInputEvents = New obj_FilterInputEvents
        Set m_filterInputEvents.ExcelApplication = Application
    End If
End Sub

Private Function private_GlobalMacroRef(ByVal macroName As String) As String
    private_GlobalMacroRef = "'" & Replace$(ThisWorkbook.Name, "'", "''") & _
        "'!" & macroName
End Function