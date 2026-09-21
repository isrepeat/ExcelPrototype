Option Explicit

' Reads configuration rows from wsConfig/tbConfig by ASCII keys.
' Values may be Unicode because Excel stores worksheet text without an ANSI code page.
Private Const CONFIG_SHEET_NAME As String = "wsConfig"
Private Const CONFIG_TABLE_NAME As String = "tbConfig"
Private Const CONFIG_KIND_COLUMN_NAME As String = "Kind"
Private Const CONFIG_KEY_COLUMN_NAME As String = "Key"
Private Const CONFIG_VALUE_COLUMN_NAME As String = "Value"
Private Const IMPORT_KIND_NAME As String = "importCfg"

Private configValues As Object
Private configIsLoaded As Boolean

' --------------------------------------
' namespace API {
' --------------------------------------
' Resets the cache after a hot reload or a configuration-table change.
Public Sub fn_ResetCache()
    Set configValues = Nothing
    configIsLoaded = False
End Sub

Public Function fn_TryGetText(ByVal configKey As String, ByRef outValue As String) As Boolean
    outValue = VBA.vbNullString
    If Not private_TryEnsureLoaded() Then Exit Function
    If Not configValues.Exists(configKey) Then
        ex_Helpers.ex_ShowMessage "Configuration key was not found: " & configKey, VBA.vbExclamation, "Document Generation"
        Exit Function
    End If
    outValue = VBA.CStr(configValues(configKey))
    fn_TryGetText = True
End Function

Public Function fn_TryGetLong(ByVal configKey As String, ByRef outValue As Long) As Boolean
    Dim valueText As String
    outValue = 0
    If Not fn_TryGetText(configKey, valueText) Then Exit Function
    If Not VBA.IsNumeric(valueText) Then
        ex_Helpers.ex_ShowMessage "Configuration key must contain a number: " & configKey, VBA.vbExclamation, "Document Generation"
        Exit Function
    End If
    outValue = VBA.CLng(valueText)
    fn_TryGetLong = True
End Function
' --------------------------------------
' } // namespace API
' --------------------------------------

Private Function private_TryEnsureLoaded() As Boolean
    If configIsLoaded Then
        private_TryEnsureLoaded = True
        Exit Function
    End If
    ' The active workbook configuration takes precedence over imported values.
    Set configValues = VBA.CreateObject("Scripting.Dictionary")
    configValues.CompareMode = VBA.vbBinaryCompare
    If Not private_TryLoadWorkbookConfig(ThisWorkbook, True) Then Exit Function
    configIsLoaded = True
    private_TryEnsureLoaded = True
End Function

Private Function private_TryLoadWorkbookConfig(ByVal configWorkbook As Workbook, ByVal isLocalConfig As Boolean) As Boolean
    Dim configTable As ListObject, configRow As ListRow
    Dim kindText As String, keyText As String, valueText As String
    Dim dependencyWorkbook As Workbook, openedByLoader As Boolean

    On Error GoTo EH
    ' Sheet, table, and column names are intentionally ASCII and ACP-independent.
    Set configTable = configWorkbook.Worksheets(CONFIG_SHEET_NAME).ListObjects(CONFIG_TABLE_NAME)
    For Each configRow In configTable.ListRows
        kindText = VBA.Trim$(VBA.CStr(configRow.Range.Cells(1, configTable.ListColumns(CONFIG_KIND_COLUMN_NAME).Index).Value2))
        keyText = VBA.Trim$(VBA.CStr(configRow.Range.Cells(1, configTable.ListColumns(CONFIG_KEY_COLUMN_NAME).Index).Value2))
        valueText = VBA.CStr(configRow.Range.Cells(1, configTable.ListColumns(CONFIG_VALUE_COLUMN_NAME).Index).Value2)
        If VBA.Len(keyText) > 0 Then
            ' importCfg is allowed only in the local workbook. A shared config
            ' cannot implicitly load another dependency.
            If isLocalConfig And VBA.StrComp(kindText, IMPORT_KIND_NAME, VBA.vbTextCompare) = 0 Then
                If Not private_TryOpenWorkbook(private_ResolvePath(configWorkbook, valueText), dependencyWorkbook, openedByLoader) Then Exit Function
                If Not private_TryLoadWorkbookConfig(dependencyWorkbook, False) Then GoTo CleanFail
                If openedByLoader Then dependencyWorkbook.Close False
                Set dependencyWorkbook = Nothing
            ElseIf VBA.StrComp(kindText, IMPORT_KIND_NAME, VBA.vbTextCompare) <> 0 Then
                ' A shared config fills only missing keys. The local workbook
                ' can explicitly override any shared value.
                If isLocalConfig Or Not configValues.Exists(keyText) Then configValues(keyText) = valueText
            End If
        End If
    Next configRow
    private_TryLoadWorkbookConfig = True
    Exit Function
CleanFail:
    ' Close only workbooks opened by this loader. A user-opened workbook
    ' remains in its original state.
    On Error Resume Next
    If openedByLoader Then dependencyWorkbook.Close False
    On Error GoTo 0
    Exit Function
EH:
    ex_Helpers.ex_ShowMessage "Cannot read configuration workbook '" & configWorkbook.Name & "': " & Err.Description, VBA.vbExclamation, "Document Generation"
    Resume CleanFail
End Function

Private Function private_TryOpenWorkbook(ByVal workbookPath As String, ByRef outWorkbook As Workbook, ByRef outOpenedByLoader As Boolean) As Boolean
    Dim openWorkbook As Workbook
    outOpenedByLoader = False
    Set outWorkbook = Nothing
    If VBA.Len(VBA.Dir$(workbookPath)) = 0 Then
        ex_Helpers.ex_ShowMessage "Configuration dependency was not found: " & workbookPath, VBA.vbExclamation, "Document Generation"
        Exit Function
    End If
    ' Do not reopen a dependency that the user already has open.
    For Each openWorkbook In Application.Workbooks
        If VBA.StrComp(openWorkbook.FullName, workbookPath, VBA.vbTextCompare) = 0 Then
            Set outWorkbook = openWorkbook
            private_TryOpenWorkbook = True
            Exit Function
        End If
    Next openWorkbook
    ' The dependency is read-only. Do not update links or the recent-files list.
    Set outWorkbook = Application.Workbooks.Open(Filename:=workbookPath, ReadOnly:=True, UpdateLinks:=False, AddToMru:=False)
    outOpenedByLoader = True
    private_TryOpenWorkbook = True
End Function

Private Function private_ResolvePath(ByVal baseWorkbook As Workbook, ByVal configuredPath As String) As String
    configuredPath = VBA.Trim$(configuredPath)
    ' Relative paths are resolved from the workbook containing importCfg,
    ' which allows moving the complete project as one directory.
    If VBA.Len(configuredPath) >= 2 And VBA.Mid$(configuredPath, 2, 1) = ":" Then
        private_ResolvePath = configuredPath
    ElseIf VBA.Left$(configuredPath, 2) = "\\" Then
        private_ResolvePath = configuredPath
    Else
        private_ResolvePath = baseWorkbook.Path & Application.PathSeparator & configuredPath
    End If
End Function