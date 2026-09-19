Option Explicit

Private configuredInputSheetName As String
Private configuredLookupCellAddresses As Collection
Private configuredLookupFieldTitles As Object
Private configuredCandidateStartCellAddress As String
Private configuredTableTitlePrefix As String
Private configuredHideCommandText As String
Private configuredMaxCandidateCount As Long
Private configuredQueryCallbackName As String
Private configuredSelectedValueIndex As Long
Private configuredColumns As Collection
Private configuredStyles As Object
Private configuredTableStyleName As String
Private configuredChromeStyleName As String
Private configuredCommandStyleName As String
Private configuredTitleStyleName As String
Private configuredSelectedStyleName As String
Private originalCandidateCellStyles As Object
Private activeLookupCellAddress As String
Private selectedCandidateRow As Range
Private isConfigured As Boolean
Private areSelectionRoutesRegistered As Boolean

' --------------------------------------
' namespace API {
' --------------------------------------
' Настраивает универсальный поиск кандидатов. SQL и предметная логика остаются
' в callback, имя которого передаёт владелец формы.
Public Function fn_Configure(ByVal candidatesConfig As Object) As Boolean
    If isConfigured Then
        If Not private_Candidates_TryRestoreSavedStyles() Then Exit Function
    End If
    If Not private_Candidates_TryReadConfig(candidatesConfig) Then Exit Function
    Set originalCandidateCellStyles = VBA.CreateObject("Scripting.Dictionary")
    activeLookupCellAddress = VBA.vbNullString
    Set selectedCandidateRow = Nothing
    areSelectionRoutesRegistered = False
    isConfigured = True
    ex_Helpers.LogDebug "Candidates configured | Sheet=" & configuredInputSheetName & _
        " | LookupCells=" & VBA.CStr(configuredLookupCellAddresses.Count) & _
        " | CandidateStartCell=" & configuredCandidateStartCellAddress & _
        " | QueryCallback=" & configuredQueryCallbackName
    fn_Configure = True
End Function

Public Function fn_RegisterRoutes() As Boolean
    Dim lookupCellAddress As Variant

    On Error GoTo EH
    If Not private_Candidates_TryEnsureConfigured() Then Exit Function
    For Each lookupCellAddress In configuredLookupCellAddresses
        If Not ex_CellChangeRouter.fn_RegisterChangeRoute( _
            configuredInputSheetName, VBA.CStr(lookupCellAddress), _
            "ex_Candidates.fn_OnInputChanged") Then Exit Function
    Next lookupCellAddress
    ex_Helpers.LogDebug "Candidate input routes registered | Sheet=" & _
        configuredInputSheetName & " | LookupCells=" & _
        VBA.CStr(configuredLookupCellAddresses.Count)
    fn_RegisterRoutes = True
    Exit Function
EH:
    ex_Helpers.LogError "Failed to register candidate routes | Number=" & _
        VBA.CStr(VBA.Err.Number) & " | Description=" & VBA.Err.Description
    VBA.MsgBox "Failed to register candidate routes: [" & _
        VBA.CStr(VBA.Err.Number) & "] " & VBA.Err.Description, _
        VBA.vbExclamation, "Document Generation"
End Function

Public Sub fn_OnInputChanged(ByVal changedSheet As Object, ByVal target As Range)
    Dim sourceSheet As Worksheet

    If changedSheet Is Nothing Or target Is Nothing Then Exit Sub
    If Not TypeOf changedSheet Is Worksheet Then Exit Sub
    If Not private_Candidates_TryEnsureConfigured() Then Exit Sub
    If target.CountLarge <> 1 Then
        VBA.MsgBox "Candidate search requires exactly one changed cell.", _
            VBA.vbExclamation, "Document Generation"
        Exit Sub
    End If
    Set sourceSheet = changedSheet
    If Not private_Candidates_TryValidateSourceSheet(sourceSheet) Then Exit Sub
    activeLookupCellAddress = target.Address(False, False)
    ex_Helpers.LogDebug "Candidate search started | Sheet=" & sourceSheet.Name & _
        " | Cell=" & activeLookupCellAddress
    private_Candidates_Refresh sourceSheet, target
End Sub

Public Sub fn_OnCandidateSelected(ByVal changedSheet As Object, ByVal target As Range)
    Dim sourceSheet As Worksheet

    If changedSheet Is Nothing Or target Is Nothing Then Exit Sub
    If Not TypeOf changedSheet Is Worksheet Or target.CountLarge <> 1 Then Exit Sub
    If Not private_Candidates_TryEnsureConfigured() Then Exit Sub
    Set sourceSheet = changedSheet
    If Not private_Candidates_TryValidateSourceSheet(sourceSheet) Then Exit Sub
    If VBA.Len(ex_Helpers.private_Text_Normalize(VBA.CStr(target.Value2))) = 0 Then Exit Sub
    If VBA.Len(activeLookupCellAddress) = 0 Then
        VBA.MsgBox "Select a lookup field before choosing a candidate.", _
            VBA.vbExclamation, "Document Generation"
        Exit Sub
    End If
    private_Candidates_Accept sourceSheet, target
End Sub

' Демонтирует таблицу кандидатов по клику на команду «Сховати».
Public Sub fn_OnClearCandidatesRequested(ByVal changedSheet As Object, ByVal target As Range)
    Dim sourceSheet As Worksheet
    Dim candidateRange As Range
    Dim screenUpdatingWasEnabled As Boolean
    Dim eventsWereEnabled As Boolean
    Dim hideStartedAt As Single

    If changedSheet Is Nothing Or target Is Nothing Then Exit Sub
    If Not TypeOf changedSheet Is Worksheet Then Exit Sub
    If Not private_Candidates_TryEnsureConfigured() Then Exit Sub
    Set sourceSheet = changedSheet
    If Not private_Candidates_TryValidateSourceSheet(sourceSheet) Then Exit Sub
    If Application.CutCopyMode <> False Then
        ex_Helpers.LogDebug "Candidate hide command skipped while Copy/Cut mode is active"
        Exit Sub
    End If
    On Error GoTo EH
    screenUpdatingWasEnabled = Application.ScreenUpdating
    eventsWereEnabled = Application.EnableEvents
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    hideStartedAt = VBA.Timer
    If Not private_Candidates_TryGetCandidateRange(sourceSheet, candidateRange) Then GoTo CleanExit
    If Not private_Candidates_TryClearTable(candidateRange) Then GoTo CleanExit
    private_Candidates_UnregisterSelectionRoutes
    activeLookupCellAddress = VBA.vbNullString
    ex_Helpers.LogDebug "Candidate table cleared by clear cell | Sheet=" & _
        sourceSheet.Name & " | Cell=" & target.Address(False, False) & _
        " | HideMs=" & VBA.CStr( _
        private_Candidates_GetElapsedMilliseconds(hideStartedAt))
CleanExit:
    Application.EnableEvents = eventsWereEnabled
    Application.ScreenUpdating = screenUpdatingWasEnabled
    Exit Sub
EH:
    Application.EnableEvents = eventsWereEnabled
    Application.ScreenUpdating = screenUpdatingWasEnabled
    ex_Helpers.LogError "Failed to hide candidate table | Number=" & _
        VBA.CStr(VBA.Err.Number) & " | Description=" & VBA.Err.Description
    VBA.MsgBox "Failed to hide candidate table: [" & _
        VBA.CStr(VBA.Err.Number) & "] " & VBA.Err.Description, _
        VBA.vbExclamation, "Document Generation"
End Sub
' --------------------------------------
' } // namespace API
' --------------------------------------

' --------------------------------------
' namespace Candidates {
' --------------------------------------
Private Function private_Candidates_TryReadConfig(ByVal candidatesConfig As Object) As Boolean
    Dim lookupCellAddresses As Collection

    isConfigured = False
    configuredInputSheetName = VBA.vbNullString
    Set configuredLookupCellAddresses = Nothing
    Set configuredLookupFieldTitles = Nothing
    configuredCandidateStartCellAddress = VBA.vbNullString
    configuredTableTitlePrefix = VBA.vbNullString
    configuredHideCommandText = VBA.vbNullString
    configuredMaxCandidateCount = 0
    configuredQueryCallbackName = VBA.vbNullString
    configuredSelectedValueIndex = 0
    Set configuredColumns = Nothing
    Set configuredStyles = Nothing
    configuredTableStyleName = VBA.vbNullString
    configuredChromeStyleName = VBA.vbNullString
    configuredCommandStyleName = VBA.vbNullString
    configuredTitleStyleName = VBA.vbNullString
    configuredSelectedStyleName = VBA.vbNullString
    If candidatesConfig Is Nothing Then
        VBA.MsgBox "Candidate configuration was not provided.", VBA.vbExclamation, _
            "Document Generation"
        Exit Function
    End If
    If Not private_Candidates_TryGetRequiredText(candidatesConfig, _
        "InputSheetName", configuredInputSheetName) Then Exit Function
    If Not private_Candidates_TryGetRequiredCollection(candidatesConfig, _
        "LookupCellAddresses", lookupCellAddresses) Then Exit Function
    If lookupCellAddresses.Count = 0 Then
        VBA.MsgBox "Candidate configuration has no lookup cell addresses.", _
            VBA.vbExclamation, "Document Generation"
        Exit Function
    End If
    If Not private_Candidates_TryGetRequiredObject( _
        candidatesConfig, "LookupFieldTitles", configuredLookupFieldTitles) Then Exit Function
    If Not private_Candidates_TryValidateLookupFieldTitles( _
        lookupCellAddresses) Then Exit Function
    If Not private_Candidates_TryGetRequiredText(candidatesConfig, _
        "CandidateStartCellAddress", configuredCandidateStartCellAddress) Then Exit Function
    If Not private_Candidates_TryGetRequiredText(candidatesConfig, _
        "TableTitlePrefix", configuredTableTitlePrefix) Then Exit Function
    If Not private_Candidates_TryGetRequiredText(candidatesConfig, _
        "HideCommandText", configuredHideCommandText) Then Exit Function
    If Not private_Candidates_TryGetRequiredLong(candidatesConfig, _
        "MaxCandidateCount", configuredMaxCandidateCount) Then Exit Function
    If configuredMaxCandidateCount <= 0 Then
        VBA.MsgBox "Candidate configuration MaxCandidateCount must be greater than zero.", _
            VBA.vbExclamation, "Document Generation"
        Exit Function
    End If
    If Not private_Candidates_TryGetRequiredText(candidatesConfig, _
        "QueryCallbackName", configuredQueryCallbackName) Then Exit Function
    If Not private_Candidates_TryGetRequiredLong(candidatesConfig, _
        "SelectedValueIndex", configuredSelectedValueIndex) Then Exit Function
    If configuredSelectedValueIndex < 0 Then
        VBA.MsgBox "Candidate selected value index cannot be negative.", _
            VBA.vbExclamation, "Document Generation"
        Exit Function
    End If
    If Not private_Candidates_TryGetRequiredCollection(candidatesConfig, _
        "Columns", configuredColumns) Then Exit Function
    If Not private_Candidates_TryValidateColumns(configuredColumns) Then Exit Function
    If Not private_Candidates_TryGetRequiredObject( _
        candidatesConfig, "Styles", configuredStyles) Then Exit Function
    If Not private_Candidates_TryGetRequiredText( _
        candidatesConfig, "TableStyleName", configuredTableStyleName) Then Exit Function
    If Not private_Candidates_TryGetRequiredText( _
        candidatesConfig, "ChromeStyleName", configuredChromeStyleName) Then Exit Function
    If Not private_Candidates_TryGetRequiredText( _
        candidatesConfig, "CommandStyleName", configuredCommandStyleName) Then Exit Function
    If Not private_Candidates_TryGetRequiredText( _
        candidatesConfig, "TitleStyleName", configuredTitleStyleName) Then Exit Function
    If Not private_Candidates_TryGetRequiredText( _
        candidatesConfig, "SelectedStyleName", configuredSelectedStyleName) Then Exit Function
    If Not private_Candidates_TryValidateConfiguredStyle( _
        configuredTableStyleName) Then Exit Function
    If Not private_Candidates_TryValidateConfiguredStyle( _
        configuredChromeStyleName) Then Exit Function
    If Not private_Candidates_TryValidateConfiguredStyle( _
        configuredCommandStyleName) Then Exit Function
    If Not private_Candidates_TryValidateConfiguredStyle( _
        configuredTitleStyleName) Then Exit Function
    If Not private_Candidates_TryValidateConfiguredStyle( _
        configuredSelectedStyleName) Then Exit Function
    Set configuredLookupCellAddresses = lookupCellAddresses
    private_Candidates_TryReadConfig = True
End Function

Private Function private_Candidates_TryGetRequiredObject(ByVal candidatesConfig As Object, _
    ByVal keyName As String, ByRef outObject As Object) As Boolean
    On Error GoTo EH
    Set outObject = Nothing
    If Not candidatesConfig.Exists(keyName) Then GoTo MissingValue
    Set outObject = candidatesConfig.Item(keyName)
    If outObject Is Nothing Then GoTo MissingValue
    private_Candidates_TryGetRequiredObject = True
    Exit Function
MissingValue:
    VBA.MsgBox "Candidate configuration object '" & keyName & "' is required.", _
        VBA.vbExclamation, "Document Generation"
    Exit Function
EH:
    VBA.MsgBox "Candidate configuration value '" & keyName & "' must be an object.", _
        VBA.vbExclamation, "Document Generation"
End Function

Private Function private_Candidates_TryValidateColumns( _
    ByVal columns As Collection _
) As Boolean
    Dim columnConfig As Object
    Dim sourceIndex As Long
    Dim numberFormat As String
    Dim headerText As String
    Dim sourceIndexes As Object

    If columns.Count = 0 Then
        VBA.MsgBox "Candidate configuration has no output columns.", _
            VBA.vbExclamation, "Document Generation"
        Exit Function
    End If
    Set sourceIndexes = VBA.CreateObject("Scripting.Dictionary")
    sourceIndexes.CompareMode = VBA.vbBinaryCompare
    For Each columnConfig In columns
        If Not private_Candidates_TryGetRequiredLong( _
            columnConfig, "SourceIndex", sourceIndex) Then Exit Function
        If sourceIndex < 0 Then
            VBA.MsgBox "Candidate column SourceIndex cannot be negative.", _
                VBA.vbExclamation, "Document Generation"
            Exit Function
        End If
        If Not private_Candidates_TryGetRequiredText( _
            columnConfig, "NumberFormat", numberFormat) Then Exit Function
        If Not private_Candidates_TryGetRequiredText( _
            columnConfig, "Header", headerText) Then Exit Function
        If sourceIndexes.Exists(VBA.CStr(sourceIndex)) Then
            VBA.MsgBox "Candidate column SourceIndex is duplicated: " & _
                VBA.CStr(sourceIndex) & ".", VBA.vbExclamation, _
                "Document Generation"
            Exit Function
        End If
        sourceIndexes.Add VBA.CStr(sourceIndex), True
    Next columnConfig
    If Not sourceIndexes.Exists(VBA.CStr(configuredSelectedValueIndex)) Then
        VBA.MsgBox "SelectedValueIndex is not included in candidate Columns.", _
            VBA.vbExclamation, "Document Generation"
        Exit Function
    End If
    private_Candidates_TryValidateColumns = True
End Function

Private Function private_Candidates_TryValidateLookupFieldTitles( _
    ByVal lookupCellAddresses As Collection _
) As Boolean
    Dim lookupCellAddress As Variant
    Dim lookupFieldTitle As String

    On Error GoTo EH
    For Each lookupCellAddress In lookupCellAddresses
        If Not configuredLookupFieldTitles.Exists(VBA.CStr(lookupCellAddress)) Then
            VBA.MsgBox "Candidate title is not configured for lookup cell '" & _
                VBA.CStr(lookupCellAddress) & "'.", VBA.vbExclamation, _
                "Document Generation"
            Exit Function
        End If
        lookupFieldTitle = VBA.Trim$(VBA.CStr(configuredLookupFieldTitles.Item( _
            VBA.CStr(lookupCellAddress))))
        If VBA.Len(lookupFieldTitle) = 0 Then
            VBA.MsgBox "Candidate title is empty for lookup cell '" & _
                VBA.CStr(lookupCellAddress) & "'.", VBA.vbExclamation, _
                "Document Generation"
            Exit Function
        End If
    Next lookupCellAddress
    private_Candidates_TryValidateLookupFieldTitles = True
    Exit Function
EH:
    VBA.MsgBox "Candidate LookupFieldTitles must be a dictionary of text values.", _
        VBA.vbExclamation, "Document Generation"
End Function

Private Function private_Candidates_TryGetRequiredText(ByVal candidatesConfig As Object, _
    ByVal keyName As String, ByRef outText As String) As Boolean
    On Error GoTo EH
    outText = VBA.vbNullString
    If Not candidatesConfig.Exists(keyName) Then GoTo MissingValue
    outText = VBA.Trim$(VBA.CStr(candidatesConfig.Item(keyName)))
    If VBA.Len(outText) = 0 Then GoTo MissingValue
    private_Candidates_TryGetRequiredText = True
    Exit Function
MissingValue:
    VBA.MsgBox "Candidate configuration value '" & keyName & "' is required.", _
        VBA.vbExclamation, "Document Generation"
    Exit Function
EH:
    VBA.MsgBox "Candidate configuration does not provide key '" & keyName & "'.", _
        VBA.vbExclamation, "Document Generation"
End Function

Private Function private_Candidates_TryGetRequiredCollection(ByVal candidatesConfig As Object, _
    ByVal keyName As String, ByRef outCollection As Collection) As Boolean
    On Error GoTo EH
    Set outCollection = Nothing
    If Not candidatesConfig.Exists(keyName) Then GoTo MissingValue
    Set outCollection = candidatesConfig.Item(keyName)
    If outCollection Is Nothing Then GoTo MissingValue
    private_Candidates_TryGetRequiredCollection = True
    Exit Function
MissingValue:
    VBA.MsgBox "Candidate configuration collection '" & keyName & "' is required.", _
        VBA.vbExclamation, "Document Generation"
    Exit Function
EH:
    VBA.MsgBox "Candidate configuration value '" & keyName & "' must be a Collection.", _
        VBA.vbExclamation, "Document Generation"
End Function

Private Function private_Candidates_TryGetRequiredLong(ByVal candidatesConfig As Object, _
    ByVal keyName As String, ByRef outValue As Long) As Boolean
    On Error GoTo EH
    outValue = 0
    If Not candidatesConfig.Exists(keyName) Then GoTo MissingValue
    If Not VBA.IsNumeric(candidatesConfig.Item(keyName)) Then GoTo MissingValue
    outValue = VBA.CLng(candidatesConfig.Item(keyName))
    private_Candidates_TryGetRequiredLong = True
    Exit Function
MissingValue:
    VBA.MsgBox "Candidate configuration numeric value '" & keyName & "' is required.", _
        VBA.vbExclamation, "Document Generation"
    Exit Function
EH:
    VBA.MsgBox "Candidate configuration value '" & keyName & "' is invalid.", _
        VBA.vbExclamation, "Document Generation"
End Function

Private Function private_Candidates_TryGetRequiredBoolean(ByVal candidatesConfig As Object, _
    ByVal keyName As String, ByRef outValue As Boolean) As Boolean
    On Error GoTo EH
    outValue = False
    If Not candidatesConfig.Exists(keyName) Then GoTo MissingValue
    If VBA.VarType(candidatesConfig.Item(keyName)) <> VBA.vbBoolean Then GoTo MissingValue
    outValue = VBA.CBool(candidatesConfig.Item(keyName))
    private_Candidates_TryGetRequiredBoolean = True
    Exit Function
MissingValue:
    VBA.MsgBox "Candidate configuration Boolean value '" & keyName & "' is required.", _
        VBA.vbExclamation, "Document Generation"
    Exit Function
EH:
    VBA.MsgBox "Candidate configuration value '" & keyName & "' is invalid.", _
        VBA.vbExclamation, "Document Generation"
End Function

Private Function private_Candidates_TryGetRequiredDouble(ByVal candidatesConfig As Object, _
    ByVal keyName As String, ByRef outValue As Double) As Boolean
    On Error GoTo EH
    outValue = 0
    If Not candidatesConfig.Exists(keyName) Then GoTo MissingValue
    If Not VBA.IsNumeric(candidatesConfig.Item(keyName)) Then GoTo MissingValue
    outValue = VBA.CDbl(candidatesConfig.Item(keyName))
    private_Candidates_TryGetRequiredDouble = True
    Exit Function
MissingValue:
    VBA.MsgBox "Candidate configuration numeric value '" & keyName & "' is required.", _
        VBA.vbExclamation, "Document Generation"
    Exit Function
EH:
    VBA.MsgBox "Candidate configuration value '" & keyName & "' is invalid.", _
        VBA.vbExclamation, "Document Generation"
End Function

' Единый контракт стиля ячейки. Роли таблицы задаются только ссылками на
' именованные стили в конфигурации, а не специальными свойствами модуля.
Private Function private_Candidates_TryValidateConfiguredStyle( _
    ByVal styleName As String _
) As Boolean
    Dim cellStyle As Object
    Dim fontSize As Double

    If Not private_Candidates_TryGetConfiguredStyle(styleName, cellStyle) Then Exit Function
    If cellStyle.Count = 0 Then
        VBA.MsgBox "Candidate style '" & styleName & "' must define at least one property.", _
            VBA.vbExclamation, "Document Generation"
        Exit Function
    End If
    If cellStyle.Exists("FontSize") Then
        If Not private_Candidates_TryGetRequiredDouble(cellStyle, "FontSize", fontSize) Then Exit Function
        If fontSize <= 0 Then
            VBA.MsgBox "Candidate style '" & styleName & "' FontSize must be greater than zero.", _
                VBA.vbExclamation, "Document Generation"
            Exit Function
        End If
    End If
    private_Candidates_TryValidateConfiguredStyle = True
End Function

Private Function private_Candidates_TryGetConfiguredStyle( _
    ByVal styleName As String, _
    ByRef outCellStyle As Object _
) As Boolean
    On Error GoTo EH
    Set outCellStyle = Nothing
    If configuredStyles Is Nothing Then GoTo MissingStyle
    If Not configuredStyles.Exists(styleName) Then GoTo MissingStyle
    Set outCellStyle = configuredStyles.Item(styleName)
    If outCellStyle Is Nothing Then GoTo MissingStyle
    private_Candidates_TryGetConfiguredStyle = True
    Exit Function
MissingStyle:
    VBA.MsgBox "Candidate cell style '" & styleName & "' is required.", _
        VBA.vbExclamation, "Document Generation"
    Exit Function
EH:
    VBA.MsgBox "Candidate cell style '" & styleName & "' must be an object.", _
        VBA.vbExclamation, "Document Generation"
End Function

Private Function private_Candidates_TryApplyCellStyle( _
    ByVal targetRange As Range, _
    ByVal styleName As String _
) As Boolean
    Dim cellStyle As Object

    On Error GoTo EH
    If Not private_Candidates_TryGetConfiguredStyle(styleName, cellStyle) Then Exit Function
    If cellStyle.Exists("FontName") Then targetRange.Font.Name = _
        VBA.CStr(cellStyle.Item("FontName"))
    If cellStyle.Exists("FontSize") Then targetRange.Font.Size = _
        VBA.CDbl(cellStyle.Item("FontSize"))
    If cellStyle.Exists("FontColor") Then targetRange.Font.Color = _
        VBA.CLng(cellStyle.Item("FontColor"))
    If cellStyle.Exists("FillColor") Then targetRange.Interior.Color = _
        VBA.CLng(cellStyle.Item("FillColor"))
    If cellStyle.Exists("HorizontalAlignment") Then targetRange.HorizontalAlignment = _
        VBA.CLng(cellStyle.Item("HorizontalAlignment"))
    If cellStyle.Exists("VerticalAlignment") Then targetRange.VerticalAlignment = _
        VBA.CLng(cellStyle.Item("VerticalAlignment"))
    If cellStyle.Exists("WrapText") Then targetRange.WrapText = _
        VBA.CBool(cellStyle.Item("WrapText"))
    private_Candidates_TryApplyCellStyle = True
    Exit Function
EH:
    ex_Helpers.LogError "Failed to apply candidate cell style | Style=" & _
        styleName & " | Number=" & VBA.CStr(VBA.Err.Number) & _
        " | Description=" & VBA.Err.Description
    VBA.MsgBox "Failed to apply candidate cell style '" & styleName & "': [" & _
        VBA.CStr(VBA.Err.Number) & "] " & VBA.Err.Description, _
        VBA.vbExclamation, "Document Generation"
End Function

Private Function private_Candidates_TryEnsureConfigured() As Boolean
    If isConfigured Then
        private_Candidates_TryEnsureConfigured = True
        Exit Function
    End If
    VBA.MsgBox "Candidate module is not configured.", VBA.vbExclamation, _
        "Document Generation"
End Function

Private Function private_Candidates_TryValidateSourceSheet(ByVal sourceSheet As Worksheet) As Boolean
    If VBA.StrComp(sourceSheet.Name, configuredInputSheetName, VBA.vbTextCompare) = 0 Then
        private_Candidates_TryValidateSourceSheet = True
        Exit Function
    End If
    VBA.MsgBox "Candidate callback was received from unexpected sheet '" & _
        sourceSheet.Name & "'.", VBA.vbExclamation, "Document Generation"
End Function

Private Function private_Candidates_TryRegisterSelectionRoutes( _
    ByVal sourceSheet As Worksheet, _
    ByVal candidateRange As Range _
) As Boolean
    Dim clearCandidatesCell As Range

    If areSelectionRoutesRegistered Then
        private_Candidates_TryRegisterSelectionRoutes = True
        Exit Function
    End If
    If candidateRange.Row <= 2 Then
        VBA.MsgBox "CandidateStartCellAddress must leave rows above for header and clear command.", _
            VBA.vbExclamation, "Document Generation"
        Exit Function
    End If
    Set clearCandidatesCell = candidateRange.Cells(1, 1).Offset(-2, 0)
    If Not ex_CellChangeRouter.fn_RegisterSelectionRoute( _
        sourceSheet.Name, candidateRange.Address(False, False), _
        "ex_Candidates.fn_OnCandidateSelected") Then Exit Function
    If Not ex_CellChangeRouter.fn_RegisterSelectionRoute( _
        sourceSheet.Name, clearCandidatesCell.Address(False, False), _
        "ex_Candidates.fn_OnClearCandidatesRequested") Then Exit Function
    areSelectionRoutesRegistered = True
    ex_Helpers.LogDebug "Candidate selection routes registered | Sheet=" & _
        sourceSheet.Name & " | CandidateRange=" & _
        candidateRange.Address(False, False) & " | ClearCell=" & _
        clearCandidatesCell.Address(False, False)
    private_Candidates_TryRegisterSelectionRoutes = True
End Function

Private Sub private_Candidates_UnregisterSelectionRoutes()
    If Not areSelectionRoutesRegistered Then Exit Sub
    ex_CellChangeRouter.fn_UnregisterSelectionRoutes _
        configuredInputSheetName, "ex_Candidates.fn_OnCandidateSelected"
    ex_CellChangeRouter.fn_UnregisterSelectionRoutes _
        configuredInputSheetName, "ex_Candidates.fn_OnClearCandidatesRequested"
    areSelectionRoutesRegistered = False
    ex_Helpers.LogDebug "Candidate selection routes unregistered | Sheet=" & _
        configuredInputSheetName
End Sub

Private Function private_Candidates_TryClearTable( _
    ByVal candidateRange As Range _
) As Boolean
    Dim candidateTableRange As Range

    On Error GoTo EH
    If Not private_Candidates_TryClear(candidateRange) Then Exit Function
    Set candidateTableRange = candidateRange.Rows(1).Offset(-3, 0).Resize( _
        candidateRange.Rows.Count + 3, candidateRange.Columns.Count)
    candidateTableRange.Rows(1).Resize(3, candidateRange.Columns.Count).ClearContents
    private_Candidates_TryClearTable = True
    Exit Function
EH:
    ex_Helpers.LogError "Failed to clear candidate table | Number=" & _
        VBA.CStr(VBA.Err.Number) & " | Description=" & VBA.Err.Description
    VBA.MsgBox "Failed to clear candidate table: [" & _
        VBA.CStr(VBA.Err.Number) & "] " & VBA.Err.Description, _
        VBA.vbExclamation, "Document Generation"
End Function

Private Function private_Candidates_TryApplyHeaders() As Boolean
    Dim sourceSheet As Worksheet
    Dim candidateRange As Range
    Dim headerRange As Range
    Dim commandCell As Range
    Dim titleCell As Range
    Dim headerCell As Range
    Dim columnConfig As Object
    Dim lookupFieldTitle As String
    Dim outputColumnIndex As Long

    On Error GoTo EH
    Set sourceSheet = ThisWorkbook.Worksheets(configuredInputSheetName)
    If Not private_Candidates_TryGetCandidateRange(sourceSheet, candidateRange) Then Exit Function
    If candidateRange.Row <= 2 Then
        VBA.MsgBox "CandidateStartCellAddress must leave rows above for title, headers and command.", _
            VBA.vbExclamation, "Document Generation"
        Exit Function
    End If
    Set headerRange = candidateRange.Rows(1).Offset(-1, 0)
    Set commandCell = headerRange.Cells(1, 1).Offset(-1, 0)
    Set titleCell = commandCell.Offset(-1, 0)
    If Not configuredLookupFieldTitles.Exists(activeLookupCellAddress) Then
        VBA.MsgBox "Candidate title is not configured for active lookup cell '" & _
            activeLookupCellAddress & "'.", VBA.vbExclamation, "Document Generation"
        Exit Function
    End If
    lookupFieldTitle = VBA.CStr(configuredLookupFieldTitles.Item( _
        activeLookupCellAddress))
    If VBA.Len(VBA.Trim$(lookupFieldTitle)) = 0 Then
        VBA.MsgBox "Candidate title is empty for active lookup cell '" & _
            activeLookupCellAddress & "'.", VBA.vbExclamation, "Document Generation"
        Exit Function
    End If
    If Not private_Candidates_TrySaveCellStyle(titleCell) Then Exit Function
    If Not private_Candidates_TrySaveCellStyle(commandCell) Then Exit Function
    For Each headerCell In headerRange.Cells
        If Not private_Candidates_TrySaveCellStyle(headerCell) Then Exit Function
    Next headerCell
    If Not private_Candidates_TryApplyCellStyle( _
        titleCell, configuredTitleStyleName) Then Exit Function
    titleCell.Value = configuredTableTitlePrefix & " """ & lookupFieldTitle & """:"
    If Not private_Candidates_TryApplyCellStyle( _
        commandCell, configuredCommandStyleName) Then Exit Function
    commandCell.Value = configuredHideCommandText
    If Not private_Candidates_TryApplyCellStyle( _
        headerRange, configuredChromeStyleName) Then Exit Function
    outputColumnIndex = 0
    For Each columnConfig In configuredColumns
        outputColumnIndex = outputColumnIndex + 1
        headerRange.Cells(1, outputColumnIndex).Value = _
            VBA.CStr(columnConfig.Item("Header"))
    Next columnConfig
    ex_Helpers.LogDebug "Candidate headers applied | Sheet=" & sourceSheet.Name & _
        " | Range=" & headerRange.Address(False, False)
    private_Candidates_TryApplyHeaders = True
    Exit Function
EH:
    ex_Helpers.LogError "Failed to apply candidate headers | Number=" & _
        VBA.CStr(VBA.Err.Number) & " | Description=" & VBA.Err.Description
    VBA.MsgBox "Failed to apply candidate headers: [" & _
        VBA.CStr(VBA.Err.Number) & "] " & VBA.Err.Description, _
        VBA.vbExclamation, "Document Generation"
End Function

Private Function private_Candidates_TryGetCandidateRange( _
    ByVal sourceSheet As Worksheet, _
    ByRef outCandidateRange As Range _
) As Boolean
    Dim candidateStartCell As Range

    On Error GoTo EH
    Set outCandidateRange = Nothing
    Set candidateStartCell = sourceSheet.Range(configuredCandidateStartCellAddress)
    If candidateStartCell.Cells.CountLarge <> 1 Then
        VBA.MsgBox "CandidateStartCellAddress must reference exactly one cell.", _
            VBA.vbExclamation, "Document Generation"
        Exit Function
    End If
    Set outCandidateRange = candidateStartCell.Offset(2, 0).Resize( _
        configuredMaxCandidateCount, configuredColumns.Count)
    private_Candidates_TryGetCandidateRange = True
    Exit Function
EH:
    VBA.MsgBox "Failed to resolve candidate range from start cell '" & _
        configuredCandidateStartCellAddress & "': " & VBA.Err.Description, _
        VBA.vbExclamation, "Document Generation"
End Function

Private Sub private_Candidates_Refresh(ByVal sourceSheet As Worksheet, ByVal lookupCell As Range)
    Dim inputText As String
    Dim candidates As Collection
    Dim candidateRange As Range
    Dim queryStartedAt As Single
    Dim renderStartedAt As Single
    Dim screenUpdatingWasEnabled As Boolean

    On Error GoTo EH
    screenUpdatingWasEnabled = Application.ScreenUpdating
    Application.ScreenUpdating = False
    If Not private_Candidates_TryGetCandidateRange(sourceSheet, candidateRange) Then GoTo CleanExit
    If Not private_Candidates_TryClearTable(candidateRange) Then GoTo CleanExit
    private_Candidates_UnregisterSelectionRoutes
    inputText = ex_Helpers.private_Text_Normalize(VBA.CStr(lookupCell.Value2))
    If VBA.Len(inputText) = 0 Then
        ex_Helpers.LogDebug "Candidate search cleared | Sheet=" & sourceSheet.Name & _
            " | Cell=" & activeLookupCellAddress
        activeLookupCellAddress = VBA.vbNullString
        GoTo CleanExit
    End If
    queryStartedAt = VBA.Timer
    If Not private_Candidates_TryFindCandidates(inputText, candidates) Then GoTo CleanExit
    If candidates.Count > candidateRange.Rows.Count Then
        VBA.MsgBox "Candidate query returned more rows than configured maximum.", _
            VBA.vbExclamation, "Document Generation"
        GoTo CleanExit
    End If
    renderStartedAt = VBA.Timer
    ex_Helpers.LogDebug "Candidate search completed | Sheet=" & sourceSheet.Name & _
        " | Cell=" & activeLookupCellAddress & " | Count=" & _
        VBA.CStr(candidates.Count) & " | QueryMs=" & _
        VBA.CStr(private_Candidates_GetElapsedMilliseconds(queryStartedAt))
    If candidates.Count = 0 Then GoTo CleanExit
    If Not private_Candidates_TryApplyHeaders() Then GoTo CleanExit
    If Not private_Candidates_TryRenderCandidates(candidateRange, candidates) Then GoTo CleanExit
    If Not private_Candidates_TryRegisterSelectionRoutes( _
        sourceSheet, candidateRange) Then GoTo CleanExit
    ex_Helpers.LogDebug "Candidate table rendered | Sheet=" & sourceSheet.Name & _
        " | Count=" & VBA.CStr(candidates.Count) & " | RenderMs=" & _
        VBA.CStr(private_Candidates_GetElapsedMilliseconds(renderStartedAt))
CleanExit:
    Application.ScreenUpdating = screenUpdatingWasEnabled
    Exit Sub
EH:
    Application.ScreenUpdating = screenUpdatingWasEnabled
    ex_Helpers.LogError "Failed to refresh candidate table | Number=" & _
        VBA.CStr(VBA.Err.Number) & " | Description=" & VBA.Err.Description
    VBA.MsgBox "Failed to refresh candidate table: [" & _
        VBA.CStr(VBA.Err.Number) & "] " & VBA.Err.Description, _
        VBA.vbExclamation, "Document Generation"
End Sub

Private Function private_Candidates_GetElapsedMilliseconds( _
    ByVal startedAt As Single _
) As Long
    Dim elapsedSeconds As Single

    elapsedSeconds = VBA.Timer - startedAt
    If elapsedSeconds < 0 Then elapsedSeconds = elapsedSeconds + 86400!
    private_Candidates_GetElapsedMilliseconds = VBA.CLng(elapsedSeconds * 1000!)
End Function

Private Function private_Candidates_TryFindCandidates(ByVal searchText As String, _
    ByRef outCandidates As Collection) As Boolean
    Dim macroReference As String

    On Error GoTo EH
    Set outCandidates = Nothing
    macroReference = "'" & VBA.Replace$(ThisWorkbook.Name, "'", "''") & _
        "'!" & configuredQueryCallbackName
    Set outCandidates = Application.Run(macroReference, searchText, configuredMaxCandidateCount)
    If outCandidates Is Nothing Then
        VBA.MsgBox "Candidate query callback returned no result: " & _
            configuredQueryCallbackName, VBA.vbExclamation, "Document Generation"
        Exit Function
    End If
    private_Candidates_TryFindCandidates = True
    Exit Function
EH:
    ex_Helpers.LogError "Candidate query callback failed | Callback=" & _
        configuredQueryCallbackName & " | Number=" & VBA.CStr(VBA.Err.Number) & _
        " | Description=" & VBA.Err.Description
    VBA.MsgBox "Candidate query callback failed: [" & VBA.CStr(VBA.Err.Number) & _
        "] " & VBA.Err.Description, VBA.vbExclamation, "Document Generation"
End Function

Private Function private_Candidates_TryRenderCandidates( _
    ByVal candidateRange As Range, _
    ByVal candidates As Collection _
) As Boolean
    Dim columnConfig As Object
    Dim candidateItem As Variant
    Dim candidateCell As Range
    Dim sourceIndex As Long
    Dim numberFormat As String
    Dim outputColumnIndex As Long
    Dim outputRowIndex As Long
    Dim outputRange As Range
    Dim outputValues() As Variant

    If candidates.Count = 0 Then
        private_Candidates_TryRenderCandidates = True
        Exit Function
    End If
    On Error GoTo InvalidCandidate
    Set outputRange = candidateRange.Resize( _
        candidates.Count, configuredColumns.Count)
    For Each candidateCell In outputRange.Cells
        If Not private_Candidates_TrySaveCellStyle(candidateCell) Then Exit Function
    Next candidateCell
    If Not private_Candidates_TryApplyCellStyle( _
        outputRange, configuredTableStyleName) Then Exit Function
    outputColumnIndex = 0
    For Each columnConfig In configuredColumns
        outputColumnIndex = outputColumnIndex + 1
        numberFormat = VBA.CStr(columnConfig.Item("NumberFormat"))
        outputRange.Columns(outputColumnIndex).NumberFormat = numberFormat
    Next columnConfig
    ReDim outputValues(1 To candidates.Count, 1 To configuredColumns.Count)
    outputRowIndex = 0
    For Each candidateItem In candidates
        If Not VBA.IsArray(candidateItem) Then GoTo InvalidCandidate
        outputRowIndex = outputRowIndex + 1
        outputColumnIndex = 0
        For Each columnConfig In configuredColumns
            outputColumnIndex = outputColumnIndex + 1
            sourceIndex = VBA.CLng(columnConfig.Item("SourceIndex"))
            outputValues(outputRowIndex, outputColumnIndex) = _
                VBA.CStr(candidateItem(sourceIndex))
        Next columnConfig
    Next candidateItem
    outputRange.Value2 = outputValues
    private_Candidates_TryRenderCandidates = True
    Exit Function
InvalidCandidate:
    VBA.MsgBox "Candidate has invalid data format for configured value indexes.", _
        VBA.vbExclamation, "Document Generation"
End Function

Private Sub private_Candidates_Accept(ByVal sourceSheet As Worksheet, ByVal candidateCell As Range)
    Dim selectedValue As String
    Dim eventsWereEnabled As Boolean
    Dim candidateRange As Range
    Dim selectedOutputColumnIndex As Long

    If Not private_Candidates_TryGetCandidateRange(sourceSheet, candidateRange) Then Exit Sub
    If Not private_Candidates_TryGetOutputColumnIndex( _
        configuredSelectedValueIndex, selectedOutputColumnIndex) Then Exit Sub
    selectedValue = ex_Helpers.private_Text_Normalize(VBA.CStr(sourceSheet.Cells( _
        candidateCell.Row, candidateRange.Columns( _
        selectedOutputColumnIndex).Column).Value2))
    If VBA.Len(selectedValue) = 0 Then
        VBA.MsgBox "Candidate selected value is empty in " & candidateCell.Address(False, False) & ".", _
            VBA.vbExclamation, "Document Generation"
        Exit Sub
    End If
    eventsWereEnabled = Application.EnableEvents
    On Error GoTo EH
    Application.EnableEvents = False
    sourceSheet.Range(activeLookupCellAddress).Value = selectedValue
    If Not private_Candidates_TryApplySelectedCellStyle(candidateCell) Then GoTo CleanExit
    ex_Helpers.LogDebug "Candidate accepted | Sheet=" & sourceSheet.Name & _
        " | TargetCell=" & activeLookupCellAddress & " | CandidateRow=" & _
        VBA.CStr(candidateCell.Row)
    Application.EnableEvents = eventsWereEnabled
    Exit Sub
CleanExit:
    Application.EnableEvents = eventsWereEnabled
    Exit Sub
EH:
    Application.EnableEvents = eventsWereEnabled
    ex_Helpers.LogError "Failed to accept candidate | Number=" & _
        VBA.CStr(VBA.Err.Number) & " | Description=" & VBA.Err.Description
    ex_Helpers.ex_ShowErrorMessage "Failed to accept candidate: " & _
        VBA.Err.Description, VBA.vbExclamation, "Document Generation"
End Sub

Private Function private_Candidates_TryApplySelectedCellStyle( _
    ByVal candidateCell As Range _
) As Boolean
    Dim candidateRow As Range

    Set candidateRow = candidateCell.Worksheet.Cells( _
        candidateCell.Row, candidateCell.Worksheet.Range( _
        configuredCandidateStartCellAddress).Column).Resize( _
        1, configuredColumns.Count)
    If Not selectedCandidateRow Is Nothing Then
        If Not private_Candidates_TryApplyCellStyle( _
            selectedCandidateRow, configuredTableStyleName) Then Exit Function
    End If
    If Not private_Candidates_TryApplyCellStyle( _
        candidateRow, configuredSelectedStyleName) Then Exit Function
    Set selectedCandidateRow = candidateRow
    private_Candidates_TryApplySelectedCellStyle = True
End Function

Private Function private_Candidates_TryGetOutputColumnIndex( _
    ByVal sourceIndex As Long, _
    ByRef outOutputColumnIndex As Long _
) As Boolean
    Dim columnConfig As Object

    outOutputColumnIndex = 0
    For Each columnConfig In configuredColumns
        outOutputColumnIndex = outOutputColumnIndex + 1
        If VBA.CLng(columnConfig.Item("SourceIndex")) = sourceIndex Then
            private_Candidates_TryGetOutputColumnIndex = True
            Exit Function
        End If
    Next columnConfig
    VBA.MsgBox "Candidate selected source index is not mapped to an output column.", _
        VBA.vbExclamation, "Document Generation"
End Function

Private Function private_Candidates_TrySaveCellStyle( _
    ByVal candidateCell As Range _
) As Boolean
    Dim cellKey As String
    Dim cellStyle As Object

    On Error GoTo EH
    If originalCandidateCellStyles Is Nothing Then
        VBA.MsgBox "Candidate cell style storage is not initialized.", _
            VBA.vbExclamation, "Document Generation"
        Exit Function
    End If
    cellKey = candidateCell.Worksheet.CodeName & "!" & _
        candidateCell.Address(False, False)
    If originalCandidateCellStyles.Exists(cellKey) Then
        private_Candidates_TrySaveCellStyle = True
        Exit Function
    End If
    Set cellStyle = VBA.CreateObject("Scripting.Dictionary")
    cellStyle.CompareMode = VBA.vbBinaryCompare
    cellStyle.Add "FontName", candidateCell.Font.Name
    cellStyle.Add "FontSize", candidateCell.Font.Size
    cellStyle.Add "FontColor", candidateCell.Font.Color
    cellStyle.Add "FillColor", candidateCell.Interior.Color
    cellStyle.Add "HorizontalAlignment", candidateCell.HorizontalAlignment
    cellStyle.Add "VerticalAlignment", candidateCell.VerticalAlignment
    cellStyle.Add "WrapText", candidateCell.WrapText
    cellStyle.Add "NumberFormat", candidateCell.NumberFormat
    cellStyle.Add "Cell", candidateCell
    originalCandidateCellStyles.Add cellKey, cellStyle
    private_Candidates_TrySaveCellStyle = True
    Exit Function
EH:
    ex_Helpers.LogError "Failed to save candidate cell style | Number=" & _
        VBA.CStr(VBA.Err.Number) & " | Description=" & VBA.Err.Description
    VBA.MsgBox "Failed to save candidate cell style: [" & _
        VBA.CStr(VBA.Err.Number) & "] " & VBA.Err.Description, _
        VBA.vbExclamation, "Document Generation"
End Function

Private Function private_Candidates_TryRestoreSavedStyles() As Boolean
    Dim cellKey As Variant
    Dim cellStyle As Object
    Dim candidateCell As Range

    On Error GoTo EH
    If originalCandidateCellStyles Is Nothing Then
        private_Candidates_TryRestoreSavedStyles = True
        Exit Function
    End If
    For Each cellKey In originalCandidateCellStyles.Keys
        Set cellStyle = originalCandidateCellStyles.Item(cellKey)
        Set candidateCell = cellStyle.Item("Cell")
        candidateCell.Font.Name = VBA.CStr(cellStyle.Item("FontName"))
        candidateCell.Font.Size = VBA.CDbl(cellStyle.Item("FontSize"))
        candidateCell.Font.Color = VBA.CLng(cellStyle.Item("FontColor"))
        candidateCell.Interior.Color = VBA.CLng(cellStyle.Item("FillColor"))
        candidateCell.HorizontalAlignment = VBA.CLng( _
            cellStyle.Item("HorizontalAlignment"))
        candidateCell.VerticalAlignment = VBA.CLng( _
            cellStyle.Item("VerticalAlignment"))
        candidateCell.WrapText = VBA.CBool(cellStyle.Item("WrapText"))
        candidateCell.NumberFormat = VBA.CStr(cellStyle.Item("NumberFormat"))
    Next cellKey
    originalCandidateCellStyles.RemoveAll
    private_Candidates_TryRestoreSavedStyles = True
    Exit Function
EH:
    ex_Helpers.LogError "Failed to restore candidate cell styles | Number=" & _
        VBA.CStr(VBA.Err.Number) & " | Description=" & VBA.Err.Description
    VBA.MsgBox "Failed to restore candidate cell styles: [" & _
        VBA.CStr(VBA.Err.Number) & "] " & VBA.Err.Description, _
        VBA.vbExclamation, "Document Generation"
End Function

Private Function private_Candidates_TryClear(ByVal candidateRange As Range) As Boolean
    On Error GoTo EH
    If Not private_Candidates_TryRestoreSavedStyles() Then Exit Function
    candidateRange.ClearContents
    Set selectedCandidateRow = Nothing
    private_Candidates_TryClear = True
    Exit Function
EH:
    ex_Helpers.LogError "Failed to clear candidate range | Number=" & _
        VBA.CStr(VBA.Err.Number) & " | Description=" & VBA.Err.Description
    VBA.MsgBox "Failed to clear candidate range: [" & _
        VBA.CStr(VBA.Err.Number) & "] " & VBA.Err.Description, _
        VBA.vbExclamation, "Document Generation"
End Function
' --------------------------------------
' } // namespace Candidates
' --------------------------------------