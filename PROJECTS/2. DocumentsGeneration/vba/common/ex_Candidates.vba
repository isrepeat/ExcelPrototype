Option Explicit

Private configuredInputSheetName As String
Private configuredLookupCellAddresses As Collection
Private configuredCandidateStartCellAddress As String
Private configuredMaxCandidateCount As Long
Private configuredQueryCallbackName As String
Private configuredSelectedValueIndex As Long
Private configuredColumns As Collection
Private configuredFontColor As Long
Private configuredFillColor As Long
Private configuredSelectedFillColor As Long
Private configuredFontName As String
Private configuredFontSize As Double
Private configuredHorizontalAlignment As Long
Private configuredVerticalAlignment As Long
Private configuredWrapText As Boolean
Private originalCandidateCellStyles As Object
Private activeLookupCellAddress As String
Private selectedCandidateRow As Range
Private isConfigured As Boolean

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
    isConfigured = True
    ex_Helpers.LogDebug "Candidates configured | Sheet=" & configuredInputSheetName & _
        " | LookupCells=" & VBA.CStr(configuredLookupCellAddresses.Count) & _
        " | CandidateStartCell=" & configuredCandidateStartCellAddress & _
        " | QueryCallback=" & configuredQueryCallbackName
    fn_Configure = True
End Function

Public Function fn_RegisterRoutes() As Boolean
    Dim lookupCellAddress As Variant
    Dim sourceSheet As Worksheet
    Dim candidateRange As Range
    Dim clearCandidatesCell As Range

    On Error GoTo EH
    If Not private_Candidates_TryEnsureConfigured() Then Exit Function
    Set sourceSheet = ThisWorkbook.Worksheets(configuredInputSheetName)
    If Not private_Candidates_TryGetCandidateRange(sourceSheet, candidateRange) Then Exit Function
    If candidateRange.Row = 1 Then
        VBA.MsgBox "Candidate start cell must not be in the first worksheet row.", _
            VBA.vbExclamation, "Document Generation"
        Exit Function
    End If
    Set clearCandidatesCell = candidateRange.Cells(1, 1).Offset(-1, 0)
    For Each lookupCellAddress In configuredLookupCellAddresses
        If Not ex_CellChangeRouter.fn_RegisterChangeRoute( _
            configuredInputSheetName, VBA.CStr(lookupCellAddress), _
            "ex_Candidates.fn_OnInputChanged") Then Exit Function
    Next lookupCellAddress
    If Not ex_CellChangeRouter.fn_RegisterSelectionRoute( _
        configuredInputSheetName, candidateRange.Address(False, False), _
        "ex_Candidates.fn_OnCandidateSelected") Then Exit Function
    If Not ex_CellChangeRouter.fn_RegisterSelectionRoute( _
        configuredInputSheetName, clearCandidatesCell.Address(False, False), _
        "ex_Candidates.fn_OnClearCandidatesRequested") Then Exit Function
    ex_Helpers.LogDebug "Candidate routes registered | Sheet=" & _
        configuredInputSheetName & " | LookupCells=" & _
        VBA.CStr(configuredLookupCellAddresses.Count) & " | CandidateRange=" & _
        candidateRange.Address(False, False) & " | ClearCell=" & _
        clearCandidatesCell.Address(False, False)
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

' Очищает выдачу по клику на ячейку заголовка над таблицей кандидатов.
Public Sub fn_OnClearCandidatesRequested(ByVal changedSheet As Object, ByVal target As Range)
    Dim sourceSheet As Worksheet
    Dim candidateRange As Range

    If changedSheet Is Nothing Or target Is Nothing Then Exit Sub
    If Not TypeOf changedSheet Is Worksheet Then Exit Sub
    If Not private_Candidates_TryEnsureConfigured() Then Exit Sub
    Set sourceSheet = changedSheet
    If Not private_Candidates_TryValidateSourceSheet(sourceSheet) Then Exit Sub
    If Not private_Candidates_TryGetCandidateRange(sourceSheet, candidateRange) Then Exit Sub
    If Not private_Candidates_TryClear(candidateRange) Then Exit Sub
    activeLookupCellAddress = VBA.vbNullString
    ex_Helpers.LogDebug "Candidate table cleared by clear cell | Sheet=" & _
        sourceSheet.Name & " | Cell=" & target.Address(False, False)
End Sub
' --------------------------------------
' } // namespace API
' --------------------------------------

' --------------------------------------
' namespace Candidates {
' --------------------------------------
Private Function private_Candidates_TryReadConfig(ByVal candidatesConfig As Object) As Boolean
    Dim lookupCellAddresses As Collection
    Dim styleConfig As Object

    isConfigured = False
    configuredInputSheetName = VBA.vbNullString
    Set configuredLookupCellAddresses = Nothing
    configuredCandidateStartCellAddress = VBA.vbNullString
    configuredMaxCandidateCount = 0
    configuredQueryCallbackName = VBA.vbNullString
    configuredSelectedValueIndex = 0
    Set configuredColumns = Nothing
    configuredFontColor = 0
    configuredFillColor = 0
    configuredSelectedFillColor = 0
    configuredFontName = VBA.vbNullString
    configuredFontSize = 0
    configuredHorizontalAlignment = 0
    configuredVerticalAlignment = 0
    configuredWrapText = False
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
    If Not private_Candidates_TryGetRequiredText(candidatesConfig, _
        "CandidateStartCellAddress", configuredCandidateStartCellAddress) Then Exit Function
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
        candidatesConfig, "Style", styleConfig) Then Exit Function
    If Not private_Candidates_TryGetRequiredLong( _
        styleConfig, "FontColor", configuredFontColor) Then Exit Function
    If Not private_Candidates_TryGetRequiredLong( _
        styleConfig, "FillColor", configuredFillColor) Then Exit Function
    If Not private_Candidates_TryGetRequiredLong( _
        styleConfig, "SelectedFillColor", configuredSelectedFillColor) Then Exit Function
    If Not private_Candidates_TryGetRequiredText( _
        styleConfig, "FontName", configuredFontName) Then Exit Function
    If Not private_Candidates_TryGetRequiredDouble( _
        styleConfig, "FontSize", configuredFontSize) Then Exit Function
    If configuredFontSize <= 0 Then
        VBA.MsgBox "Candidate style FontSize must be greater than zero.", _
            VBA.vbExclamation, "Document Generation"
        Exit Function
    End If
    If Not private_Candidates_TryGetRequiredLong( _
        styleConfig, "HorizontalAlignment", configuredHorizontalAlignment) Then Exit Function
    If Not private_Candidates_TryGetRequiredLong( _
        styleConfig, "VerticalAlignment", configuredVerticalAlignment) Then Exit Function
    If Not private_Candidates_TryGetRequiredBoolean( _
        styleConfig, "WrapText", configuredWrapText) Then Exit Function
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
    Set outCandidateRange = candidateStartCell.Resize( _
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
    Dim candidateItem As Variant
    Dim candidateRange As Range
    Dim outputRow As Range

    If Not private_Candidates_TryGetCandidateRange(sourceSheet, candidateRange) Then Exit Sub
    If Not private_Candidates_TryClear(candidateRange) Then Exit Sub
    inputText = ex_Helpers.private_Text_Normalize(VBA.CStr(lookupCell.Value2))
    If VBA.Len(inputText) = 0 Then
        ex_Helpers.LogDebug "Candidate search cleared | Sheet=" & sourceSheet.Name & _
            " | Cell=" & activeLookupCellAddress
        activeLookupCellAddress = VBA.vbNullString
        Exit Sub
    End If
    If Not private_Candidates_TryFindCandidates(inputText, candidates) Then Exit Sub
    If candidates.Count > candidateRange.Rows.Count Then
        VBA.MsgBox "Candidate query returned more rows than configured maximum.", _
            VBA.vbExclamation, "Document Generation"
        Exit Sub
    End If
    ex_Helpers.LogDebug "Candidate search completed | Sheet=" & sourceSheet.Name & _
        " | Cell=" & activeLookupCellAddress & " | Count=" & VBA.CStr(candidates.Count)
    Set outputRow = candidateRange.Rows(1)
    For Each candidateItem In candidates
        If Not private_Candidates_TryWriteCandidate(outputRow, candidateItem) Then Exit Sub
        Set outputRow = outputRow.Offset(1, 0)
    Next candidateItem
End Sub

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

Private Function private_Candidates_TryWriteCandidate(ByVal outputRow As Range, _
    ByVal candidateItem As Variant) As Boolean
    Dim columnConfig As Object
    Dim sourceIndex As Long
    Dim numberFormat As String
    Dim outputColumnIndex As Long

    If Not VBA.IsArray(candidateItem) Then GoTo InvalidCandidate
    On Error GoTo InvalidCandidate
    outputColumnIndex = 0
    For Each columnConfig In configuredColumns
        outputColumnIndex = outputColumnIndex + 1
        sourceIndex = VBA.CLng(columnConfig.Item("SourceIndex"))
        numberFormat = VBA.CStr(columnConfig.Item("NumberFormat"))
        If Not private_Candidates_TrySaveCellStyle( _
            outputRow.Cells(1, outputColumnIndex)) Then Exit Function
        outputRow.Cells(1, outputColumnIndex).Font.Name = configuredFontName
        outputRow.Cells(1, outputColumnIndex).Font.Size = configuredFontSize
        outputRow.Cells(1, outputColumnIndex).Font.Color = configuredFontColor
        outputRow.Cells(1, outputColumnIndex).Interior.Color = configuredFillColor
        outputRow.Cells(1, outputColumnIndex).HorizontalAlignment = _
            configuredHorizontalAlignment
        outputRow.Cells(1, outputColumnIndex).VerticalAlignment = _
            configuredVerticalAlignment
        outputRow.Cells(1, outputColumnIndex).WrapText = configuredWrapText
        outputRow.Cells(1, outputColumnIndex).NumberFormat = numberFormat
        outputRow.Cells(1, outputColumnIndex).Value = VBA.CStr(candidateItem(sourceIndex))
    Next columnConfig
    private_Candidates_TryWriteCandidate = True
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
    private_Candidates_ApplySelectedCellStyle candidateCell
    ex_Helpers.LogDebug "Candidate accepted | Sheet=" & sourceSheet.Name & _
        " | TargetCell=" & activeLookupCellAddress & " | CandidateRow=" & _
        VBA.CStr(candidateCell.Row)
    Application.EnableEvents = eventsWereEnabled
    Exit Sub
EH:
    Application.EnableEvents = eventsWereEnabled
    ex_Helpers.LogError "Failed to accept candidate | Number=" & _
        VBA.CStr(VBA.Err.Number) & " | Description=" & VBA.Err.Description
    ex_Helpers.ex_ShowErrorMessage "Failed to accept candidate: " & _
        VBA.Err.Description, VBA.vbExclamation, "Document Generation"
End Sub

Private Sub private_Candidates_ApplySelectedCellStyle( _
    ByVal candidateCell As Range _
)
    Dim candidateRow As Range

    Set candidateRow = candidateCell.Worksheet.Cells( _
        candidateCell.Row, candidateCell.Worksheet.Range( _
        configuredCandidateStartCellAddress).Column).Resize( _
        1, configuredColumns.Count)
    If Not selectedCandidateRow Is Nothing Then
        selectedCandidateRow.Interior.Color = configuredFillColor
    End If
    candidateRow.Interior.Color = configuredSelectedFillColor
    Set selectedCandidateRow = candidateRow
End Sub

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