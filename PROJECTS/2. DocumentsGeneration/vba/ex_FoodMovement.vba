Option Explicit

Private Const CFG_FORM_TABLE_NAME As String = "wsFoodMovement::table.form"
Private Const CFG_PEOPLE_TABLE_NAME As String = "wsFoodMovement::table.people"
Private Const CFG_TARGET_TABLE_NAME As String = "wsFoodMovement::table.target"
Private Const CFG_PROVISIONS_TABLE_NAME As String = "wsFoodMovement::table.provisions"
Private Const CFG_FORM_KEY_COLUMN_NAME As String = "wsFoodMovement::column.form.key"
Private Const CFG_FORM_VALUE_COLUMN_NAME As String = "wsFoodMovement::column.form.value"
Private Const CFG_PEOPLE_RANK_COLUMN_NAME As String = "wsFoodMovement::column.people.rank"
Private Const CFG_PEOPLE_FIO_COLUMN_NAME As String = "wsFoodMovement::column.people.fio"
Private Const CFG_PEOPLE_UNIT_COLUMN_NAME As String = "wsFoodMovement::column.people.unit"
Private Const CFG_TARGET_UNIT_COLUMN_NAME As String = "wsFoodMovement::column.target.unit"
Private Const CFG_TARGET_START_COLUMN_NAME As String = "wsFoodMovement::column.target.start"
Private Const CFG_TARGET_DURATION_COLUMN_NAME As String = "wsFoodMovement::column.target.duration"
Private Const CFG_TARGET_END_COLUMN_NAME As String = "wsFoodMovement::column.target.end"
Private Const CFG_TARGET_BASIS_COLUMN_NAME As String = "wsFoodMovement::column.target.basis"
Private Const CFG_TARGET_ORDER_COLUMN_NAME As String = "wsFoodMovement::column.target.order"
Private Const CFG_PROVISIONS_UNIT_COLUMN_NAME As String = "wsFoodMovement::column.provisions.unit"
Private Const CFG_PROVISIONS_START_COLUMN_NAME As String = "wsFoodMovement::column.provisions.start"
Private Const CFG_PROVISIONS_END_COLUMN_NAME As String = "wsFoodMovement::column.provisions.end"
Private Const CFG_FORM_KEY_ACTION As String = "wsFoodMovement::key.action"
Private Const CFG_FORM_KEY_START_DATE As String = "wsFoodMovement::key.start_date"
Private Const CFG_FORM_KEY_DURATION As String = "wsFoodMovement::key.duration"
Private Const CFG_ACTION_ENROLL As String = "wsFoodMovement::action.enroll"
Private Const CFG_ACTION_PROVISIONS_ENROLL As String = "wsFoodMovement::action.provisions.enroll"
Private Const CFG_ACTION_PROVISIONS_REMOVE As String = "wsFoodMovement::action.provisions.remove"
Private Const CFG_MESSAGE_TARGET_EXTERNAL As String = "wsFoodMovement::message.target_external"
Private Const CFG_MESSAGE_ACTION_NOT_SUPPORTED As String = "wsFoodMovement::message.action_not_supported"
Private Const CFG_MESSAGE_PEOPLE_EMPTY As String = "wsFoodMovement::message.people_empty"
Private Const CFG_MESSAGE_SUCCESS_PREFIX As String = "wsFoodMovement::message.success_prefix"
Private Const CFG_MESSAGE_SUCCESS_SUFFIX As String = "wsFoodMovement::message.success_suffix"
Private Const CFG_MESSAGE_TRANSFER_FAILED_PREFIX As String = "wsFoodMovement::message.transfer_failed_prefix"
Private Const CFG_MESSAGE_TARGET_NOT_OPEN As String = "wsFoodMovement::message.target_not_open"
Private Const CFG_MESSAGE_TARGET_AMBIGUOUS As String = "wsFoodMovement::message.target_ambiguous"
Private Const CFG_MESSAGE_TABLE_NOT_FOUND_PREFIX As String = "wsFoodMovement::message.table_not_found_prefix"
Private Const CFG_MESSAGE_FIELD_EMPTY_PREFIX As String = "wsFoodMovement::message.field_empty_prefix"
Private Const CFG_MESSAGE_FORM_KEY_NOT_FOUND_PREFIX As String = "wsFoodMovement::message.form_key_not_found_prefix"
Private Const CFG_MESSAGE_FORM_SCHEMA As String = "wsFoodMovement::message.form_schema"
Private Const CFG_MESSAGE_START_DATE_INVALID As String = "wsFoodMovement::message.start_date_invalid"
Private Const CFG_MESSAGE_DURATION_INVALID As String = "wsFoodMovement::message.duration_invalid"
Private Const CFG_MESSAGE_COLUMN_NOT_FOUND_PREFIX As String = "wsFoodMovement::message.column_not_found_prefix"
Private Const CFG_MESSAGE_PREVIOUS_EVENT_OPEN_PREFIX As String = "wsFoodMovement::message.previous_event_open_prefix"
Private Const CFG_MESSAGE_PREVIOUS_EVENT_END_INVALID_PREFIX As String = "wsFoodMovement::message.previous_event_end_invalid_prefix"
Private Const CFG_MESSAGE_PROVISIONS_ACTIVE_PREFIX As String = "wsFoodMovement::message.provisions_active_prefix"
Private Const CFG_MESSAGE_PROVISIONS_NOT_ENROLLED_PREFIX As String = "wsFoodMovement::message.provisions_not_enrolled_prefix"
Private Const CFG_MESSAGE_PROVISIONS_ALREADY_CLOSED_PREFIX As String = "wsFoodMovement::message.provisions_already_closed_prefix"
Private Const CFG_MESSAGE_PROVISIONS_DATE_INVALID_PREFIX As String = "wsFoodMovement::message.provisions_date_invalid_prefix"
Private Const CFG_LOG_FILE_SUFFIX As String = "wsFoodMovement::log.file_suffix"
Private Const CFG_FORM_SHEET_NAME As String = "wsFoodMovement::sheet.form"
Private Const CFG_MESSAGE_TARGET As String = "wsFoodMovement::message.target"
Private Const CONFIG_SHEET_NAME As String = "wsConfig"
Private Const CONFIG_TABLE_NAME As String = "tbConfig"
Private Const CONFIG_KEY_COLUMN_NAME As String = "Key"
Private Const CONFIG_VALUE_COLUMN_NAME As String = "Value"

Private FORM_TABLE_NAME As String
Private PEOPLE_TABLE_NAME As String
Private TARGET_TABLE_NAME As String
Private PROVISIONS_TABLE_NAME As String
Private FORM_KEY_COLUMN_NAME As String
Private FORM_VALUE_COLUMN_NAME As String
Private PEOPLE_RANK_COLUMN_NAME As String
Private PEOPLE_FIO_COLUMN_NAME As String
Private PEOPLE_UNIT_COLUMN_NAME As String
Private TARGET_UNIT_COLUMN_NAME As String
Private TARGET_START_COLUMN_NAME As String
Private TARGET_DURATION_COLUMN_NAME As String
Private TARGET_END_COLUMN_NAME As String
Private TARGET_BASIS_COLUMN_NAME As String
Private TARGET_ORDER_COLUMN_NAME As String
Private PROVISIONS_UNIT_COLUMN_NAME As String
Private PROVISIONS_START_COLUMN_NAME As String
Private PROVISIONS_END_COLUMN_NAME As String
Private FORM_KEY_ACTION As String
Private FORM_KEY_START_DATE As String
Private FORM_KEY_DURATION As String
Private ACTION_ENROLL As String
Private ACTION_PROVISIONS_ENROLL As String
Private ACTION_PROVISIONS_REMOVE As String
Private LOG_FILE_SUFFIX As String
Private FORM_SHEET_NAME As String
Private MESSAGE_TARGET As String

' --------------------------------------
' namespace API {
' --------------------------------------
' Додає список військовослужбовців із форми до відкритої таблиці СУХПРОД.
Public Sub fn_Execute()
    Dim formTable As ListObject
    Dim peopleTable As ListObject
    Dim targetTable As ListObject
    Dim actionText As String
    Dim startDate As Date
    Dim durationDays As Long
    Dim addedCount As Long
    Dim performanceStart As Single

    On Error GoTo EH
    performanceStart = VBA.Timer
    If Not private_InitializeTextValues() Then Exit Sub
    If Not ex_Helpers.ex_TryConfigureLogFileSuffix(LOG_FILE_SUFFIX) Then Exit Sub
    ex_Helpers.LogDebug "Food movement execution started"
    private_Performance_LogCheckpoint performanceStart, "Configuration initialized"
    If Not private_TryGetLocalTable(FORM_TABLE_NAME, formTable) Then Exit Sub
    private_Performance_LogCheckpoint performanceStart, "Form table located"
    If Not private_TryGetLocalTable(PEOPLE_TABLE_NAME, peopleTable) Then Exit Sub
    private_Performance_LogCheckpoint performanceStart, "People table located"
    If Not private_TryReadForm(formTable, FORM_KEY_ACTION, actionText) Then Exit Sub
    private_Performance_LogCheckpoint performanceStart, "Action read"
    If Not private_TryReadStartDate(formTable, startDate) Then Exit Sub
    private_Performance_LogCheckpoint performanceStart, "Start date read"
    If Not private_TryValidatePeopleTable(peopleTable) Then Exit Sub
    Select Case VBA.LCase$(actionText)
        Case VBA.LCase$(ACTION_ENROLL)
            If Not private_TryFindOpenTargetTable(TARGET_TABLE_NAME, targetTable) Then Exit Sub
            If Not private_TryValidateTargetTable(targetTable) Then Exit Sub
            If Not private_TryReadDuration(formTable, durationDays) Then Exit Sub
            If Not private_TryValidatePeoplePreviousEvents(peopleTable, targetTable, startDate) Then Exit Sub
            addedCount = private_AppendPeople(peopleTable, targetTable, startDate, durationDays)
        Case VBA.LCase$(ACTION_PROVISIONS_ENROLL)
            If Not private_TryFindOpenTargetTable(PROVISIONS_TABLE_NAME, targetTable) Then Exit Sub
            If Not private_TryValidateProvisionsTable(targetTable) Then Exit Sub
            If Not private_TryValidateProvisionsEnroll(peopleTable, targetTable, startDate) Then Exit Sub
            addedCount = private_AppendProvisions(peopleTable, targetTable, startDate)
        Case VBA.LCase$(ACTION_PROVISIONS_REMOVE)
            If Not private_TryFindOpenTargetTable(PROVISIONS_TABLE_NAME, targetTable) Then Exit Sub
            If Not private_TryValidateProvisionsTable(targetTable) Then Exit Sub
            If Not private_TryValidateProvisionsRemove(peopleTable, targetTable, startDate) Then Exit Sub
            addedCount = private_CloseProvisions(peopleTable, targetTable, startDate)
        Case Else
            private_ShowError private_Message(CFG_MESSAGE_ACTION_NOT_SUPPORTED)
            Exit Sub
    End Select
    private_Performance_LogCheckpoint performanceStart, "Target operation completed"
    If addedCount = 0 Then private_ShowError private_Message(CFG_MESSAGE_PEOPLE_EMPTY): Exit Sub
    private_ShowSuccessStatus private_Message(CFG_MESSAGE_SUCCESS_PREFIX) & " " & _
        VBA.CStr(addedCount) & private_Message(CFG_MESSAGE_SUCCESS_SUFFIX)
    private_Performance_LogCheckpoint performanceStart, "Status written"
    ex_Helpers.LogDebug "Food movement execution completed | Added=" & _
        VBA.CStr(addedCount)
    private_Performance_LogCheckpoint performanceStart, "Completed"
    Exit Sub
EH:
    ex_Helpers.LogError "Food movement execution failed | Number=" & _
        VBA.CStr(VBA.Err.Number) & " | Description=" & VBA.Err.Description
    private_ShowError private_Message(CFG_MESSAGE_TRANSFER_FAILED_PREFIX) & _
        " " & VBA.Err.Description
End Sub
' --------------------------------------
' } // namespace API
' --------------------------------------

' --------------------------------------
' namespace Performance {
' --------------------------------------
' Пише у журнал тривалість виконання від натискання кнопки.
Private Sub private_Performance_LogCheckpoint( _
    ByVal startTime As Single, _
    ByVal checkpointName As String _
)
    Dim elapsedSeconds As Single

    elapsedSeconds = VBA.Timer - startTime
    If elapsedSeconds < 0 Then elapsedSeconds = elapsedSeconds + 86400!
    ex_Helpers.WriteLog "PERF | ElapsedMs=" & _
        VBA.Format$(elapsedSeconds * 1000!, "0") & " | Checkpoint=" & checkpointName
End Sub
' --------------------------------------
' } // namespace Performance
' --------------------------------------

Private Function private_InitializeTextValues() As Boolean
    If Not private_LoadConfigText(CFG_FORM_TABLE_NAME, FORM_TABLE_NAME) Then Exit Function
    If Not private_LoadConfigText(CFG_PEOPLE_TABLE_NAME, PEOPLE_TABLE_NAME) Then Exit Function
    If Not private_LoadConfigText(CFG_TARGET_TABLE_NAME, TARGET_TABLE_NAME) Then Exit Function
    If Not private_LoadConfigText(CFG_PROVISIONS_TABLE_NAME, PROVISIONS_TABLE_NAME) Then Exit Function
    If Not private_LoadConfigText(CFG_FORM_KEY_COLUMN_NAME, FORM_KEY_COLUMN_NAME) Then Exit Function
    If Not private_LoadConfigText(CFG_FORM_VALUE_COLUMN_NAME, FORM_VALUE_COLUMN_NAME) Then Exit Function
    If Not private_LoadConfigText(CFG_PEOPLE_RANK_COLUMN_NAME, PEOPLE_RANK_COLUMN_NAME) Then Exit Function
    If Not private_LoadConfigText(CFG_PEOPLE_FIO_COLUMN_NAME, PEOPLE_FIO_COLUMN_NAME) Then Exit Function
    If Not private_LoadConfigText(CFG_PEOPLE_UNIT_COLUMN_NAME, PEOPLE_UNIT_COLUMN_NAME) Then Exit Function
    If Not private_LoadConfigText(CFG_TARGET_UNIT_COLUMN_NAME, TARGET_UNIT_COLUMN_NAME) Then Exit Function
    If Not private_LoadConfigText(CFG_TARGET_START_COLUMN_NAME, TARGET_START_COLUMN_NAME) Then Exit Function
    If Not private_LoadConfigText(CFG_TARGET_DURATION_COLUMN_NAME, TARGET_DURATION_COLUMN_NAME) Then Exit Function
    If Not private_LoadConfigText(CFG_TARGET_END_COLUMN_NAME, TARGET_END_COLUMN_NAME) Then Exit Function
    If Not private_LoadConfigText(CFG_TARGET_BASIS_COLUMN_NAME, TARGET_BASIS_COLUMN_NAME) Then Exit Function
    If Not private_LoadConfigText(CFG_TARGET_ORDER_COLUMN_NAME, TARGET_ORDER_COLUMN_NAME) Then Exit Function
    If Not private_LoadConfigText(CFG_PROVISIONS_UNIT_COLUMN_NAME, PROVISIONS_UNIT_COLUMN_NAME) Then Exit Function
    If Not private_LoadConfigText(CFG_PROVISIONS_START_COLUMN_NAME, PROVISIONS_START_COLUMN_NAME) Then Exit Function
    If Not private_LoadConfigText(CFG_PROVISIONS_END_COLUMN_NAME, PROVISIONS_END_COLUMN_NAME) Then Exit Function
    If Not private_LoadConfigText(CFG_FORM_KEY_ACTION, FORM_KEY_ACTION) Then Exit Function
    If Not private_LoadConfigText(CFG_FORM_KEY_START_DATE, FORM_KEY_START_DATE) Then Exit Function
    If Not private_LoadConfigText(CFG_FORM_KEY_DURATION, FORM_KEY_DURATION) Then Exit Function
    If Not private_LoadConfigText(CFG_ACTION_ENROLL, ACTION_ENROLL) Then Exit Function
    If Not private_LoadConfigText(CFG_ACTION_PROVISIONS_ENROLL, ACTION_PROVISIONS_ENROLL) Then Exit Function
    If Not private_LoadConfigText(CFG_ACTION_PROVISIONS_REMOVE, ACTION_PROVISIONS_REMOVE) Then Exit Function
    If Not private_LoadConfigText(CFG_LOG_FILE_SUFFIX, LOG_FILE_SUFFIX) Then Exit Function
    If Not private_LoadConfigText(CFG_FORM_SHEET_NAME, FORM_SHEET_NAME) Then Exit Function
    If Not private_LoadConfigText(CFG_MESSAGE_TARGET, MESSAGE_TARGET) Then Exit Function
    private_InitializeTextValues = True
End Function

Private Function private_LoadConfigText( _
    ByVal configKey As String, _
    ByRef outValue As String _
) As Boolean
    Dim configTable As ListObject
    Dim configRow As ListRow
    Dim keyColumnIndex As Long
    Dim valueColumnIndex As Long

    On Error GoTo EH
    outValue = VBA.vbNullString
    Set configTable = ThisWorkbook.Worksheets(CONFIG_SHEET_NAME).ListObjects(CONFIG_TABLE_NAME)
    keyColumnIndex = configTable.ListColumns(CONFIG_KEY_COLUMN_NAME).Index
    valueColumnIndex = configTable.ListColumns(CONFIG_VALUE_COLUMN_NAME).Index
    For Each configRow In configTable.ListRows
        If VBA.StrComp(VBA.CStr(configRow.Range.Cells(1, keyColumnIndex).Value2), _
                configKey, VBA.vbBinaryCompare) = 0 Then
            outValue = VBA.CStr(configRow.Range.Cells(1, valueColumnIndex).Value2)
            private_LoadConfigText = True
            Exit Function
        End If
    Next configRow
    VBA.MsgBox "Configuration key was not found: " & configKey, _
        VBA.vbExclamation, "FoodMovement"
    Exit Function
EH:
    VBA.MsgBox "Cannot read wsConfig/tbConfig: " & VBA.Err.Description, _
        VBA.vbExclamation, "FoodMovement"
End Function

Private Function private_Message(ByVal configKey As String) As String
    If Not private_LoadConfigText(configKey, private_Message) Then _
        private_Message = "Configuration message is missing: " & configKey
End Function

Private Function private_TryFindOpenTargetTable( _
    ByVal tableName As String, _
    ByRef outTable As ListObject _
) As Boolean
    Dim workbookObj As Workbook
    Dim worksheetObj As Worksheet
    Dim tableObj As ListObject
    Dim matchCount As Long

    Set outTable = Nothing
    For Each workbookObj In Application.Workbooks
        For Each worksheetObj In workbookObj.Worksheets
            For Each tableObj In worksheetObj.ListObjects
                If VBA.StrComp(tableObj.Name, tableName, VBA.vbTextCompare) = 0 Then
                    matchCount = matchCount + 1
                    If matchCount = 1 Then Set outTable = tableObj
                End If
            Next tableObj
        Next worksheetObj
    Next workbookObj
    If matchCount = 0 Then
        private_ShowError private_Message(CFG_MESSAGE_TARGET_NOT_OPEN)
        Exit Function
    End If
    If matchCount > 1 Then
        private_ShowError private_Message(CFG_MESSAGE_TARGET_AMBIGUOUS)
        Set outTable = Nothing
        Exit Function
    End If
    private_TryFindOpenTargetTable = True
End Function

Private Function private_TryGetLocalTable( _
    ByVal tableName As String, _
    ByRef outTable As ListObject _
) As Boolean
    Dim worksheetObj As Worksheet
    Dim tableObj As ListObject

    Set outTable = Nothing
    For Each worksheetObj In ThisWorkbook.Worksheets
        For Each tableObj In worksheetObj.ListObjects
            If VBA.StrComp(tableObj.Name, tableName, VBA.vbTextCompare) = 0 Then
                Set outTable = tableObj
                private_TryGetLocalTable = True
                Exit Function
            End If
        Next tableObj
    Next worksheetObj
    private_ShowError private_Message(CFG_MESSAGE_TABLE_NOT_FOUND_PREFIX) & _
        " " & tableName & "."
End Function

Private Function private_TryReadForm( _
    ByVal formTable As ListObject, _
    ByVal keyText As String, _
    ByRef outValue As String _
) As Boolean
    Dim formRow As ListRow
    Dim keyIndex As Long
    Dim valueIndex As Long

    On Error GoTo EH
    keyIndex = formTable.ListColumns(FORM_KEY_COLUMN_NAME).Index
    valueIndex = formTable.ListColumns(FORM_VALUE_COLUMN_NAME).Index
    For Each formRow In formTable.ListRows
        If VBA.StrComp(private_Normalize(formRow.Range.Cells(1, keyIndex).Value2), _
                keyText, VBA.vbTextCompare) = 0 Then
            outValue = private_Normalize(formRow.Range.Cells(1, valueIndex).Value2)
            If VBA.Len(outValue) = 0 Then
                private_ShowError private_Message(CFG_MESSAGE_FIELD_EMPTY_PREFIX) & _
                    " " & keyText & "."
                Exit Function
            End If
            private_TryReadForm = True
            Exit Function
        End If
    Next formRow
    private_ShowError private_Message(CFG_MESSAGE_FORM_KEY_NOT_FOUND_PREFIX) & _
        " " & keyText & "."
    Exit Function
EH:
    private_ShowError private_Message(CFG_MESSAGE_FORM_SCHEMA)
End Function

Private Function private_TryReadStartDate( _
    ByVal formTable As ListObject, _
    ByRef outDate As Date _
) As Boolean
    Dim valueInput As Variant

    If Not private_TryReadFormValue(formTable, FORM_KEY_START_DATE, valueInput) Then Exit Function
    If VBA.IsNumeric(valueInput) Then
        outDate = VBA.CDate(VBA.CDbl(valueInput))
    ElseIf VBA.IsDate(valueInput) Then
        outDate = VBA.CDate(valueInput)
    Else
        private_ShowError private_Message(CFG_MESSAGE_START_DATE_INVALID)
        Exit Function
    End If
    private_TryReadStartDate = True
End Function

Private Function private_TryReadFormValue( _
    ByVal formTable As ListObject, _
    ByVal keyText As String, _
    ByRef outValue As Variant _
) As Boolean
    Dim formRow As ListRow
    Dim keyIndex As Long
    Dim valueIndex As Long

    On Error GoTo EH
    keyIndex = formTable.ListColumns(FORM_KEY_COLUMN_NAME).Index
    valueIndex = formTable.ListColumns(FORM_VALUE_COLUMN_NAME).Index
    For Each formRow In formTable.ListRows
        If VBA.StrComp(VBA.CStr(formRow.Range.Cells(1, keyIndex).Value2), _
                keyText, VBA.vbTextCompare) = 0 Then
            outValue = formRow.Range.Cells(1, valueIndex).Value2
            If VBA.IsEmpty(outValue) Or VBA.IsNull(outValue) Then
                private_ShowError private_Message(CFG_MESSAGE_FIELD_EMPTY_PREFIX) & keyText & "."
                Exit Function
            End If
            private_TryReadFormValue = True
            Exit Function
        End If
    Next formRow
    private_ShowError private_Message(CFG_MESSAGE_FORM_KEY_NOT_FOUND_PREFIX) & keyText & "."
    Exit Function
EH:
    private_ShowError private_Message(CFG_MESSAGE_FORM_SCHEMA)
End Function

Private Function private_TryReadDuration( _
    ByVal formTable As ListObject, _
    ByRef outDuration As Long _
) As Boolean
    Dim valueText As String

    If Not private_TryReadForm(formTable, FORM_KEY_DURATION, valueText) Then Exit Function
    If Not VBA.IsNumeric(valueText) Then
        private_ShowError private_Message(CFG_MESSAGE_DURATION_INVALID)
        Exit Function
    End If
    outDuration = VBA.CLng(valueText)
    If outDuration <= 0 Or VBA.CDbl(outDuration) <> VBA.CDbl(valueText) Then
        private_ShowError private_Message(CFG_MESSAGE_DURATION_INVALID)
        Exit Function
    End If
    private_TryReadDuration = True
End Function

Private Function private_TryValidatePeopleTable(ByVal peopleTable As ListObject) As Boolean
    private_TryValidatePeopleTable = private_TableHasColumn(peopleTable, PEOPLE_RANK_COLUMN_NAME) And _
        private_TableHasColumn(peopleTable, PEOPLE_FIO_COLUMN_NAME) And _
        private_TableHasColumn(peopleTable, PEOPLE_UNIT_COLUMN_NAME)
End Function

Private Function private_TryValidateTargetTable(ByVal targetTable As ListObject) As Boolean
    private_TryValidateTargetTable = private_TableHasColumn(targetTable, PEOPLE_RANK_COLUMN_NAME) And _
        private_TableHasColumn(targetTable, PEOPLE_FIO_COLUMN_NAME) And _
        private_TableHasColumn(targetTable, TARGET_UNIT_COLUMN_NAME) And _
        private_TableHasColumn(targetTable, TARGET_START_COLUMN_NAME) And _
        private_TableHasColumn(targetTable, TARGET_DURATION_COLUMN_NAME) And _
        private_TableHasColumn(targetTable, TARGET_END_COLUMN_NAME) And _
        private_TableHasColumn(targetTable, TARGET_BASIS_COLUMN_NAME) And _
        private_TableHasColumn(targetTable, TARGET_ORDER_COLUMN_NAME)
End Function

Private Function private_TryValidateProvisionsTable(ByVal targetTable As ListObject) As Boolean
    private_TryValidateProvisionsTable = private_TableHasColumn(targetTable, PEOPLE_RANK_COLUMN_NAME) And _
        private_TableHasColumn(targetTable, PEOPLE_FIO_COLUMN_NAME) And _
        private_TableHasColumn(targetTable, PROVISIONS_UNIT_COLUMN_NAME) And _
        private_TableHasColumn(targetTable, PROVISIONS_START_COLUMN_NAME) And _
        private_TableHasColumn(targetTable, PROVISIONS_END_COLUMN_NAME)
End Function

Private Function private_TryValidateProvisionsEnroll( _
    ByVal peopleTable As ListObject, ByVal targetTable As ListObject, ByVal operationDate As Date _
) As Boolean
    Dim personRow As ListRow, targetRow As ListRow, fioIndex As Long, targetFioIndex As Long, endIndex As Long
    Dim fioText As String, targetFioText As String, endValue As Variant, endDate As Date

    fioIndex = peopleTable.ListColumns(PEOPLE_FIO_COLUMN_NAME).Index
    targetFioIndex = targetTable.ListColumns(PEOPLE_FIO_COLUMN_NAME).Index
    endIndex = targetTable.ListColumns(PROVISIONS_END_COLUMN_NAME).Index
    For Each personRow In peopleTable.ListRows
        fioText = private_Normalize(personRow.Range.Cells(1, fioIndex).Value2)
        If VBA.Len(fioText) > 0 Then
            For Each targetRow In targetTable.ListRows
                targetFioText = private_Normalize(targetRow.Range.Cells(1, targetFioIndex).Value2)
                If VBA.StrComp(private_PersonKey(targetFioText), private_PersonKey(fioText), VBA.vbBinaryCompare) = 0 Then
                    endValue = targetRow.Range.Cells(1, endIndex).Value2
                    If VBA.Len(private_Normalize(endValue)) = 0 Then
                        private_ShowError private_Message(CFG_MESSAGE_PROVISIONS_ACTIVE_PREFIX) & " " & fioText & ".": Exit Function
                    End If
                    If Not private_TryReadTargetEndDate(endValue, fioText, endDate) Then Exit Function
                    If operationDate < endDate Then
                        private_ShowError private_Message(CFG_MESSAGE_PROVISIONS_ACTIVE_PREFIX) & " " & fioText & ".": Exit Function
                    End If
                End If
            Next targetRow
        End If
    Next personRow
    private_TryValidateProvisionsEnroll = True
End Function

Private Function private_TryValidateProvisionsRemove( _
    ByVal peopleTable As ListObject, ByVal targetTable As ListObject, ByVal operationDate As Date _
) As Boolean
    Dim personRow As ListRow, targetRow As ListRow, fioIndex As Long, targetFioIndex As Long, startIndex As Long, endIndex As Long
    Dim fioText As String, endText As String, startDate As Date, activeCount As Long, historyCount As Long

    fioIndex = peopleTable.ListColumns(PEOPLE_FIO_COLUMN_NAME).Index
    targetFioIndex = targetTable.ListColumns(PEOPLE_FIO_COLUMN_NAME).Index
    startIndex = targetTable.ListColumns(PROVISIONS_START_COLUMN_NAME).Index
    endIndex = targetTable.ListColumns(PROVISIONS_END_COLUMN_NAME).Index
    For Each personRow In peopleTable.ListRows
        fioText = private_Normalize(personRow.Range.Cells(1, fioIndex).Value2): activeCount = 0: historyCount = 0
        If VBA.Len(fioText) > 0 Then
            For Each targetRow In targetTable.ListRows
                If VBA.StrComp(private_PersonKey(targetRow.Range.Cells(1, targetFioIndex).Value2), private_PersonKey(fioText), VBA.vbBinaryCompare) = 0 Then
                    historyCount = historyCount + 1: endText = private_Normalize(targetRow.Range.Cells(1, endIndex).Value2)
                    If VBA.Len(endText) = 0 Then
                        activeCount = activeCount + 1
                        If Not private_TryReadTargetEndDate(targetRow.Range.Cells(1, startIndex).Value2, fioText, startDate) Then Exit Function
                        If operationDate < startDate Then private_ShowError private_Message(CFG_MESSAGE_PROVISIONS_DATE_INVALID_PREFIX) & " " & fioText & ".": Exit Function
                    End If
                End If
            Next targetRow
            If activeCount = 0 Then
                If historyCount = 0 Then private_ShowError private_Message(CFG_MESSAGE_PROVISIONS_NOT_ENROLLED_PREFIX) & " " & fioText & "." Else private_ShowError private_Message(CFG_MESSAGE_PROVISIONS_ALREADY_CLOSED_PREFIX) & " " & fioText & "."
                Exit Function
            End If
            If activeCount > 1 Then private_ShowError private_Message(CFG_MESSAGE_PROVISIONS_ACTIVE_PREFIX) & " " & fioText & ".": Exit Function
        End If
    Next personRow
    private_TryValidateProvisionsRemove = True
End Function

Private Function private_AppendProvisions( _
    ByVal peopleTable As ListObject, ByVal targetTable As ListObject, ByVal operationDate As Date _
) As Long
    Dim personRow As ListRow, targetRow As ListRow, rankIndex As Long, fioIndex As Long, unitIndex As Long, fioText As String
    rankIndex = peopleTable.ListColumns(PEOPLE_RANK_COLUMN_NAME).Index: fioIndex = peopleTable.ListColumns(PEOPLE_FIO_COLUMN_NAME).Index: unitIndex = peopleTable.ListColumns(PEOPLE_UNIT_COLUMN_NAME).Index
    For Each personRow In peopleTable.ListRows
        fioText = private_Normalize(personRow.Range.Cells(1, fioIndex).Value2)
        If VBA.Len(fioText) > 0 Then
            Set targetRow = targetTable.ListRows.Add
            targetRow.Range.Cells(1, targetTable.ListColumns(PEOPLE_RANK_COLUMN_NAME).Index).Value = private_Normalize(personRow.Range.Cells(1, rankIndex).Value2)
            targetRow.Range.Cells(1, targetTable.ListColumns(PEOPLE_FIO_COLUMN_NAME).Index).Value = fioText
            targetRow.Range.Cells(1, targetTable.ListColumns(PROVISIONS_UNIT_COLUMN_NAME).Index).Value = private_Normalize(personRow.Range.Cells(1, unitIndex).Value2)
            targetRow.Range.Cells(1, targetTable.ListColumns(PROVISIONS_START_COLUMN_NAME).Index).Value = operationDate
            private_AppendProvisions = private_AppendProvisions + 1
        End If
    Next personRow
End Function

Private Function private_CloseProvisions( _
    ByVal peopleTable As ListObject, ByVal targetTable As ListObject, ByVal operationDate As Date _
) As Long
    Dim personRow As ListRow, targetRow As ListRow, fioIndex As Long, targetFioIndex As Long, endIndex As Long, fioText As String
    fioIndex = peopleTable.ListColumns(PEOPLE_FIO_COLUMN_NAME).Index: targetFioIndex = targetTable.ListColumns(PEOPLE_FIO_COLUMN_NAME).Index: endIndex = targetTable.ListColumns(PROVISIONS_END_COLUMN_NAME).Index
    For Each personRow In peopleTable.ListRows
        fioText = private_Normalize(personRow.Range.Cells(1, fioIndex).Value2)
        If VBA.Len(fioText) > 0 Then
            For Each targetRow In targetTable.ListRows
                If VBA.StrComp(private_PersonKey(targetRow.Range.Cells(1, targetFioIndex).Value2), private_PersonKey(fioText), VBA.vbBinaryCompare) = 0 And VBA.Len(private_Normalize(targetRow.Range.Cells(1, endIndex).Value2)) = 0 Then
                    targetRow.Range.Cells(1, endIndex).Value = operationDate: private_CloseProvisions = private_CloseProvisions + 1
                End If
            Next targetRow
        End If
    Next personRow
End Function

Private Function private_TableHasColumn( _
    ByVal tableObj As ListObject, _
    ByVal columnName As String _
) As Boolean
    Dim columnIndex As Long

    On Error GoTo EH
    columnIndex = tableObj.ListColumns(columnName).Index
    private_TableHasColumn = True
    Exit Function
EH:
    private_ShowError private_Message(CFG_MESSAGE_COLUMN_NOT_FOUND_PREFIX) & _
        " " & tableObj.Name & ": " & columnName & "."
End Function

Private Function private_TryValidatePeoplePreviousEvents( _
    ByVal peopleTable As ListObject, _
    ByVal targetTable As ListObject, _
    ByVal startDate As Date _
) As Boolean
    Dim personRow As ListRow
    Dim targetRow As ListRow
    Dim peopleFioIndex As Long
    Dim targetFioIndex As Long
    Dim targetEndIndex As Long
    Dim fioText As String
    Dim personKey As String
    Dim targetFioText As String
    Dim targetPersonKey As String
    Dim endDate As Date
    Dim latestEndDate As Date
    Dim hasPreviousEvent As Boolean
    Dim matchingEventsCount As Long

    peopleFioIndex = peopleTable.ListColumns(PEOPLE_FIO_COLUMN_NAME).Index
    targetFioIndex = targetTable.ListColumns(PEOPLE_FIO_COLUMN_NAME).Index
    targetEndIndex = targetTable.ListColumns(TARGET_END_COLUMN_NAME).Index

    For Each personRow In peopleTable.ListRows
        fioText = private_Normalize(personRow.Range.Cells(1, peopleFioIndex).Value2)
        If VBA.Len(fioText) > 0 Then
            personKey = private_PersonKey(fioText)
            hasPreviousEvent = False
            matchingEventsCount = 0
            For Each targetRow In targetTable.ListRows
                targetFioText = private_Normalize( _
                    targetRow.Range.Cells(1, targetFioIndex).Value2)
                targetPersonKey = private_PersonKey(targetFioText)
                If VBA.StrComp(targetPersonKey, personKey, VBA.vbBinaryCompare) = 0 Then
                    If Not private_TryReadTargetEndDate( _
                        targetRow.Range.Cells(1, targetEndIndex).Value2, _
                        fioText, endDate) Then Exit Function
                    matchingEventsCount = matchingEventsCount + 1
                    If Not hasPreviousEvent Or endDate > latestEndDate Then
                        latestEndDate = endDate
                        hasPreviousEvent = True
                    End If
                End If
            Next targetRow
            ex_Helpers.LogDebug "Food movement previous-event validation | Person=" & _
                fioText & " | Matches=" & VBA.CStr(matchingEventsCount) & _
                " | LatestEnd=" & IIf(hasPreviousEvent, _
                VBA.Format$(latestEndDate, "yyyy-mm-dd"), "<none>") & _
                " | Start=" & VBA.Format$(startDate, "yyyy-mm-dd")
            If hasPreviousEvent And startDate < latestEndDate Then
                private_ShowError private_Message(CFG_MESSAGE_PREVIOUS_EVENT_OPEN_PREFIX) & _
                    " " & fioText & "."
                Exit Function
            End If
        End If
    Next personRow
    private_TryValidatePeoplePreviousEvents = True
End Function

Private Function private_PersonKey(ByVal valueInput As Variant) As String
    Dim normalizedText As String

    normalizedText = private_Normalize(valueInput)
    normalizedText = VBA.Replace$(normalizedText, VBA.ChrW$(160), " ")
    normalizedText = VBA.Replace$(normalizedText, VBA.vbCr, " ")
    normalizedText = VBA.Replace$(normalizedText, VBA.vbLf, " ")
    private_PersonKey = VBA.UCase$(Application.WorksheetFunction.Trim(normalizedText))
End Function

Private Function private_TryReadTargetEndDate( _
    ByVal valueInput As Variant, _
    ByVal fioText As String, _
    ByRef outDate As Date _
) As Boolean
    If VBA.IsNumeric(valueInput) Then
        outDate = VBA.CDate(VBA.CDbl(valueInput))
    ElseIf VBA.IsDate(valueInput) Then
        outDate = VBA.CDate(valueInput)
    Else
        private_ShowError private_Message( _
            CFG_MESSAGE_PREVIOUS_EVENT_END_INVALID_PREFIX) & " " & fioText & "."
        Exit Function
    End If
    private_TryReadTargetEndDate = True
End Function

Private Function private_AppendPeople( _
    ByVal peopleTable As ListObject, _
    ByVal targetTable As ListObject, _
    ByVal startDate As Date, _
    ByVal durationDays As Long _
) As Long
    Dim personRow As ListRow
    Dim rankIndex As Long
    Dim fioIndex As Long
    Dim unitIndex As Long
    Dim fioText As String
    Dim peopleCount As Long
    Dim valueRowIndex As Long
    Dim firstTargetDataRow As Long
    Dim rankValues() As Variant
    Dim fioValues() As Variant
    Dim unitValues() As Variant
    Dim startValues() As Variant
    Dim durationValues() As Variant
    Dim endValues() As Variant
    Dim screenUpdatingEnabled As Boolean
    Dim eventsEnabled As Boolean
    Dim originalCalculation As XlCalculation
    Dim errorNumber As Long
    Dim errorSource As String
    Dim errorDescription As String

    rankIndex = peopleTable.ListColumns(PEOPLE_RANK_COLUMN_NAME).Index
    fioIndex = peopleTable.ListColumns(PEOPLE_FIO_COLUMN_NAME).Index
    unitIndex = peopleTable.ListColumns(PEOPLE_UNIT_COLUMN_NAME).Index
    For Each personRow In peopleTable.ListRows
        fioText = private_Normalize(personRow.Range.Cells(1, fioIndex).Value2)
        If VBA.Len(fioText) > 0 Then peopleCount = peopleCount + 1
    Next personRow
    If peopleCount = 0 Then Exit Function

    ReDim rankValues(1 To peopleCount, 1 To 1)
    ReDim fioValues(1 To peopleCount, 1 To 1)
    ReDim unitValues(1 To peopleCount, 1 To 1)
    ReDim startValues(1 To peopleCount, 1 To 1)
    ReDim durationValues(1 To peopleCount, 1 To 1)
    ReDim endValues(1 To peopleCount, 1 To 1)
    For Each personRow In peopleTable.ListRows
        fioText = private_Normalize(personRow.Range.Cells(1, fioIndex).Value2)
        If VBA.Len(fioText) > 0 Then
            valueRowIndex = valueRowIndex + 1
            rankValues(valueRowIndex, 1) = private_Normalize( _
                personRow.Range.Cells(1, rankIndex).Value2)
            fioValues(valueRowIndex, 1) = fioText
            unitValues(valueRowIndex, 1) = private_Normalize( _
                personRow.Range.Cells(1, unitIndex).Value2)
            startValues(valueRowIndex, 1) = VBA.CDbl(startDate)
            durationValues(valueRowIndex, 1) = durationDays
            endValues(valueRowIndex, 1) = VBA.CDbl( _
                VBA.DateAdd("d", durationDays - 1, startDate))
        End If
    Next personRow

    screenUpdatingEnabled = Application.ScreenUpdating
    eventsEnabled = Application.EnableEvents
    originalCalculation = Application.Calculation
    On Error GoTo EH
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual

    firstTargetDataRow = targetTable.ListRows.Count + 1
    targetTable.Resize targetTable.Range.Resize( _
        targetTable.Range.Rows.Count + peopleCount, targetTable.Range.Columns.Count)
    targetTable.DataBodyRange.Cells(firstTargetDataRow, _
        targetTable.ListColumns(PEOPLE_RANK_COLUMN_NAME).Index).Resize(peopleCount, 1).Value2 = rankValues
    targetTable.DataBodyRange.Cells(firstTargetDataRow, _
        targetTable.ListColumns(PEOPLE_FIO_COLUMN_NAME).Index).Resize(peopleCount, 1).Value2 = fioValues
    targetTable.DataBodyRange.Cells(firstTargetDataRow, _
        targetTable.ListColumns(TARGET_UNIT_COLUMN_NAME).Index).Resize(peopleCount, 1).Value2 = unitValues
    targetTable.DataBodyRange.Cells(firstTargetDataRow, _
        targetTable.ListColumns(TARGET_START_COLUMN_NAME).Index).Resize(peopleCount, 1).Value2 = startValues
    targetTable.DataBodyRange.Cells(firstTargetDataRow, _
        targetTable.ListColumns(TARGET_DURATION_COLUMN_NAME).Index).Resize(peopleCount, 1).Value2 = durationValues
    targetTable.DataBodyRange.Cells(firstTargetDataRow, _
        targetTable.ListColumns(TARGET_END_COLUMN_NAME).Index).Resize(peopleCount, 1).Value2 = endValues
    private_AppendPeople = peopleCount

CleanExit:
    Application.Calculation = originalCalculation
    Application.EnableEvents = eventsEnabled
    Application.ScreenUpdating = screenUpdatingEnabled
    Exit Function
EH:
    errorNumber = VBA.Err.Number
    errorSource = VBA.Err.Source
    errorDescription = VBA.Err.Description
    On Error Resume Next
    Application.Calculation = originalCalculation
    Application.EnableEvents = eventsEnabled
    Application.ScreenUpdating = screenUpdatingEnabled
    On Error GoTo 0
    VBA.Err.Raise errorNumber, errorSource, errorDescription
End Function

Private Function private_Normalize(ByVal valueInput As Variant) As String
    If VBA.IsError(valueInput) Or VBA.IsNull(valueInput) Or VBA.IsEmpty(valueInput) Then Exit Function
    private_Normalize = VBA.Trim$(VBA.CStr(valueInput))
End Function

Private Sub private_ShowError(ByVal messageText As String)
    private_ShowStatus messageText, True
End Sub

Private Sub private_ShowSuccessStatus(ByVal messageText As String)
    private_ShowStatus messageText, False
End Sub

Private Sub private_ShowStatus( _
    ByVal messageText As String, _
    ByVal isErrorMessage As Boolean _
)
    If Not ex_Helpers.ex_TryConfigureMessageTarget( _
        FORM_SHEET_NAME, MESSAGE_TARGET) Then Exit Sub

    ex_Helpers.ex_ShowStatusMessage messageText, isErrorMessage
    ex_Helpers.ex_ClearMessageTarget
End Sub