Option Explicit

Private Const CFG_FORM_TABLE_NAME As String = "wsFoodMovement::table.form"
Private Const CFG_PEOPLE_TABLE_NAME As String = "wsFoodMovement::table.people"
Private Const CFG_TARGET_TABLE_NAME As String = "wsFoodMovement::table.target"
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
Private Const CFG_FORM_KEY_ACTION As String = "wsFoodMovement::key.action"
Private Const CFG_FORM_KEY_START_DATE As String = "wsFoodMovement::key.start_date"
Private Const CFG_FORM_KEY_DURATION As String = "wsFoodMovement::key.duration"
Private Const CFG_ACTION_ENROLL As String = "wsFoodMovement::action.enroll"
Private Const CFG_MESSAGE_TITLE As String = "wsFoodMovement::message.title"
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
Private Const CFG_LOG_FILE_SUFFIX As String = "wsFoodMovement::log.file_suffix"
Private Const CONFIG_SHEET_NAME As String = "wsConfig"
Private Const CONFIG_TABLE_NAME As String = "tbConfig"
Private Const CONFIG_KEY_COLUMN_NAME As String = "Key"
Private Const CONFIG_VALUE_COLUMN_NAME As String = "Value"

Private FORM_TABLE_NAME As String
Private PEOPLE_TABLE_NAME As String
Private TARGET_TABLE_NAME As String
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
Private FORM_KEY_ACTION As String
Private FORM_KEY_START_DATE As String
Private FORM_KEY_DURATION As String
Private ACTION_ENROLL As String
Private MESSAGE_TITLE As String
Private LOG_FILE_SUFFIX As String

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

    On Error GoTo EH
    If Not private_InitializeTextValues() Then Exit Sub
    If Not ex_Helpers.ex_TryConfigureLogFileSuffix(LOG_FILE_SUFFIX) Then Exit Sub
    ex_Helpers.LogDebug "Food movement execution started"
    If Not private_TryGetLocalTable(FORM_TABLE_NAME, formTable) Then Exit Sub
    If Not private_TryGetLocalTable(PEOPLE_TABLE_NAME, peopleTable) Then Exit Sub
    If Not private_TryFindOpenTargetTable(targetTable) Then Exit Sub
    If targetTable.Parent.Parent Is ThisWorkbook Then
        private_ShowError private_Message(CFG_MESSAGE_TARGET_EXTERNAL)
        Exit Sub
    End If
    If Not private_TryReadForm(formTable, FORM_KEY_ACTION, actionText) Then Exit Sub
    If VBA.StrComp(actionText, ACTION_ENROLL, VBA.vbTextCompare) <> 0 Then
        private_ShowError private_Message(CFG_MESSAGE_ACTION_NOT_SUPPORTED)
        Exit Sub
    End If
    If Not private_TryReadStartDate(formTable, startDate) Then Exit Sub
    If Not private_TryReadDuration(formTable, durationDays) Then Exit Sub
    If Not private_TryValidatePeopleTable(peopleTable) Then Exit Sub
    If Not private_TryValidateTargetTable(targetTable) Then Exit Sub

    addedCount = private_AppendPeople(peopleTable, targetTable, startDate, durationDays)
    If addedCount = 0 Then
        private_ShowError private_Message(CFG_MESSAGE_PEOPLE_EMPTY)
        Exit Sub
    End If
    private_ShowMessage private_Message(CFG_MESSAGE_SUCCESS_PREFIX) & _
        " " & VBA.CStr(addedCount) & private_Message(CFG_MESSAGE_SUCCESS_SUFFIX), _
        VBA.vbInformation
    ex_Helpers.LogDebug "Food movement execution completed | Added=" & _
        VBA.CStr(addedCount)
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

Private Function private_InitializeTextValues() As Boolean
    If Not private_LoadConfigText(CFG_FORM_TABLE_NAME, FORM_TABLE_NAME) Then Exit Function
    If Not private_LoadConfigText(CFG_PEOPLE_TABLE_NAME, PEOPLE_TABLE_NAME) Then Exit Function
    If Not private_LoadConfigText(CFG_TARGET_TABLE_NAME, TARGET_TABLE_NAME) Then Exit Function
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
    If Not private_LoadConfigText(CFG_FORM_KEY_ACTION, FORM_KEY_ACTION) Then Exit Function
    If Not private_LoadConfigText(CFG_FORM_KEY_START_DATE, FORM_KEY_START_DATE) Then Exit Function
    If Not private_LoadConfigText(CFG_FORM_KEY_DURATION, FORM_KEY_DURATION) Then Exit Function
    If Not private_LoadConfigText(CFG_ACTION_ENROLL, ACTION_ENROLL) Then Exit Function
    If Not private_LoadConfigText(CFG_MESSAGE_TITLE, MESSAGE_TITLE) Then Exit Function
    If Not private_LoadConfigText(CFG_LOG_FILE_SUFFIX, LOG_FILE_SUFFIX) Then Exit Function
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
                If VBA.StrComp(tableObj.Name, TARGET_TABLE_NAME, VBA.vbTextCompare) = 0 Then
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

Private Function private_AppendPeople( _
    ByVal peopleTable As ListObject, _
    ByVal targetTable As ListObject, _
    ByVal startDate As Date, _
    ByVal durationDays As Long _
) As Long
    Dim personRow As ListRow
    Dim targetRow As ListRow
    Dim rankIndex As Long
    Dim fioIndex As Long
    Dim unitIndex As Long
    Dim fioText As String

    rankIndex = peopleTable.ListColumns(PEOPLE_RANK_COLUMN_NAME).Index
    fioIndex = peopleTable.ListColumns(PEOPLE_FIO_COLUMN_NAME).Index
    unitIndex = peopleTable.ListColumns(PEOPLE_UNIT_COLUMN_NAME).Index
    For Each personRow In peopleTable.ListRows
        fioText = private_Normalize(personRow.Range.Cells(1, fioIndex).Value2)
        If VBA.Len(fioText) > 0 Then
            Set targetRow = targetTable.ListRows.Add
            targetRow.Range.Cells(1, targetTable.ListColumns(PEOPLE_RANK_COLUMN_NAME).Index).Value = _
                private_Normalize(personRow.Range.Cells(1, rankIndex).Value2)
            targetRow.Range.Cells(1, targetTable.ListColumns(PEOPLE_FIO_COLUMN_NAME).Index).Value = fioText
            targetRow.Range.Cells(1, targetTable.ListColumns(TARGET_UNIT_COLUMN_NAME).Index).Value = _
                private_Normalize(personRow.Range.Cells(1, unitIndex).Value2)
            targetRow.Range.Cells(1, targetTable.ListColumns(TARGET_START_COLUMN_NAME).Index).Value = startDate
            targetRow.Range.Cells(1, targetTable.ListColumns(TARGET_DURATION_COLUMN_NAME).Index).Value = durationDays
            targetRow.Range.Cells(1, targetTable.ListColumns(TARGET_END_COLUMN_NAME).Index).Value = _
                VBA.DateAdd("d", durationDays - 1, startDate)
            private_AppendPeople = private_AppendPeople + 1
        End If
    Next personRow
End Function

Private Function private_Normalize(ByVal valueInput As Variant) As String
    If VBA.IsError(valueInput) Or VBA.IsNull(valueInput) Or VBA.IsEmpty(valueInput) Then Exit Function
    private_Normalize = VBA.Trim$(VBA.CStr(valueInput))
End Function

Private Sub private_ShowError(ByVal messageText As String)
    private_ShowMessage messageText, VBA.vbExclamation
End Sub

Private Sub private_ShowMessage(ByVal messageText As String, ByVal messageIcon As VbMsgBoxStyle)
    Call ex_Helpers.ex_ShowMessage(messageText, messageIcon, MESSAGE_TITLE)
End Sub