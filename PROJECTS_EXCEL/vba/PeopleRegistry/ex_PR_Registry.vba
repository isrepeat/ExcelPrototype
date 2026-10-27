Option Explicit

Private m_originAliases As Object
Private m_database As Workbook
Private m_savedID As String
Private m_savedName As String
Private m_savedOrigin As String
Private m_busy As Boolean

' --------------------------------------
' namespace Lifecycle {
' --------------------------------------
Public Sub fn_Module_Dispose()
    Set m_originAliases = Nothing
    Set m_database = Nothing
    m_savedID = vbNullString
    m_savedName = vbNullString
    m_savedOrigin = vbNullString
End Sub
' --------------------------------------
' } // namespace Lifecycle
' --------------------------------------

' --------------------------------------
' namespace API {
' --------------------------------------
Public Function fn_Register(ByVal bindings As obj_UiBindingContext) As Boolean
    fn_Register = private_Execute(bindings, False)
End Function

Public Function fn_ReplaceId(ByVal bindings As obj_UiBindingContext) As Boolean
    fn_ReplaceId = private_Execute(bindings, True)
End Function
' --------------------------------------
' } // namespace API
' --------------------------------------

Private Function private_Execute( _
    ByVal bindings As obj_UiBindingContext, _
    ByVal replaceMode As Boolean _
) As Boolean
    Dim runtimeContext As Object
    Dim historyTable As ListObject
    Dim history As Object
    Dim values As Variant
    Dim diagnostic As String
    Dim previousEvents As Boolean
    Dim previousUpdating As Boolean
    Dim runtimeEntered As Boolean
    Dim startedAt As Double
    Dim stageStartedAt As Double

    startedAt = Timer
    If m_busy Then
        ex_WindowsUi.fn_ShowMessage private_Cfg("Message.RuntimeBusy"), vbExclamation
        Exit Function
    End If
    If Not ex_RuntimeLifecycle.fn_TryEnter(runtimeContext) Then
        Exit Function
    End If
    runtimeEntered = True
    m_busy = True
    previousEvents = Application.EnableEvents
    previousUpdating = Application.ScreenUpdating
    ReDim values(1 To 1, 1 To 7)
    On Error GoTo Failed
    Application.EnableEvents = False
    Application.ScreenUpdating = False
    If Not bindings.SetValue("Form", "Result", vbNullString) Then
        Err.Raise 5, , Replace(private_Cfg("Message.BindingUnavailable"), "{name}", "Form.Result")
    End If

    ' Resolve the writable database in this Excel instance; never open it implicitly.
    Set m_database = private_RequireDatabase()
    Set historyTable = private_Table("History")
    If replaceMode Then
        values(1, 4) = private_ReadField(bindings, "PreviousId")
        values(1, 3) = private_ReadField(bindings, "NewId")
        If Len(private_Normalize(private_InputText(values, 4))) = 0 Or _
                Len(private_Normalize(private_InputText(values, 3))) = 0 Then
            Err.Raise 5, , private_Cfg("Message.ReplaceRequired")
        End If
        private_ValidateNewID private_Normalize(private_InputText(values, 3))
    Else
        values(1, 1) = private_ReadField(bindings, "PersonName")
        values(1, 2) = private_ReadField(bindings, "Origin")
        If Len(private_Normalize(private_InputText(values, 1))) = 0 Then
            Err.Raise 5, , private_Cfg("Message.NameRequired")
        End If
        If Len(private_Normalize(private_InputText(values, 2))) = 0 Then
            Err.Raise 5, , private_Cfg("Message.OriginRequired")
        End If
    End If

    ' Read current unsaved database edits; this database is the sole identity authority.
    stageStartedAt = Timer
    Set history = private_LoadHistory(historyTable)
    ex_Core.fn_Diagnostic_WritePerf "PR.Registry.ReadHistory", stageStartedAt
    stageStartedAt = Timer
    Set m_originAliases = private_LoadOriginAliases()
    ex_Core.fn_Diagnostic_WritePerf "PR.Registry.ReadAliases", stageStartedAt
    stageStartedAt = Timer
    If replaceMode Then
        private_Replace values, historyTable, history
    Else
        private_Register values, historyTable, history
    End If
    ex_Core.fn_Diagnostic_WritePerf "PR.Registry.ProcessAndWrite", stageStartedAt
    If CStr(values(1, 6)) = private_Cfg("Status.Error") Then
        ex_WindowsUi.fn_ShowMessage CStr(values(1, 7)), vbExclamation
    End If
    private_Execute = CStr(values(1, 6)) <> private_Cfg("Status.Error") And _
        CStr(values(1, 6)) <> private_Cfg("Status.Cancelled")
Report:
    On Error GoTo ReportFailed
    private_WriteResult bindings, values, replaceMode
Cleanup:
    Set m_database = Nothing
    Set m_originAliases = Nothing
    Application.EnableEvents = previousEvents
    Application.ScreenUpdating = previousUpdating
    m_busy = False
    ex_Core.fn_Diagnostic_WritePerf "PR.Registry.Total", startedAt
    If runtimeEntered Then
        ex_RuntimeLifecycle.fn_Leave runtimeContext
    End If
    Exit Function
Failed:
    diagnostic = Err.Description
    Resume HandleFailure
HandleFailure:
    On Error GoTo ReportFailed
    private_Execute = False
    values(1, 5) = vbNullString
    values(1, 6) = private_Cfg("Status.Error")
    values(1, 7) = diagnostic
    ex_WindowsUi.fn_ShowMessage diagnostic, vbExclamation
    GoTo Report
ReportFailed:
    ex_WindowsUi.fn_ShowMessage Err.Description, vbExclamation
    Resume Cleanup
End Function

Private Sub private_Register( _
    ByRef values As Variant, _
    ByVal historyTable As ListObject, _
    ByVal history As Object _
)
    Dim fullName As String
    Dim origin As String
    Dim originalCount As Long

    fullName = private_Normalize(private_InputText(values, 1))
    origin = private_Normalize(private_InputText(values, 2))
    If StrComp(fullName, m_savedName, vbTextCompare) = 0 And _
            StrComp(origin, m_savedOrigin, vbTextCompare) = 0 Then
        If history.Exists(m_savedID) Then
            values(1, 4) = m_savedID
        End If
    End If
    originalCount = history.Count
    If Not private_ProcessRow(values, history) Then
        Exit Sub
    End If
    If history.Count <> originalCount Then
        private_WriteHistory historyTable, history, CStr(values(1, 5))
    End If
    m_savedID = CStr(values(1, 5))
    m_savedName = fullName
    m_savedOrigin = origin
End Sub

Private Sub private_Replace( _
    ByRef values As Variant, _
    ByVal historyTable As ListObject, _
    ByVal history As Object _
)
    Dim oldID As String
    Dim newID As String
    Dim currentID As String
    Dim fullName As String
    Dim origin As String
    Dim record As Variant
    Dim message As String

    oldID = private_Normalize(private_InputText(values, 4))
    newID = private_Normalize(private_InputText(values, 3))
    If history.Exists(oldID) Then
        record = history(oldID)
        currentID = CStr(record(0))
        record = history(currentID)
        fullName = CStr(record(1))
        origin = CStr(record(2))
    Else
        Err.Raise 5, , private_Cfg("Message.UnknownID")
    End If
    If StrComp(newID, currentID, vbTextCompare) = 0 Then
        values(1, 5) = currentID
        values(1, 6) = private_Cfg("Status.Existing")
        values(1, 7) = private_Cfg("Message.SameID")
        Exit Sub
    End If
    If history.Exists(newID) Then
        Err.Raise 5, , private_Cfg("Message.IDInUse")
    End If
    message = Replace(private_Cfg("Message.ConfirmReplacement"), "{name}", fullName)
    message = Replace(message, "{old}", currentID)
    message = Replace(message, "{new}", newID)
    message = Replace(message, "{newline}", vbLf)
    If ex_WindowsUi.fn_ShowMessage(message, vbYesNo Or vbExclamation Or vbDefaultButton2, _
            private_Cfg("Title.Replace")) <> vbYes Then
        values(1, 6) = private_Cfg("Status.Cancelled")
        values(1, 7) = private_Cfg("Message.ReplaceCancelled")
        Exit Sub
    End If
    values(1, 1) = fullName
    values(1, 2) = origin
    If Not private_ProcessRow(values, history) Then
        Exit Sub
    End If
    private_WriteHistory historyTable, history
    If StrComp(m_savedID, currentID, vbTextCompare) = 0 Then
        m_savedID = CStr(values(1, 5))
    End If
End Sub

Private Function private_RequireDatabase() As Workbook
    Dim candidate As Workbook
    Dim fileSystem As Object
    Dim path As String

    Set fileSystem = CreateObject("Scripting.FileSystemObject")
    path = private_Cfg("registry.path")
    If fileSystem.GetDriveName(path) = vbNullString And Left$(path, 2) <> "\\" Then
        path = ThisWorkbook.Path & "\" & path
    End If
    path = fileSystem.GetAbsolutePathName(path)
    For Each candidate In Application.Workbooks
        If StrComp(candidate.FullName, path, vbTextCompare) = 0 Then
            If candidate.ReadOnly Then
                Err.Raise 5, , Replace(private_Cfg("Message.BaseReadOnly"), "{path}", path)
            End If
            Set private_RequireDatabase = candidate
            Exit Function
        End If
    Next candidate
    Err.Raise 5, , Replace(private_Cfg("Message.BaseOpenRequired"), "{path}", path)
End Function

Private Function private_Table(ByVal role As String) As ListObject
    Dim table As ListObject
    Dim headers As Variant
    Dim column As Long

    On Error GoTo Missing
    Set table = m_database.Worksheets(private_Cfg("registry." & LCase$(role) & ".sheet")) _
        .ListObjects(private_Cfg("registry." & LCase$(role) & ".table"))
    On Error GoTo 0
    headers = Split(private_Cfg("Header." & role), "|")
    If table.ListColumns.Count <> UBound(headers) + 1 Then
        Err.Raise 5, , private_Cfg("Message.SchemaInvalid")
    End If
    For column = 1 To table.ListColumns.Count
        If StrComp(table.ListColumns(column).Name, CStr(headers(column - 1)), vbBinaryCompare) <> 0 Then
            Err.Raise 5, , private_Cfg("Message.SchemaInvalid")
        End If
    Next column
    If role = "History" And table.Parent.ProtectContents Then
        Err.Raise 5, , private_Cfg("Message.BaseProtected")
    End If
    Set private_Table = table
    Exit Function
Missing:
    Err.Raise 5, , private_Cfg("Message.SchemaInvalid")
End Function

Private Sub private_WriteHistory( _
    ByVal table As ListObject, _
    ByVal history As Object, _
    Optional ByVal registrationID As String = vbNullString _
)
    Dim original As Variant
    Dim values As Variant
    Dim record As Variant
    Dim key As Variant
    Dim oldCount As Long
    Dim row As Long
    Dim column As Long
    Dim diagnostic As String
    Dim rollbackDiagnostic As String
    Dim growth As Range
    Dim cell As Range
    Dim mutationStarted As Boolean
    Dim appendOnly As Boolean
    Dim stageStartedAt As Double

    stageStartedAt = Timer
    oldCount = table.ListRows.Count
    appendOnly = Len(registrationID) > 0 And history.Count = oldCount + 1
    If oldCount > 0 And Not appendOnly Then
        original = table.DataBodyRange.Value2
    End If
    On Error GoTo Failed
    If history.Count > oldCount Then
        Set growth = table.HeaderRowRange.Offset(oldCount + 1).Resize(history.Count - oldCount, 5)
        For Each cell In growth.Cells
            If Not IsEmpty(cell.Value2) Or cell.HasFormula Then
                Err.Raise 5, , private_Cfg("Message.BaseNoSpace")
            End If
        Next cell
    End If
    If appendOnly Then
        ReDim values(1 To 1, 1 To 5)
        record = history(registrationID)
        values(1, 1) = registrationID
        For column = 0 To 3
            values(1, column + 2) = CStr(record(column))
        Next column
    Else
        ReDim values(1 To history.Count, 1 To 5)
        row = 0
        For Each key In history.Keys
            row = row + 1
            record = history(key)
            values(row, 1) = CStr(key)
            For column = 0 To 3
                values(row, column + 2) = CStr(record(column))
            Next column
        Next key
    End If

    ' Update the open workbook only; the user controls when it is saved to disk.
    mutationStarted = True
    table.Resize table.HeaderRowRange.Resize(history.Count + 1, 5)
    If Not growth Is Nothing Then
        growth.NumberFormat = "@"
    End If
    If appendOnly Then
        growth.Value2 = values
    Else
        table.HeaderRowRange.Offset(1).Resize(history.Count, 5).Value2 = values
    End If
    ex_Core.fn_Diagnostic_WritePerf "PR.Registry.WriteHistory", stageStartedAt
    Exit Sub
Failed:
    diagnostic = Err.Description
    If mutationStarted Then
        On Error Resume Next
        Err.Clear
        If appendOnly Then
            growth.ClearContents
        Else
            table.HeaderRowRange.Offset(1).Resize(history.Count, 5).ClearContents
        End If
        table.Resize table.HeaderRowRange.Resize(oldCount + 1, 5)
        If oldCount > 0 And Not appendOnly Then
            table.HeaderRowRange.Offset(1).Resize(oldCount, 5).Value2 = original
        End If
        If Err.Number <> 0 Then
            rollbackDiagnostic = Err.Description
        End If
        On Error GoTo 0
    End If
    If Len(rollbackDiagnostic) > 0 Then
        diagnostic = diagnostic & vbLf & rollbackDiagnostic
    End If
    Err.Raise 5, "private_WriteHistory", Replace(private_Cfg("Message.WriteFailed"), "{error}", diagnostic)
End Sub

Private Function private_ReadField( _
    ByVal bindings As obj_UiBindingContext, _
    ByVal name As String _
) As Variant
    Dim value As Variant
    Dim objectValue As Object
    Dim isObject As Boolean

    If Not bindings.TryGetValue("Form", name, value, objectValue, isObject) Then
        Err.Raise 5, , Replace(private_Cfg("Message.BindingUnavailable"), "{name}", "Form." & name)
    End If
    If isObject Or IsError(value) Or IsNull(value) Then
        Err.Raise 5, , private_Cfg("Message.CellError")
    End If
    private_ReadField = value
End Function

Private Sub private_WriteResult( _
    ByVal bindings As obj_UiBindingContext, _
    ByRef values As Variant, _
    ByVal replaceMode As Boolean _
)
    Dim text As String
    Dim updates As Object
    Dim diagnostic As String

    text = Replace(private_Cfg("Form.ResultTemplate"), "{id}", CStr(values(1, 5)))
    text = Replace(text, "{status}", CStr(values(1, 6)))
    text = Replace(text, "{message}", CStr(values(1, 7)))
    Set updates = private_Dictionary()
    updates.Add "Form.Result", Replace(text, "{newline}", vbLf)
    If replaceMode And CStr(values(1, 7)) = private_Cfg("Message.Migrated") Then
        updates.Add "Form.PreviousId", CStr(values(1, 5))
        updates.Add "Form.NewId", vbNullString
    End If
    If Not bindings.TryApplyValues(updates, diagnostic) Then
        Err.Raise 5, , diagnostic
    End If
End Sub

Private Function private_ProcessRow( _
    ByRef inputRow As Variant, _
    ByVal history As Object _
) As Boolean
    Dim fullName As String
    Dim origin As String
    Dim newID As String
    Dim existingID As String
    Dim candidates As String
    Dim selectedID As String
    Dim resultID As String
    Dim message As String
    Dim record As Variant
    Dim key As Variant
    Dim isNew As Boolean
    Dim isMigration As Boolean
    Dim mutationStarted As Boolean

    On Error GoTo Failed
    fullName = private_Normalize(private_InputText(inputRow, 1))
    origin = private_Normalize(private_InputText(inputRow, 2))
    newID = private_Normalize(private_InputText(inputRow, 3))
    existingID = private_Normalize(private_InputText(inputRow, 4))
    If Len(fullName) = 0 Then
        Err.Raise vbObjectError + 2001, , private_Cfg("Message.NameRequired")
    End If
    If Len(newID) > 0 Then
        private_ValidateNewID newID
    End If

    ' Resolve explicit identifiers; matching names alone never merge people.
    If Len(existingID) > 0 Then
        If history.Exists(existingID) Then
            record = history(existingID)
            private_CheckName CStr(record(1)), fullName, existingID
            selectedID = CStr(record(0))
        Else
            Err.Raise vbObjectError + 2005, , private_Cfg("Message.UnknownID")
        End If
    End If
    If Len(newID) > 0 Then
        If history.Exists(newID) Then
            record = history(newID)
            private_CheckName CStr(record(1)), fullName, newID
            If Len(selectedID) > 0 And StrComp(selectedID, CStr(record(0)), vbTextCompare) <> 0 Then
                Err.Raise vbObjectError + 2007, , private_Cfg("Message.IDConflict")
            End If
            selectedID = CStr(record(0))
        End If
    End If
    isNew = Len(selectedID) = 0

    ' A blank new ID generates for a new person and resolves an existing person's current ID.
    If Len(selectedID) = 0 Then
        If Len(newID) > 0 Then
            resultID = newID
        Else
            If Len(origin) = 0 Then
                Err.Raise vbObjectError + 2010, , private_Cfg("Message.OriginRequired")
            End If
            resultID = private_UniqueID(private_BaseID(fullName, origin), history)
        End If
    Else
        resultID = selectedID
        If Len(newID) > 0 And StrComp(newID, selectedID, vbTextCompare) <> 0 Then
            If Not history.Exists(newID) Then
                resultID = newID
                isMigration = True
            End If
        End If
    End If

    If Len(origin) = 0 And Len(selectedID) > 0 Then
        If history.Exists(selectedID) Then
            record = history(selectedID)
            origin = CStr(record(2))
        End If
    End If

    ' Confirm same-name creation before mutating history.
    If isNew Then
        candidates = private_Candidates(fullName, history)
        If Len(candidates) > 0 Then
            message = Replace(private_Cfg("Message.ConfirmNewPerson"), "{candidates}", candidates)
            If ex_WindowsUi.fn_ShowMessage(message, vbYesNo Or vbExclamation Or vbDefaultButton2, _
                private_Cfg("Title.Input")) <> vbYes Then
                inputRow(1, 5) = vbNullString
                inputRow(1, 6) = private_Cfg("Status.Cancelled")
                inputRow(1, 7) = private_Cfg("Message.Cancelled")
                Exit Function
            End If
        End If
    End If

    ' Flatten replacements directly to the current ID; no internal person keys or chains.
    If isMigration Then
        mutationStarted = True
        For Each key In history.Keys
            record = history(key)
            If StrComp(CStr(record(0)), selectedID, vbTextCompare) = 0 Then
                record(0) = resultID
                history(key) = record
            End If
        Next key
        If Not history.Exists(selectedID) Then
            history.Add selectedID, Array(resultID, fullName, origin, private_Timestamp())
        End If
        history.Add resultID, Array(resultID, fullName, origin, private_Timestamp())
    ElseIf isNew Then
        mutationStarted = True
        history.Add resultID, Array(resultID, fullName, origin, private_Timestamp())
    End If

    inputRow(1, 5) = resultID
    If isMigration Then
        message = private_Cfg("Message.Migrated")
    ElseIf isNew Then
        message = private_Cfg("Message.Created")
    Else
        message = private_Cfg("Message.Found")
    End If
    If Not isNew And Not isMigration Then
        inputRow(1, 6) = private_Cfg("Status.Existing")
    Else
        inputRow(1, 6) = private_Cfg("Status.Applied")
    End If
    inputRow(1, 7) = message
    private_ProcessRow = True
    Exit Function
Failed:
    message = Err.Description
    If mutationStarted Then
        ' A failed mutation aborts before any history is persisted.
        Err.Raise vbObjectError + 2012, "private_ProcessRow", message
    End If
    inputRow(1, 5) = vbNullString
    inputRow(1, 6) = private_Cfg("Status.Error")
    inputRow(1, 7) = message
End Function

Private Function private_InputText( _
    ByRef values As Variant, _
    ByVal column As Long _
) As String
    If IsError(values(1, column)) Then
        Err.Raise vbObjectError + 2040, , private_Cfg("Message.CellError")
    End If
    private_InputText = CStr(values(1, column))
End Function

Private Sub private_ValidateNewID( _
    ByVal identifier As String _
)
    Dim index As Long
    Dim character As String
    Dim invalidMessage As String

    invalidMessage = Replace(private_Cfg("Message.InvalidID"), "{max}", private_Cfg("Limit.IDLength"))
    If Len(identifier) > CLng(private_Cfg("Limit.IDLength")) Then
        Err.Raise vbObjectError + 2061, , invalidMessage
    End If
    For index = 1 To Len(identifier)
        character = Mid$(identifier, index, 1)
        If AscW(character) >= 0 And AscW(character) < 32 Then
            Err.Raise vbObjectError + 2061, , invalidMessage
        End If
        If InStr(1, private_Cfg("Value.InvalidIDCharacters"), character, vbBinaryCompare) > 0 Then
            Err.Raise vbObjectError + 2061, , invalidMessage
        End If
    Next index
End Sub

Private Function private_LoadHistory(ByVal table As ListObject) As Object
    Dim history As Object
    Dim values As Variant
    Dim row As Long
    Dim column As Long
    Dim populated As Boolean
    Dim identifier As String
    Dim currentID As String
    Dim fullName As String
    Dim key As Variant
    Dim record As Variant
    Dim current As Variant

    Set history = private_Dictionary()
    If table.ListColumns.Count <> 5 Then
        Err.Raise vbObjectError + 2052, , private_Cfg("Message.SchemaInvalid")
    End If
    If Not table.DataBodyRange Is Nothing Then
        values = table.DataBodyRange.Value2
        For row = 1 To UBound(values, 1)
            populated = False
            For column = 1 To 5
                If IsError(values(row, column)) Then
                    Err.Raise vbObjectError + 2040, , private_Cfg("Message.CellError")
                End If
                If Not IsEmpty(values(row, column)) Then
                    populated = True
                End If
            Next column
            If populated Then
                identifier = private_Normalize(CStr(values(row, 1)))
                currentID = private_Normalize(CStr(values(row, 2)))
                fullName = private_Normalize(CStr(values(row, 3)))
                If Len(identifier) = 0 Or Len(currentID) = 0 Or Len(fullName) = 0 Or history.Exists(identifier) Then
                    Err.Raise vbObjectError + 2020, , private_Cfg("Message.RegistryInvalid")
                End If
                history.Add identifier, Array(currentID, fullName, CStr(values(row, 4)), CStr(values(row, 5)))
            End If
        Next row
    End If
    For Each key In history.Keys
        record = history(key)
        If Not history.Exists(CStr(record(0))) Then
            Err.Raise vbObjectError + 2021, , private_Cfg("Message.RegistryInvalid")
        End If
        current = history(CStr(record(0)))
        If StrComp(CStr(current(0)), CStr(record(0)), vbTextCompare) <> 0 Or _
                StrComp(CStr(current(1)), CStr(record(1)), vbTextCompare) <> 0 Then
            Err.Raise vbObjectError + 2022, , private_Cfg("Message.RegistryInvalid")
        End If
    Next key
    Set private_LoadHistory = history
End Function

Private Sub private_CheckName( _
    ByVal expected As String, _
    ByVal actual As String, _
    ByVal identifier As String _
)
    If StrComp(private_Normalize(expected), actual, vbTextCompare) <> 0 Then
        Err.Raise vbObjectError + 2104, , Replace(private_Cfg("Message.IDMismatch"), "{id}", identifier)
    End If
End Sub

Private Function private_Candidates( _
    ByVal fullName As String, _
    ByVal history As Object _
) As String
    Dim candidates As Object
    Dim key As Variant
    Dim record As Variant
    Dim count As Long
    Dim result As String

    Set candidates = private_Dictionary()
    For Each key In history.Keys
        record = history(key)
        If StrComp(CStr(record(1)), fullName, vbTextCompare) = 0 Then
            candidates(CStr(record(0))) = CStr(record(2))
        End If
    Next key
    For Each key In candidates.Keys
        count = count + 1
        If count > CLng(private_Cfg("Limit.Candidates")) Then
            Exit For
        End If
        result = result & CStr(key) & " [" & CStr(candidates(key)) & "]; "
    Next key
    private_Candidates = result
End Function

Private Function private_Timestamp() As String
    private_Timestamp = Format$(Now, private_Cfg("Format.Timestamp"))
End Function

Private Function private_BaseID( _
    ByVal fullName As String, _
    ByVal origin As String) As String

    Dim words As Variant
    Dim result As String
    Dim index As Long
    Dim originToken As String

    words = Split(fullName, " ")
    result = private_Token(CStr(words(0)))
    result = UCase$(Left$(result, 1)) & LCase$(Mid$(result, 2))
    For index = 1 To UBound(words)
        result = result & UCase$(Left$(private_Token(CStr(words(index))), 1))
    Next index
    originToken = private_Token(origin)
    If m_originAliases.Exists(private_OriginKey(origin)) Then
        originToken = CStr(m_originAliases(private_OriginKey(origin)))
    End If
    result = result & private_Cfg("Value.Separator") & originToken
    If Len(private_Token(CStr(words(0)))) = 0 Or Len(originToken) = 0 Then
        Err.Raise vbObjectError + 2030, , private_Cfg("Message.EmptyToken")
    End If
    private_BaseID = Left$(result, CLng(private_Cfg("Limit.IDLength")))
End Function

Private Function private_OriginKey(ByVal value As String) As String
    value = Replace(private_Normalize(value), ChrW(&H2019), "'")
    value = Replace(value, ChrW(&H2BC), "'")
    private_OriginKey = private_Token(value)
End Function

Private Function private_LoadOriginAliases() As Object
    Dim aliases As Object
    Dim table As ListObject
    Dim values As Variant
    Dim row As Long
    Dim messageTemplate As String
    Dim maxLength As Long
    Dim headerRow As Long
    Dim fullName As String
    Dim shortCode As String
    Dim key As String
    Dim message As String

    Set aliases = private_Dictionary()
    Set table = private_Table("Aliases")
    If table.ListColumns.Count <> 2 Then
        Err.Raise vbObjectError + 2052, , private_Cfg("Message.SchemaInvalid")
    End If
    messageTemplate = private_Cfg("Message.AliasInvalid")
    maxLength = CLng(private_Cfg("Limit.AliasLength"))
    headerRow = table.HeaderRowRange.Row
    If Not table.DataBodyRange Is Nothing Then
        values = table.DataBodyRange.Value2
        For row = 1 To UBound(values, 1)
            If IsError(values(row, 1)) Or IsError(values(row, 2)) Then
                Err.Raise vbObjectError + 2040, , private_Cfg("Message.CellError")
            End If
            If Not IsEmpty(values(row, 1)) Or Not IsEmpty(values(row, 2)) Then
                fullName = private_Normalize(CStr(values(row, 1)))
                shortCode = private_Normalize(CStr(values(row, 2)))
                key = private_OriginKey(fullName)
                message = Replace(messageTemplate, "{row}", CStr(headerRow + row))
                If Len(key) = 0 Or Len(shortCode) = 0 Or _
                        Len(shortCode) > maxLength Or _
                        StrComp(private_Token(shortCode), shortCode, vbBinaryCompare) <> 0 Then
                    Err.Raise vbObjectError + 2063, , message
                End If
                If aliases.Exists(key) Then
                    If StrComp(CStr(aliases(key)), shortCode, vbTextCompare) <> 0 Then
                        Err.Raise vbObjectError + 2063, , message
                    End If
                Else
                    aliases.Add key, shortCode
                End If
            End If
        Next row
    End If
    Set private_LoadOriginAliases = aliases
End Function

Private Function private_UniqueID( _
    ByVal baseID As String, _
    ByVal history As Object) As String

    Dim candidate As String
    Dim suffix As String
    Dim counter As Long
    Dim maxLength As Long

    candidate = baseID
    counter = 1
    maxLength = CLng(private_Cfg("Limit.IDLength"))
    Do While history.Exists(candidate)
        counter = counter + 1
        suffix = private_Cfg("Value.Separator") & CStr(counter)
        candidate = Left$(baseID, maxLength - Len(suffix)) & suffix
    Loop
    private_UniqueID = candidate
End Function

Private Function private_Normalize(ByVal value As String) As String
    value = Replace(value, ChrW(160), " ")
    value = Replace(value, ChrW(&H202F), " ")
    value = Replace(value, vbTab, " ")
    value = Replace(value, vbCr, " ")
    value = Replace(value, vbLf, " ")
    value = Trim$(value)
    Do While InStr(value, "  ") > 0
        value = Replace(value, "  ", " ")
    Loop
    private_Normalize = value
End Function

Private Function private_Token(ByVal value As String) As String
    Dim index As Long
    Dim character As String
    Dim result As String
    Dim removedCharacters As String

    removedCharacters = private_Cfg("Value.RemovedCharacters")
    For index = 1 To Len(value)
        character = Mid$(value, index, 1)
        If InStr(1, removedCharacters, character, vbBinaryCompare) = 0 Then
            result = result & character
        End If
    Next index
    private_Token = result
End Function

Private Function private_Dictionary() As Object
    Dim dictionary As Object

    Set dictionary = CreateObject("Scripting.Dictionary")
    dictionary.CompareMode = vbTextCompare
    Set private_Dictionary = dictionary
End Function

Private Function private_Cfg(ByVal key As String) As String
    private_Cfg = ex_PRM.fn_Config(key)
End Function