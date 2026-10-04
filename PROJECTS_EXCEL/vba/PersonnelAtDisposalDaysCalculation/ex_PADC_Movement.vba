Attribute VB_Name = "ex_PADC_Movement"
Option Explicit

Private SOURCE_SHEET_NAME As String
Private SOURCE_TABLE_NAME As String
Private SOURCE_COL_NAME As String
Private ROSTER_SHEET_NAME As String
Private ROSTER_TABLE_NAME As String
Private ROSTER_COL_NAME As String
Private ROSTER_COL_TAX_ID As String
Private MSG_PERSON_UNVERIFIED As String
Private MSG_AMBIGUOUS_NAME As String
Private SOURCE_COL_TAX_ID As String
Private SOURCE_COL_EVENT As String
Private SOURCE_COL_PERIOD_FROM As String
Private SOURCE_COL_PERIOD_TO As String

Private TARGET_COL_NAME As String
Private TARGET_COL_TAX_ID As String
Private TARGET_COL_START As String
Private TARGET_COL_DAYS As String
Private TARGET_COL_THRESHOLD_DATE As String
Private COUNTED_DAYS_THRESHOLD As Long
Private ERROR_COL_NAME As String
Private ERROR_COL_TAX_ID As String
Private ERROR_COL_DESCRIPTION As String
Private MSG_SKIPPED_PEOPLE As String
Private TARGET_COL_TRIPS As String
Private TARGET_COL_PERIODS As String

Private Const INPUT_REFERENCE_SEPARATOR As String = "|"
Private Const INPUT_SHEET_SEPARATOR As String = "!"
Private Const PARAM_INPUT_PATH_CELL As String = "I2"

Private PARAM_COL_NAME As String
Private PARAM_COL_TAX_ID As String
Private PARAM_COL_START As String
Private PARAM_COL_TRIPS As String

Private EXCLUDED_EVENTS As String

Private OPTIONAL_EXCLUDED_EVENTS As String

Private PRESENT_EVENT_NAME As String
Private m_includeTripValues As String
Private m_excludeTripValues As String

Private Const UI_YIELD_INTERVAL As Long = 100
Private Const UI_YIELD_SECONDS As Single = 0.1

Private Const FORMAT_DATE As String = "dd.mm.yyyy"
Private Const DATE_TEXT_PATTERN As String = "##.##.####"
Private Const DATE_1904_OFFSET As Long = 1462
Private Const MIN_DATE_SERIAL As Long = 61
Private Const MAX_DATE_SERIAL As Long = 2958465
Private Const NBSP_CODE As Long = 160

Private MSG_ALREADY_RUNNING As String
Private MSG_PARAMETER_LOADING_STARTED As String
Private MSG_SELECT_TOOL_SHEET As String
Private MSG_WRONG_WORKBOOK As String
Private MSG_PARAMETER_WORKSHEET As String
Private MSG_PARAMETER_LOADING_COMPLETED As String
Private MSG_INPUT_FILE As String
Private MSG_END_DATE As String
Private MSG_OPERATION_CANCELLED As String
Private MSG_OPERATION_STOPPED As String
Private MSG_CANCELLING_OPERATION As String
Private MSG_RUKH_SOURCE As String
Private MSG_PARAMETER_CLOSE_FAILED As String
Private MSG_PARAMETER_SHEET_MISSING As String
Private MSG_PARAMETERS_EMPTY As String
Private MSG_PEOPLE As String
Private MSG_RUKH_EVENTS As String
Private MSG_WORKSHEET_ROW As String
Private MSG_TAX_ID As String
Private MSG_DUPLICATE_TAX_ID As String
Private MSG_TIME_POINT As String
Private MSG_START_AFTER_END As String
Private MSG_FULL_NAME As String
Private MSG_EVENT As String
Private MSG_DEPARTURE As String
Private MSG_ARRIVAL_CELL_ERROR As String
Private MSG_ARRIVAL As String
Private MSG_ARRIVAL_BEFORE_DEPARTURE As String
Private MSG_WRITING_RESULT As String
Private MSG_ROWS As String
Private MSG_TARGET_PREPARE_FAILED As String
Private MSG_WRITING_RESULTS As String
Private MSG_CALCULATION_COMPLETED_PEOPLE As String
Private MSG_RELATIVE_PATH_BASE As String
Private MSG_AMBIGUOUS_PATH As String
Private MSG_FILE_NOT_FOUND_OR_INACCESSIBLE As String
Private MSG_CANCELLED As String
Private MSG_CALCULATING_DAYS As String
Private MSG_MULTIPLE_SOURCES As String
Private MSG_OPEN_SOURCE As String
Private MSG_TABLE As String
Private MSG_TABLE_NOT_FOUND As String
Private MSG_ON_WORKSHEET As String
Private MSG_CHECK_TABLE_CONFIGURATION As String
Private MSG_TABLE_TEXT As String
Private MSG_IS_MISSING_COLUMN As String
Private MSG_CELL_ERROR As String
Private MSG_REQUIRED_VALUE As String
Private MSG_SUBTRACT_BUSINESS_TRIPS As String
Private MSG_INVALID_TRIP_FLAG As String
Private MSG_INVALID_DATE As String
Private MSG_INVALID_REFERENCE As String
Private MSG_PARAMETER_TABLE_MISSING As String

Private Const CANCEL_ERROR As Long = vbObjectError + 2101
Private Const DICTIONARY_PROG_ID As String = "Scripting.Dictionary"
Private Const FILE_SYSTEM_PROG_ID As String = "Scripting.FileSystemObject"
Private Const ERROR_SOURCE As String = "PersonnelAtDisposalDays"
Private Const OPERATION_ERROR As Long = vbObjectError + 2100

Private m_pageBase As obj_PageBase
Private mRunning As Boolean
Private mCancel As Boolean
Private mLastUiYield As Single

Private Type CalculationParameters
    InputPath As String
    SheetName As String
    HeaderAddress As String
    EndDay As Long
End Type

' --------------------------------------
' namespace API {
' --------------------------------------
Public Sub CalculateMovementDays(ByVal pageBase As obj_PageBase)
    Dim parameterSheet As Worksheet
    Dim calculationParameters As CalculationParameters
    Dim parameterWorkbook As Workbook
    Dim params As ListObject
    Dim openedHere As Boolean
    Dim oldStatus As Variant
    Dim errorText As String
    Dim errorNumber As Long

    If mRunning Then
        VBA.MsgBox MSG_ALREADY_RUNNING, vbExclamation
        Exit Sub
    End If
    oldStatus = Application.StatusBar
    On Error GoTo Failed
    mRunning = True
    mCancel = False
    Set m_pageBase = pageBase
    If Not private_LoadConfiguration() Then
        GoTo Cleanup
    End If
    private_ClearLog
    private_LogDebug MSG_PARAMETER_LOADING_STARTED
    If Not TypeOf Application.ActiveSheet Is Worksheet Then
        private_Fail MSG_SELECT_TOOL_SHEET
    End If
    Set parameterSheet = Application.ActiveSheet
    If Not parameterSheet.Parent Is ThisWorkbook Then
        private_Fail MSG_WRONG_WORKBOOK
    End If
    private_CheckCancel 0
    private_LogDebug MSG_PARAMETER_WORKSHEET & parameterSheet.Name
    private_ReadCalculationParameters parameterSheet, calculationParameters
    private_CheckCancel 0
    private_LogDebug MSG_PARAMETER_LOADING_COMPLETED
    Set parameterWorkbook = private_OpenParameterWorkbook(calculationParameters.InputPath, openedHere)
    private_CheckCancel 0
    Set params = private_ReadParameterTable(parameterWorkbook, calculationParameters)
    private_BuildMovementDays params, calculationParameters.EndDay
Cleanup:
    If openedHere Then
        private_CloseParameterWorkbook parameterWorkbook
    End If
    Application.StatusBar = oldStatus
    mRunning = False
    mCancel = False
    Set m_pageBase = Nothing
    Exit Sub
Failed:
    errorText = VBA.Err.Description
    errorNumber = VBA.Err.Number
    private_LogFailure errorNumber, errorText
    If errorNumber = CANCEL_ERROR Then
        VBA.MsgBox MSG_OPERATION_CANCELLED, vbInformation
    Else
        VBA.MsgBox MSG_OPERATION_STOPPED & errorText, vbExclamation
    End If
    Resume Cleanup
End Sub

Public Sub CancelMovementDays()
    If mRunning Then
        mCancel = True
        Application.StatusBar = MSG_CANCELLING_OPERATION
    End If
End Sub

' --------------------------------------
' } // namespace API
' --------------------------------------

Private Sub private_BuildMovementDays( _
    ByVal params As ListObject, _
    ByVal lastDay As Long _
)
    Dim source As ListObject
    Dim roster As ListObject
    Dim calculationService As obj_PADC_CalculationService
    Dim people As Object
    Dim intervals As Object
    Dim eventTaxIndex As Object
    Dim eventNameIndex As Object
    Dim eventNameIds As Object
    Dim rosterTaxIndex As Object
    Dim rosterNameIndex As Object
    Dim rosterNameIds As Object
    Dim rows As Collection
    Dim p As Variant
    Dim s As Variant
    Dim rosterData As Variant
    Dim matches As Collection
    Dim rosterMatches As Collection
    Dim rowNumber As Variant
    Dim sourceNameCol As Long
    Dim rosterTaxCol As Long
    Dim rosterNameCol As Long
    Dim matchError As String
    Dim resolvedTaxId As String
    Dim personErrors() As String
    Dim resultNames() As Variant
    Dim resultTaxIds() As Variant
    Dim resultStarts() As Variant
    Dim resultDays() As Variant
    Dim resultThresholdDates() As Variant
    Dim thresholdDate As Variant
    Dim resultPeriods() As Variant
    Dim periodsText As String
    Dim resultTrips() As Variant
    Dim starts() As Long
    Dim trips() As Boolean
    Dim taxCol As Long
    Dim eventCol As Long
    Dim fromCol As Long
    Dim toCol As Long
    Dim pTax As Long
    Dim pName As Long
    Dim pStart As Long
    Dim pTrip As Long
    Dim i As Long
    Dim person As Long
    Dim count As Long
    Dim parameterRows() As Long
    Dim failures As Collection
    Dim validCount As Long
    Dim firstDay As Long
    Dim arrival As Long
    Dim taxId As String
    Dim eventName As String
    Dim context As String

    Set source = private_FindSource()
    Set roster = private_RequireTable(ROSTER_SHEET_NAME, ROSTER_TABLE_NAME, source.Parent.Parent)
    rosterTaxCol = private_ColumnIndex(roster, ROSTER_COL_TAX_ID)
    rosterNameCol = private_ColumnIndex(roster, ROSTER_COL_NAME)
    rosterData = private_ReadQueryTable(roster)
    private_CheckCancel 0

    private_LogDebug MSG_RUKH_SOURCE & source.Parent.Parent.Name & " / " & source.Parent.Name & " / " & source.Name

    pTax = private_ColumnIndex(params, PARAM_COL_TAX_ID)
    pName = private_ColumnIndex(params, PARAM_COL_NAME)
    pStart = private_ColumnIndex(params, PARAM_COL_START)
    pTrip = private_ColumnIndex(params, PARAM_COL_TRIPS)
    sourceNameCol = private_ColumnIndex(source, SOURCE_COL_NAME)
    taxCol = private_ColumnIndex(source, SOURCE_COL_TAX_ID)
    eventCol = private_ColumnIndex(source, SOURCE_COL_EVENT)
    fromCol = private_ColumnIndex(source, SOURCE_COL_PERIOD_FROM)
    toCol = private_ColumnIndex(source, SOURCE_COL_PERIOD_TO)
    If params.DataBodyRange Is Nothing Then
        private_Fail MSG_PARAMETERS_EMPTY
    End If
    p = private_ReadQueryTable(params)
    s = private_ReadQueryTable(source)
    private_BuildPersonIndex s, source.ListRows.Count, taxCol, sourceNameCol, _
        eventTaxIndex, eventNameIndex, eventNameIds
    private_BuildPersonIndex rosterData, roster.ListRows.Count, rosterTaxCol, rosterNameCol, _
        rosterTaxIndex, rosterNameIndex, rosterNameIds
    count = private_CompactParameterRows(p, parameterRows)
    If count = 0 Then
        private_Fail MSG_PARAMETERS_EMPTY
    End If
    private_LogDebug MSG_PEOPLE & count & MSG_RUKH_EVENTS & source.ListRows.Count
    ReDim personErrors(1 To count)
    ReDim starts(1 To count)
    ReDim trips(1 To count)
    ReDim resultNames(1 To count, 1 To 1)
    ReDim resultTaxIds(1 To count, 1 To 1)
    ReDim resultStarts(1 To count, 1 To 1)
    ReDim resultDays(1 To count, 1 To 1)
    ReDim resultThresholdDates(1 To count, 1 To 1)
    ReDim resultPeriods(1 To count, 1 To 1)
    ReDim resultTrips(1 To count, 1 To 1)
    Set failures = New Collection
    Set people = VBA.CreateObject(DICTIONARY_PROG_ID)
    Set intervals = VBA.CreateObject(DICTIONARY_PROG_ID)
    For i = 1 To count
        context = params.Parent.Name & MSG_WORKSHEET_ROW & params.DataBodyRange.Row + parameterRows(i) - 1
        taxId = private_RequiredText(p(i, pTax), context & MSG_TAX_ID)
        If people.Exists(taxId) Then
            private_Fail context & MSG_DUPLICATE_TAX_ID & taxId
        End If
        people.Add taxId, i
        starts(i) = private_ReadDay(p(i, pStart), params.Parent.Parent.Date1904, context & MSG_TIME_POINT)
        If starts(i) > lastDay Then
            private_Fail context & MSG_START_AFTER_END
        End If
        trips(i) = private_ReadFlag(p(i, pTrip), context)
        resultTrips(i, 1) = p(i, pTrip)
        resultTaxIds(i, 1) = taxId
        resultNames(i, 1) = private_RequiredText(p(i, pName), context & MSG_FULL_NAME)
        resultStarts(i, 1) = VBA.CDate(starts(i))
        Set rows = New Collection
        intervals.Add taxId, rows
        private_CheckCancel i
    Next i
    For person = 1 To count
        private_CheckCancel 0
        taxId = resultTaxIds(person, 1)
        Set matches = private_FindPersonRows(eventTaxIndex, eventNameIndex, eventNameIds, _
            taxId, VBA.CStr(resultNames(person, 1)), matchError)
        If VBA.Len(matchError) > 0 Then
            personErrors(person) = matchError
            GoTo NextMatchedPerson
        End If
        If matches.Count = 0 Then
            Set rosterMatches = private_FindPersonRows(rosterTaxIndex, rosterNameIndex, rosterNameIds, _
                taxId, VBA.CStr(resultNames(person, 1)), matchError)
            If VBA.Len(matchError) > 0 Then
                personErrors(person) = matchError
                GoTo NextMatchedPerson
            End If
            If rosterMatches.Count = 0 Then
                personErrors(person) = MSG_PERSON_UNVERIFIED & resultNames(person, 1) & " / " & taxId
                GoTo NextMatchedPerson
            End If
            resolvedTaxId = private_MatchText(rosterData(rosterMatches(1), rosterTaxCol))
            If VBA.Len(resolvedTaxId) > 0 Then
                Set matches = private_FindPersonRows(eventTaxIndex, eventNameIndex, eventNameIds, _
                    resolvedTaxId, VBA.CStr(resultNames(person, 1)), matchError)
                If VBA.Len(matchError) > 0 Then
                    personErrors(person) = matchError
                    GoTo NextMatchedPerson
                End If
            End If
        End If
        For Each rowNumber In matches
            private_CheckCancel VBA.CLng(rowNumber)
            context = source.Parent.Parent.Name & " / " & SOURCE_SHEET_NAME & _
                MSG_WORKSHEET_ROW & source.DataBodyRange.Row + VBA.CLng(rowNumber) - 1
            eventName = private_RequiredText(s(VBA.CLng(rowNumber), eventCol), context & MSG_EVENT)
            firstDay = private_ReadDay(s(VBA.CLng(rowNumber), fromCol), source.Parent.Parent.Date1904, context & MSG_DEPARTURE)
            If VBA.IsError(s(VBA.CLng(rowNumber), toCol)) Then
                private_Fail context & MSG_ARRIVAL_CELL_ERROR
            End If
            If VBA.Len(VBA.Trim$(VBA.CStr(s(VBA.CLng(rowNumber), toCol)))) = 0 Then
                arrival = lastDay
            Else
                arrival = private_ReadDay(s(VBA.CLng(rowNumber), toCol), source.Parent.Parent.Date1904, context & MSG_ARRIVAL)
                If arrival < firstDay Then
                    private_Fail context & MSG_ARRIVAL_BEFORE_DEPARTURE
                End If
            End If
            If firstDay < starts(person) Then
                firstDay = starts(person)
            End If
            If arrival > lastDay Then
                arrival = lastDay
            End If
            If firstDay < arrival Then
                Set rows = intervals(VBA.CStr(resultTaxIds(person, 1)))
                rows.Add VBA.Array(firstDay, arrival, eventName, private_IsExcluded(eventName, Not trips(person)))
            End If
        Next rowNumber
NextMatchedPerson:
    Next person
    Set calculationService = New obj_PADC_CalculationService
    If Not calculationService.Initialize(COUNTED_DAYS_THRESHOLD, PRESENT_EVENT_NAME) Then
        private_Fail MSG_TARGET_PREPARE_FAILED
    End If
    For i = 1 To count
        private_CheckCancel i
        taxId = resultTaxIds(i, 1)
        If VBA.Len(personErrors(i)) > 0 Then
            failures.Add VBA.Array(resultNames(i, 1), taxId, personErrors(i))
            private_LogWarning personErrors(i)
            GoTo NextResultPerson
        End If
        Set rows = intervals(taxId)
        validCount = validCount + 1
        resultDays(validCount, 1) = calculationService.Calculate(rows, starts(i), lastDay, periodsText, thresholdDate, mCancel)
        resultPeriods(validCount, 1) = periodsText
        resultThresholdDates(validCount, 1) = thresholdDate
        resultNames(validCount, 1) = resultNames(i, 1)
        resultTaxIds(validCount, 1) = taxId
        resultStarts(validCount, 1) = resultStarts(i, 1)
        resultTrips(validCount, 1) = resultTrips(i, 1)
NextResultPerson:
    Next i
    calculationService.Dispose
    count = validCount
    private_CheckCancel 0
    private_LogDebug MSG_WRITING_RESULT & count & MSG_ROWS
    Application.StatusBar = MSG_WRITING_RESULTS
    private_PublishResults resultNames, resultTaxIds, resultStarts, resultTrips, resultPeriods, _
        resultDays, resultThresholdDates, count, failures
    private_LogDebug MSG_CALCULATION_COMPLETED_PEOPLE & count & MSG_SKIPPED_PEOPLE & failures.Count
    VBA.MsgBox MSG_CALCULATION_COMPLETED_PEOPLE & count & MSG_SKIPPED_PEOPLE & failures.Count, vbInformation
End Sub

Private Function private_FindPersonRows( _
    ByVal taxIndex As Object, _
    ByVal nameIndex As Object, _
    ByVal nameIds As Object, _
    ByVal taxId As String, _
    ByVal fullName As String, _
    ByRef matchError As String _
) As Collection
    Dim ids As Object

    matchError = vbNullString
    If taxIndex.Exists(taxId) Then
        Set private_FindPersonRows = taxIndex(taxId)
    ElseIf nameIndex.Exists(fullName) Then
        Set private_FindPersonRows = nameIndex(fullName)
        Set ids = nameIds(fullName)
        If ids.Count > 1 Then
            matchError = MSG_AMBIGUOUS_NAME & fullName
        End If
    Else
        Set private_FindPersonRows = New Collection
    End If
End Function

Private Sub private_BuildPersonIndex( _
    ByRef data As Variant, _
    ByVal count As Long, _
    ByVal taxColumn As Long, _
    ByVal nameColumn As Long, _
    ByRef taxIndex As Object, _
    ByRef nameIndex As Object, _
    ByRef nameIds As Object _
)
    Dim i As Long
    Dim taxId As String
    Dim fullName As String
    Dim rows As Collection
    Dim ids As Object

    Set taxIndex = VBA.CreateObject(DICTIONARY_PROG_ID)
    Set nameIndex = VBA.CreateObject(DICTIONARY_PROG_ID)
    Set nameIds = VBA.CreateObject(DICTIONARY_PROG_ID)
    nameIndex.CompareMode = vbTextCompare
    nameIds.CompareMode = vbTextCompare
    For i = 1 To count
        private_CheckCancel i
        taxId = private_MatchText(data(i, taxColumn))
        fullName = private_MatchText(data(i, nameColumn))
        If VBA.Len(taxId) > 0 Then
            If Not taxIndex.Exists(taxId) Then
                Set rows = New Collection
                taxIndex.Add taxId, rows
            End If
            Set rows = taxIndex(taxId)
            rows.Add i
        End If
        If VBA.Len(fullName) > 0 Then
            If Not nameIndex.Exists(fullName) Then
                Set rows = New Collection
                nameIndex.Add fullName, rows
                Set ids = VBA.CreateObject(DICTIONARY_PROG_ID)
                nameIds.Add fullName, ids
            End If
            Set rows = nameIndex(fullName)
            rows.Add i
            Set ids = nameIds(fullName)
            If VBA.Len(taxId) > 0 Then
                ids(taxId) = True
            End If
        End If
    Next i
End Sub

Private Function private_MatchText(ByVal value As Variant) As String
    If VBA.IsError(value) Or VBA.IsNull(value) Then
        Exit Function
    End If
    private_MatchText = VBA.Trim$(VBA.Replace(VBA.CStr(value), VBA.ChrW(NBSP_CODE), " "))
End Function

Private Function private_OpenParameterWorkbook( _
    ByVal filePath As String, _
    ByRef openedHere As Boolean _
) As Workbook
    Dim workbook As Workbook
    Dim oldSecurity As Long
    Dim errorNumber As Long
    Dim errorText As String

    For Each workbook In Application.Workbooks
        If VBA.StrComp(workbook.FullName, filePath, vbTextCompare) = 0 Then
            Set private_OpenParameterWorkbook = workbook
            Exit Function
        End If
    Next workbook
    oldSecurity = Application.AutomationSecurity
    On Error GoTo Failed
    Application.AutomationSecurity = msoAutomationSecurityForceDisable
    Set private_OpenParameterWorkbook = Application.Workbooks.Open(Filename:=filePath, _
        UpdateLinks:=0, ReadOnly:=True, AddToMru:=False, IgnoreReadOnlyRecommended:=True)
    openedHere = True
    Application.AutomationSecurity = oldSecurity
    Exit Function
Failed:
    errorNumber = VBA.Err.Number
    errorText = VBA.Err.Description
    Application.AutomationSecurity = oldSecurity
    VBA.Err.Raise errorNumber, ERROR_SOURCE, errorText
End Function

Private Sub private_CloseParameterWorkbook(ByVal workbook As Workbook)
    On Error GoTo Failed
    workbook.Close SaveChanges:=False
    Exit Sub
Failed:
    VBA.MsgBox MSG_PARAMETER_CLOSE_FAILED & VBA.Err.Description, vbExclamation
End Sub

Private Function private_ReadParameterTable( _
    ByVal workbook As Workbook, _
    ByRef parameters As CalculationParameters _
) As ListObject
    Dim ws As Worksheet
    Dim parameterSheet As Worksheet
    Dim headerCell As Range
    Dim table As ListObject

    For Each ws In workbook.Worksheets
        If VBA.StrComp(ws.Name, parameters.SheetName, vbTextCompare) = 0 Then
            Set parameterSheet = ws
            Exit For
        End If
    Next ws
    If parameterSheet Is Nothing Then
        private_Fail MSG_PARAMETER_SHEET_MISSING & parameters.SheetName
    End If
    Set headerCell = parameterSheet.Range(parameters.HeaderAddress)
    For Each table In parameterSheet.ListObjects
        If Not table.HeaderRowRange Is Nothing Then
            If table.HeaderRowRange.Cells(1, 1).Address = headerCell.Address Then
                Set private_ReadParameterTable = table
                Exit Function
            End If
        End If
    Next table
    private_Fail MSG_PARAMETER_TABLE_MISSING & parameters.SheetName & INPUT_SHEET_SEPARATOR & parameters.HeaderAddress
End Function

Private Function private_CompactParameterRows( _
    ByRef data As Variant, _
    ByRef sourceRows() As Long _
) As Long
    Dim i As Long
    Dim j As Long
    Dim count As Long
    Dim hasValue As Boolean

    ReDim sourceRows(1 To UBound(data, 1))
    For i = 1 To UBound(data, 1)
        private_CheckCancel i
        hasValue = False
        For j = 1 To UBound(data, 2)
            If VBA.IsError(data(i, j)) Then
                hasValue = True
            ElseIf VBA.Len(VBA.Trim$(VBA.CStr(data(i, j)))) > 0 Then
                hasValue = True
            End If
        Next j
        If hasValue Then
            count = count + 1
            sourceRows(count) = i
            If count <> i Then
                For j = 1 To UBound(data, 2)
                    data(count, j) = data(i, j)
                Next j
            End If
        End If
    Next i
    private_CompactParameterRows = count
End Function

Private Sub private_ReadCalculationParameters( _
    ByVal ws As Worksheet, _
    ByRef parameters As CalculationParameters _
)
    Dim fileSystem As Object

    parameters.InputPath = private_RequiredText(private_ReadFormValue("InputReference"), "Form.InputReference")
    parameters.EndDay = private_ReadDay(private_ReadFormValue("EndDate"), ws.Parent.Date1904, "Form.EndDate")
    private_ParseInputReference parameters
    parameters.InputPath = private_ResolveInputPath(parameters.InputPath)
    private_LogDebug MSG_INPUT_FILE & parameters.InputPath
    private_LogDebug MSG_END_DATE & VBA.Format$(VBA.CDate(parameters.EndDay), FORMAT_DATE)
    private_CheckCancel 0
    Set fileSystem = VBA.CreateObject(FILE_SYSTEM_PROG_ID)
    If Not fileSystem.FileExists(parameters.InputPath) Then
        private_Fail ws.Name & "!" & PARAM_INPUT_PATH_CELL & MSG_FILE_NOT_FOUND_OR_INACCESSIBLE & parameters.InputPath
    End If
End Sub

Private Sub private_ParseInputReference(ByRef parameters As CalculationParameters)
    Dim parts As Variant
    Dim location As String
    Dim separatorPosition As Long
    Dim address As String
    Dim i As Long
    Dim character As String
    Dim digitsStarted As Boolean

    parts = VBA.Split(parameters.InputPath, INPUT_REFERENCE_SEPARATOR)
    If UBound(parts) <> 1 Then
        private_Fail MSG_INVALID_REFERENCE
    End If
    parameters.InputPath = VBA.Trim$(parts(0))
    location = VBA.Trim$(parts(1))
    separatorPosition = VBA.InStrRev(location, INPUT_SHEET_SEPARATOR)
    If separatorPosition <= 1 Then
        private_Fail MSG_INVALID_REFERENCE
    End If
    parameters.SheetName = VBA.Trim$(VBA.Left$(location, separatorPosition - 1))
    If VBA.Left$(parameters.SheetName, 1) = "'" And VBA.Right$(parameters.SheetName, 1) = "'" Then
        parameters.SheetName = VBA.Replace(VBA.Mid$(parameters.SheetName, 2, _
            VBA.Len(parameters.SheetName) - 2), "''", "'")
    End If
    address = VBA.UCase$(VBA.Replace(VBA.Trim$(VBA.Mid$(location, separatorPosition + 1)), "$", ""))
    If VBA.Len(parameters.InputPath) = 0 Or VBA.Len(parameters.SheetName) = 0 Then
        private_Fail MSG_INVALID_REFERENCE
    End If
    If Not VBA.Left$(address, 1) Like "[A-Z]" Then
        private_Fail MSG_INVALID_REFERENCE
    End If
    For i = 1 To VBA.Len(address)
        character = VBA.Mid$(address, i, 1)
        If character Like "[0-9]" Then
            digitsStarted = True
        ElseIf Not character Like "[A-Z]" Or digitsStarted Then
            private_Fail MSG_INVALID_REFERENCE
        End If
    Next i
    If Not digitsStarted Then
        private_Fail MSG_INVALID_REFERENCE
    End If
    parameters.HeaderAddress = address
End Sub

Private Function private_ResolveInputPath(ByVal inputPath As String) As String
    Dim fileSystem As Object
    Dim basePath As String

    Set fileSystem = VBA.CreateObject(FILE_SYSTEM_PROG_ID)
    inputPath = VBA.Replace(inputPath, "/", "\")
    If VBA.Left$(inputPath, 2) = "\\" Or inputPath Like "[A-Za-z]:\*" Then
        private_ResolveInputPath = fileSystem.GetAbsolutePathName(inputPath)
        Exit Function
    End If
    If VBA.Left$(inputPath, 1) = "\" Or VBA.InStr(1, inputPath, ":", vbBinaryCompare) > 0 Then
        private_Fail MSG_AMBIGUOUS_PATH & inputPath
    End If
    basePath = ThisWorkbook.Path
    If VBA.Len(basePath) = 0 Or VBA.InStr(1, basePath, "://", vbBinaryCompare) > 0 Then
        private_Fail MSG_RELATIVE_PATH_BASE
    End If
    private_ResolveInputPath = fileSystem.GetAbsolutePathName(fileSystem.BuildPath(basePath, inputPath))
End Function

Private Sub private_CheckCancel(ByVal index As Long)
    Dim currentTime As Single

    If Not mRunning Then
        Exit Sub
    End If
    If mCancel Then
        VBA.Err.Raise CANCEL_ERROR, ERROR_SOURCE, MSG_CANCELLED
    End If
    If index Mod UI_YIELD_INTERVAL <> 0 Then
        Exit Sub
    End If
    currentTime = VBA.Timer
    If index <> 0 And currentTime >= mLastUiYield Then
        If currentTime - mLastUiYield < UI_YIELD_SECONDS Then
            Exit Sub
        End If
    End If
    mLastUiYield = currentTime
    Application.StatusBar = MSG_CALCULATING_DAYS & index
    VBA.DoEvents
    If mCancel Then
        VBA.Err.Raise CANCEL_ERROR, ERROR_SOURCE, MSG_CANCELLED
    End If
End Sub

Private Function private_IsExcluded( _
    ByVal eventName As String, _
    ByVal subtractTrips As Boolean _
) As Boolean
    private_IsExcluded = private_IsEventInList(eventName, EXCLUDED_EVENTS)
    If Not private_IsExcluded And subtractTrips Then
        private_IsExcluded = private_IsEventInList(eventName, OPTIONAL_EXCLUDED_EVENTS)
    End If
End Function

Private Function private_IsEventInList( _
    ByVal eventName As String, _
    ByVal eventList As String _
) As Boolean
    eventName = VBA.Trim$(eventName)
    If VBA.Len(eventName) = 0 Then
        Exit Function
    End If
    private_IsEventInList = VBA.InStr(1, "|" & eventList & "|", "|" & eventName & "|", vbTextCompare) > 0
End Function

Private Function private_ReadQueryTable(ByVal table As ListObject) As Variant
    Dim tableSource As obj_TableSource
    Dim tableQuery As obj_TableQuery
    Dim tableQueryService As obj_TableQueryService
    Dim dataTable As obj_DataTable
    Dim diagnostic As String
    Dim errorNumber As Long
    Dim errorDescription As String

    On Error GoTo Failed
    Set tableSource = New obj_TableSource
    Set tableQuery = New obj_TableQuery
    Set tableQueryService = New obj_TableQueryService
    If Not tableSource.Initialize() Then
        private_Fail table.Name
    End If
    If Not tableQuery.Initialize() Then
        private_Fail table.Name
    End If
    If Not tableQueryService.Initialize() Then
        private_Fail table.Name
    End If

    tableSource.WorkbookPath = table.Parent.Parent.FullName
    tableSource.SheetName = table.Parent.Name
    tableSource.TableName = table.Name
    tableQueryService.Backend = QueryExcel
    If Not tableQueryService.TryExecute(tableSource, tableQuery, dataTable, diagnostic) Then
        private_Fail diagnostic
    End If
    private_ReadQueryTable = dataTable.Values
Cleanup:
    If Not dataTable Is Nothing Then
        dataTable.Dispose
    End If
    If Not tableQueryService Is Nothing Then
        tableQueryService.Dispose
    End If
    If Not tableQuery Is Nothing Then
        tableQuery.Dispose
    End If
    If Not tableSource Is Nothing Then
        tableSource.Dispose
    End If
    If errorNumber <> 0 Then
        VBA.Err.Raise errorNumber, "ReadQueryTable", errorDescription
    End If
    Exit Function
Failed:
    errorNumber = VBA.Err.Number
    errorDescription = VBA.Err.Description
    Resume Cleanup
End Function

Private Function private_FindSource() As ListObject
    Dim wb As Workbook
    Dim ws As Worksheet
    Dim lo As ListObject

    For Each wb In Application.Workbooks
        For Each ws In wb.Worksheets
            If ws.Name = SOURCE_SHEET_NAME Then
                For Each lo In ws.ListObjects
                    If lo.Name = SOURCE_TABLE_NAME Then
                        If Not private_FindSource Is Nothing Then
                            private_Fail MSG_MULTIPLE_SOURCES
                        End If
                        Set private_FindSource = lo
                    End If
                Next lo
            End If
        Next ws
    Next wb
    If private_FindSource Is Nothing Then
        private_Fail MSG_OPEN_SOURCE & SOURCE_SHEET_NAME & MSG_TABLE & SOURCE_TABLE_NAME & "'."
    End If
End Function

Private Function private_RequireTable( _
    ByVal sheetName As String, _
    ByVal tableName As String, _
    Optional ByVal workbook As Workbook = Nothing _
) As ListObject
    Dim ws As Worksheet
    Dim lo As ListObject

    If workbook Is Nothing Then
        Set workbook = ThisWorkbook
    End If
    For Each ws In workbook.Worksheets
        If ws.Name = sheetName Then
            For Each lo In ws.ListObjects
                If lo.Name = tableName Then
                    Set private_RequireTable = lo
                    Exit Function
                End If
            Next lo
        End If
    Next ws
    private_Fail MSG_TABLE_NOT_FOUND & tableName & MSG_ON_WORKSHEET & sheetName & MSG_CHECK_TABLE_CONFIGURATION
End Function

Private Function private_ColumnIndex( _
    ByVal lo As ListObject, _
    ByVal header As String _
) As Long
    Dim col As ListColumn

    For Each col In lo.ListColumns
        If col.Name = header Then
            private_ColumnIndex = col.Index
            Exit Function
        End If
    Next col
    private_Fail MSG_TABLE_TEXT & lo.Name & MSG_IS_MISSING_COLUMN & header & "'."
End Function

Private Function private_RequiredText( _
    ByVal value As Variant, _
    ByVal context As String _
) As String
    If VBA.IsError(value) Or VBA.IsNull(value) Then
        private_Fail context & MSG_CELL_ERROR
    End If
    private_RequiredText = VBA.Trim$(VBA.Replace(VBA.CStr(value), VBA.ChrW(NBSP_CODE), " "))
    If VBA.Len(private_RequiredText) = 0 Then
        private_Fail context & MSG_REQUIRED_VALUE
    End If
End Function

Private Function private_ReadFlag( _
    ByVal value As Variant, _
    ByVal context As String _
) As Boolean
    Dim flagText As String

    If Not VBA.IsError(value) And Not VBA.IsNull(value) Then
        If VBA.Len(VBA.Trim$(VBA.Replace(VBA.CStr(value), VBA.ChrW(NBSP_CODE), " "))) = 0 Then
            private_ReadFlag = False
            Exit Function
        End If
    End If
    flagText = VBA.LCase$(private_RequiredText(value, context & MSG_SUBTRACT_BUSINESS_TRIPS))
    If private_IsEventInList(flagText, m_includeTripValues) Then
        private_ReadFlag = False
    ElseIf private_IsEventInList(flagText, m_excludeTripValues) Then
        private_ReadFlag = True
    Else
        private_Fail context & MSG_INVALID_TRIP_FLAG
    End If
End Function

Private Function private_ReadDay( _
    ByVal value As Variant, _
    ByVal date1904 As Boolean, _
    ByVal context As String _
) As Long
    Dim text As String
    Dim parsed As Date
    Dim serial As Double

    On Error GoTo InvalidDate
    If VBA.IsError(value) Or VBA.IsNull(value) Or VBA.IsEmpty(value) Then
        GoTo InvalidDate
    End If
    If VBA.VarType(value) = vbString Then
        text = VBA.Trim$(VBA.CStr(value))
        If Not text Like DATE_TEXT_PATTERN Then
            GoTo InvalidDate
        End If
        parsed = VBA.DateSerial(VBA.CInt(VBA.Right$(text, 4)), VBA.CInt(VBA.Mid$(text, 4, 2)), VBA.CInt(VBA.Left$(text, 2)))
        If VBA.Format$(parsed, FORMAT_DATE) <> text Then
            GoTo InvalidDate
        End If
        serial = VBA.CDbl(parsed)
    Else
        If VBA.VarType(value) = vbBoolean Or Not VBA.IsNumeric(value) Then
            GoTo InvalidDate
        End If
        serial = VBA.Int(VBA.CDbl(value))
        If date1904 Then
            serial = serial + DATE_1904_OFFSET
        End If
    End If
    If serial < MIN_DATE_SERIAL Or serial > MAX_DATE_SERIAL Then
        GoTo InvalidDate
    End If
    private_ReadDay = VBA.CLng(serial)
    Exit Function
InvalidDate:
    private_Fail context & MSG_INVALID_DATE
End Function

Private Sub private_Fail(ByVal message As String)
    VBA.Err.Raise OPERATION_ERROR, ERROR_SOURCE, message
End Sub

Private Sub private_ClearLog()
    ex_Core.fn_Diagnostic_WriteLog "PADC_CALCULATION_STARTED"
End Sub

Private Sub private_LogError(ByVal message As String)
    ex_Core.fn_Diagnostic_WriteLog "PADC_ERROR | " & message
End Sub

Private Sub private_LogWarning(ByVal message As String)
    ex_Core.fn_Diagnostic_WriteLog "PADC_WARNING | " & message
End Sub

Private Sub private_LogDebug(ByVal message As String)
    ex_Core.fn_Diagnostic_WriteLog "PADC | " & message
End Sub

Private Sub private_LogFailure( _
    ByVal errorNumber As Long, _
    ByVal errorText As String _
)
    ex_Core.fn_Diagnostic_WriteLog "PADC_FAILED | Number=" & errorNumber & " | " & errorText
End Sub

Private Function private_LoadConfiguration() As Boolean
    Dim key As String
    Dim thresholdText As String

    key = "PersonnelAtDisposalDaysCalculation::legacy.SOURCE_SHEET_NAME"
    If Not ex_Core.fn_TryGetWorkbookConfigValue(key, SOURCE_SHEET_NAME) Then
        GoTo MissingKey
    End If
    key = "PersonnelAtDisposalDaysCalculation::legacy.SOURCE_TABLE_NAME"
    If Not ex_Core.fn_TryGetWorkbookConfigValue(key, SOURCE_TABLE_NAME) Then
        GoTo MissingKey
    End If
    key = "PersonnelAtDisposalDaysCalculation::legacy.SOURCE_COL_NAME"
    If Not ex_Core.fn_TryGetWorkbookConfigValue(key, SOURCE_COL_NAME) Then
        GoTo MissingKey
    End If
    key = "PersonnelAtDisposalDaysCalculation::legacy.ROSTER_SHEET_NAME"
    If Not ex_Core.fn_TryGetWorkbookConfigValue(key, ROSTER_SHEET_NAME) Then
        GoTo MissingKey
    End If
    key = "PersonnelAtDisposalDaysCalculation::legacy.ROSTER_TABLE_NAME"
    If Not ex_Core.fn_TryGetWorkbookConfigValue(key, ROSTER_TABLE_NAME) Then
        GoTo MissingKey
    End If
    key = "PersonnelAtDisposalDaysCalculation::legacy.ROSTER_COL_NAME"
    If Not ex_Core.fn_TryGetWorkbookConfigValue(key, ROSTER_COL_NAME) Then
        GoTo MissingKey
    End If
    key = "PersonnelAtDisposalDaysCalculation::legacy.ROSTER_COL_TAX_ID"
    If Not ex_Core.fn_TryGetWorkbookConfigValue(key, ROSTER_COL_TAX_ID) Then
        GoTo MissingKey
    End If
    key = "PersonnelAtDisposalDaysCalculation::legacy.MSG_PERSON_UNVERIFIED"
    If Not ex_Core.fn_TryGetWorkbookConfigValue(key, MSG_PERSON_UNVERIFIED) Then
        GoTo MissingKey
    End If
    key = "PersonnelAtDisposalDaysCalculation::legacy.MSG_AMBIGUOUS_NAME"
    If Not ex_Core.fn_TryGetWorkbookConfigValue(key, MSG_AMBIGUOUS_NAME) Then
        GoTo MissingKey
    End If
    key = "PersonnelAtDisposalDaysCalculation::legacy.SOURCE_COL_TAX_ID"
    If Not ex_Core.fn_TryGetWorkbookConfigValue(key, SOURCE_COL_TAX_ID) Then
        GoTo MissingKey
    End If
    key = "PersonnelAtDisposalDaysCalculation::legacy.SOURCE_COL_EVENT"
    If Not ex_Core.fn_TryGetWorkbookConfigValue(key, SOURCE_COL_EVENT) Then
        GoTo MissingKey
    End If
    key = "PersonnelAtDisposalDaysCalculation::legacy.SOURCE_COL_PERIOD_FROM"
    If Not ex_Core.fn_TryGetWorkbookConfigValue(key, SOURCE_COL_PERIOD_FROM) Then
        GoTo MissingKey
    End If
    key = "PersonnelAtDisposalDaysCalculation::legacy.SOURCE_COL_PERIOD_TO"
    If Not ex_Core.fn_TryGetWorkbookConfigValue(key, SOURCE_COL_PERIOD_TO) Then
        GoTo MissingKey
    End If
    key = "PersonnelAtDisposalDaysCalculation::legacy.ERROR_COL_NAME"
    If Not ex_Core.fn_TryGetWorkbookConfigValue(key, ERROR_COL_NAME) Then
        GoTo MissingKey
    End If
    key = "PersonnelAtDisposalDaysCalculation::legacy.ERROR_COL_TAX_ID"
    If Not ex_Core.fn_TryGetWorkbookConfigValue(key, ERROR_COL_TAX_ID) Then
        GoTo MissingKey
    End If
    key = "PersonnelAtDisposalDaysCalculation::legacy.ERROR_COL_DESCRIPTION"
    If Not ex_Core.fn_TryGetWorkbookConfigValue(key, ERROR_COL_DESCRIPTION) Then
        GoTo MissingKey
    End If
    key = "PersonnelAtDisposalDaysCalculation::legacy.MSG_SKIPPED_PEOPLE"
    If Not ex_Core.fn_TryGetWorkbookConfigValue(key, MSG_SKIPPED_PEOPLE) Then
        GoTo MissingKey
    End If
    key = "PersonnelAtDisposalDaysCalculation::legacy.PARAM_COL_NAME"
    If Not ex_Core.fn_TryGetWorkbookConfigValue(key, PARAM_COL_NAME) Then
        GoTo MissingKey
    End If
    key = "PersonnelAtDisposalDaysCalculation::legacy.PARAM_COL_TAX_ID"
    If Not ex_Core.fn_TryGetWorkbookConfigValue(key, PARAM_COL_TAX_ID) Then
        GoTo MissingKey
    End If
    key = "PersonnelAtDisposalDaysCalculation::legacy.PARAM_COL_START"
    If Not ex_Core.fn_TryGetWorkbookConfigValue(key, PARAM_COL_START) Then
        GoTo MissingKey
    End If
    key = "PersonnelAtDisposalDaysCalculation::legacy.PARAM_COL_TRIPS"
    If Not ex_Core.fn_TryGetWorkbookConfigValue(key, PARAM_COL_TRIPS) Then
        GoTo MissingKey
    End If
    key = "PersonnelAtDisposalDaysCalculation::legacy.PRESENT_EVENT_NAME"
    If Not ex_Core.fn_TryGetWorkbookConfigValue(key, PRESENT_EVENT_NAME) Then
        GoTo MissingKey
    End If
    key = "PersonnelAtDisposalDaysCalculation::legacy.MSG_ALREADY_RUNNING"
    If Not ex_Core.fn_TryGetWorkbookConfigValue(key, MSG_ALREADY_RUNNING) Then
        GoTo MissingKey
    End If
    key = "PersonnelAtDisposalDaysCalculation::legacy.MSG_PARAMETER_LOADING_STARTED"
    If Not ex_Core.fn_TryGetWorkbookConfigValue(key, MSG_PARAMETER_LOADING_STARTED) Then
        GoTo MissingKey
    End If
    key = "PersonnelAtDisposalDaysCalculation::legacy.MSG_SELECT_TOOL_SHEET"
    If Not ex_Core.fn_TryGetWorkbookConfigValue(key, MSG_SELECT_TOOL_SHEET) Then
        GoTo MissingKey
    End If
    key = "PersonnelAtDisposalDaysCalculation::legacy.MSG_WRONG_WORKBOOK"
    If Not ex_Core.fn_TryGetWorkbookConfigValue(key, MSG_WRONG_WORKBOOK) Then
        GoTo MissingKey
    End If
    key = "PersonnelAtDisposalDaysCalculation::legacy.MSG_PARAMETER_WORKSHEET"
    If Not ex_Core.fn_TryGetWorkbookConfigValue(key, MSG_PARAMETER_WORKSHEET) Then
        GoTo MissingKey
    End If
    key = "PersonnelAtDisposalDaysCalculation::legacy.MSG_PARAMETER_LOADING_COMPLETED"
    If Not ex_Core.fn_TryGetWorkbookConfigValue(key, MSG_PARAMETER_LOADING_COMPLETED) Then
        GoTo MissingKey
    End If
    key = "PersonnelAtDisposalDaysCalculation::legacy.MSG_INPUT_FILE"
    If Not ex_Core.fn_TryGetWorkbookConfigValue(key, MSG_INPUT_FILE) Then
        GoTo MissingKey
    End If
    key = "PersonnelAtDisposalDaysCalculation::legacy.MSG_END_DATE"
    If Not ex_Core.fn_TryGetWorkbookConfigValue(key, MSG_END_DATE) Then
        GoTo MissingKey
    End If
    key = "PersonnelAtDisposalDaysCalculation::legacy.MSG_OPERATION_CANCELLED"
    If Not ex_Core.fn_TryGetWorkbookConfigValue(key, MSG_OPERATION_CANCELLED) Then
        GoTo MissingKey
    End If
    key = "PersonnelAtDisposalDaysCalculation::legacy.MSG_OPERATION_STOPPED"
    If Not ex_Core.fn_TryGetWorkbookConfigValue(key, MSG_OPERATION_STOPPED) Then
        GoTo MissingKey
    End If
    key = "PersonnelAtDisposalDaysCalculation::legacy.MSG_CANCELLING_OPERATION"
    If Not ex_Core.fn_TryGetWorkbookConfigValue(key, MSG_CANCELLING_OPERATION) Then
        GoTo MissingKey
    End If
    key = "PersonnelAtDisposalDaysCalculation::legacy.MSG_RUKH_SOURCE"
    If Not ex_Core.fn_TryGetWorkbookConfigValue(key, MSG_RUKH_SOURCE) Then
        GoTo MissingKey
    End If
    key = "PersonnelAtDisposalDaysCalculation::legacy.MSG_PARAMETER_CLOSE_FAILED"
    If Not ex_Core.fn_TryGetWorkbookConfigValue(key, MSG_PARAMETER_CLOSE_FAILED) Then
        GoTo MissingKey
    End If
    key = "PersonnelAtDisposalDaysCalculation::legacy.MSG_PARAMETER_SHEET_MISSING"
    If Not ex_Core.fn_TryGetWorkbookConfigValue(key, MSG_PARAMETER_SHEET_MISSING) Then
        GoTo MissingKey
    End If
    key = "PersonnelAtDisposalDaysCalculation::legacy.MSG_PARAMETERS_EMPTY"
    If Not ex_Core.fn_TryGetWorkbookConfigValue(key, MSG_PARAMETERS_EMPTY) Then
        GoTo MissingKey
    End If
    key = "PersonnelAtDisposalDaysCalculation::legacy.MSG_PEOPLE"
    If Not ex_Core.fn_TryGetWorkbookConfigValue(key, MSG_PEOPLE) Then
        GoTo MissingKey
    End If
    key = "PersonnelAtDisposalDaysCalculation::legacy.MSG_RUKH_EVENTS"
    If Not ex_Core.fn_TryGetWorkbookConfigValue(key, MSG_RUKH_EVENTS) Then
        GoTo MissingKey
    End If
    key = "PersonnelAtDisposalDaysCalculation::legacy.MSG_WORKSHEET_ROW"
    If Not ex_Core.fn_TryGetWorkbookConfigValue(key, MSG_WORKSHEET_ROW) Then
        GoTo MissingKey
    End If
    key = "PersonnelAtDisposalDaysCalculation::legacy.MSG_TAX_ID"
    If Not ex_Core.fn_TryGetWorkbookConfigValue(key, MSG_TAX_ID) Then
        GoTo MissingKey
    End If
    key = "PersonnelAtDisposalDaysCalculation::legacy.MSG_DUPLICATE_TAX_ID"
    If Not ex_Core.fn_TryGetWorkbookConfigValue(key, MSG_DUPLICATE_TAX_ID) Then
        GoTo MissingKey
    End If
    key = "PersonnelAtDisposalDaysCalculation::legacy.MSG_TIME_POINT"
    If Not ex_Core.fn_TryGetWorkbookConfigValue(key, MSG_TIME_POINT) Then
        GoTo MissingKey
    End If
    key = "PersonnelAtDisposalDaysCalculation::legacy.MSG_START_AFTER_END"
    If Not ex_Core.fn_TryGetWorkbookConfigValue(key, MSG_START_AFTER_END) Then
        GoTo MissingKey
    End If
    key = "PersonnelAtDisposalDaysCalculation::legacy.MSG_FULL_NAME"
    If Not ex_Core.fn_TryGetWorkbookConfigValue(key, MSG_FULL_NAME) Then
        GoTo MissingKey
    End If
    key = "PersonnelAtDisposalDaysCalculation::legacy.MSG_EVENT"
    If Not ex_Core.fn_TryGetWorkbookConfigValue(key, MSG_EVENT) Then
        GoTo MissingKey
    End If
    key = "PersonnelAtDisposalDaysCalculation::legacy.MSG_DEPARTURE"
    If Not ex_Core.fn_TryGetWorkbookConfigValue(key, MSG_DEPARTURE) Then
        GoTo MissingKey
    End If
    key = "PersonnelAtDisposalDaysCalculation::legacy.MSG_ARRIVAL_CELL_ERROR"
    If Not ex_Core.fn_TryGetWorkbookConfigValue(key, MSG_ARRIVAL_CELL_ERROR) Then
        GoTo MissingKey
    End If
    key = "PersonnelAtDisposalDaysCalculation::legacy.MSG_ARRIVAL"
    If Not ex_Core.fn_TryGetWorkbookConfigValue(key, MSG_ARRIVAL) Then
        GoTo MissingKey
    End If
    key = "PersonnelAtDisposalDaysCalculation::legacy.MSG_ARRIVAL_BEFORE_DEPARTURE"
    If Not ex_Core.fn_TryGetWorkbookConfigValue(key, MSG_ARRIVAL_BEFORE_DEPARTURE) Then
        GoTo MissingKey
    End If
    key = "PersonnelAtDisposalDaysCalculation::legacy.MSG_WRITING_RESULT"
    If Not ex_Core.fn_TryGetWorkbookConfigValue(key, MSG_WRITING_RESULT) Then
        GoTo MissingKey
    End If
    key = "PersonnelAtDisposalDaysCalculation::legacy.MSG_ROWS"
    If Not ex_Core.fn_TryGetWorkbookConfigValue(key, MSG_ROWS) Then
        GoTo MissingKey
    End If
    key = "PersonnelAtDisposalDaysCalculation::legacy.MSG_TARGET_PREPARE_FAILED"
    If Not ex_Core.fn_TryGetWorkbookConfigValue(key, MSG_TARGET_PREPARE_FAILED) Then
        GoTo MissingKey
    End If
    key = "PersonnelAtDisposalDaysCalculation::legacy.MSG_WRITING_RESULTS"
    If Not ex_Core.fn_TryGetWorkbookConfigValue(key, MSG_WRITING_RESULTS) Then
        GoTo MissingKey
    End If
    key = "PersonnelAtDisposalDaysCalculation::legacy.MSG_CALCULATION_COMPLETED_PEOPLE"
    If Not ex_Core.fn_TryGetWorkbookConfigValue(key, MSG_CALCULATION_COMPLETED_PEOPLE) Then
        GoTo MissingKey
    End If
    key = "PersonnelAtDisposalDaysCalculation::legacy.MSG_RELATIVE_PATH_BASE"
    If Not ex_Core.fn_TryGetWorkbookConfigValue(key, MSG_RELATIVE_PATH_BASE) Then
        GoTo MissingKey
    End If
    key = "PersonnelAtDisposalDaysCalculation::legacy.MSG_AMBIGUOUS_PATH"
    If Not ex_Core.fn_TryGetWorkbookConfigValue(key, MSG_AMBIGUOUS_PATH) Then
        GoTo MissingKey
    End If
    key = "PersonnelAtDisposalDaysCalculation::legacy.MSG_FILE_NOT_FOUND_OR_INACCESSIBLE"
    If Not ex_Core.fn_TryGetWorkbookConfigValue(key, MSG_FILE_NOT_FOUND_OR_INACCESSIBLE) Then
        GoTo MissingKey
    End If
    key = "PersonnelAtDisposalDaysCalculation::legacy.MSG_CANCELLED"
    If Not ex_Core.fn_TryGetWorkbookConfigValue(key, MSG_CANCELLED) Then
        GoTo MissingKey
    End If
    key = "PersonnelAtDisposalDaysCalculation::legacy.MSG_CALCULATING_DAYS"
    If Not ex_Core.fn_TryGetWorkbookConfigValue(key, MSG_CALCULATING_DAYS) Then
        GoTo MissingKey
    End If
    key = "PersonnelAtDisposalDaysCalculation::legacy.MSG_MULTIPLE_SOURCES"
    If Not ex_Core.fn_TryGetWorkbookConfigValue(key, MSG_MULTIPLE_SOURCES) Then
        GoTo MissingKey
    End If
    key = "PersonnelAtDisposalDaysCalculation::legacy.MSG_OPEN_SOURCE"
    If Not ex_Core.fn_TryGetWorkbookConfigValue(key, MSG_OPEN_SOURCE) Then
        GoTo MissingKey
    End If
    key = "PersonnelAtDisposalDaysCalculation::legacy.MSG_TABLE"
    If Not ex_Core.fn_TryGetWorkbookConfigValue(key, MSG_TABLE) Then
        GoTo MissingKey
    End If
    key = "PersonnelAtDisposalDaysCalculation::legacy.MSG_TABLE_NOT_FOUND"
    If Not ex_Core.fn_TryGetWorkbookConfigValue(key, MSG_TABLE_NOT_FOUND) Then
        GoTo MissingKey
    End If
    key = "PersonnelAtDisposalDaysCalculation::legacy.MSG_ON_WORKSHEET"
    If Not ex_Core.fn_TryGetWorkbookConfigValue(key, MSG_ON_WORKSHEET) Then
        GoTo MissingKey
    End If
    key = "PersonnelAtDisposalDaysCalculation::legacy.MSG_CHECK_TABLE_CONFIGURATION"
    If Not ex_Core.fn_TryGetWorkbookConfigValue(key, MSG_CHECK_TABLE_CONFIGURATION) Then
        GoTo MissingKey
    End If
    key = "PersonnelAtDisposalDaysCalculation::legacy.MSG_TABLE_TEXT"
    If Not ex_Core.fn_TryGetWorkbookConfigValue(key, MSG_TABLE_TEXT) Then
        GoTo MissingKey
    End If
    key = "PersonnelAtDisposalDaysCalculation::legacy.MSG_IS_MISSING_COLUMN"
    If Not ex_Core.fn_TryGetWorkbookConfigValue(key, MSG_IS_MISSING_COLUMN) Then
        GoTo MissingKey
    End If
    key = "PersonnelAtDisposalDaysCalculation::legacy.MSG_CELL_ERROR"
    If Not ex_Core.fn_TryGetWorkbookConfigValue(key, MSG_CELL_ERROR) Then
        GoTo MissingKey
    End If
    key = "PersonnelAtDisposalDaysCalculation::legacy.MSG_REQUIRED_VALUE"
    If Not ex_Core.fn_TryGetWorkbookConfigValue(key, MSG_REQUIRED_VALUE) Then
        GoTo MissingKey
    End If
    key = "PersonnelAtDisposalDaysCalculation::legacy.MSG_SUBTRACT_BUSINESS_TRIPS"
    If Not ex_Core.fn_TryGetWorkbookConfigValue(key, MSG_SUBTRACT_BUSINESS_TRIPS) Then
        GoTo MissingKey
    End If
    key = "PersonnelAtDisposalDaysCalculation::legacy.MSG_INVALID_TRIP_FLAG"
    If Not ex_Core.fn_TryGetWorkbookConfigValue(key, MSG_INVALID_TRIP_FLAG) Then
        GoTo MissingKey
    End If
    key = "PersonnelAtDisposalDaysCalculation::legacy.MSG_INVALID_DATE"
    If Not ex_Core.fn_TryGetWorkbookConfigValue(key, MSG_INVALID_DATE) Then
        GoTo MissingKey
    End If
    key = "PersonnelAtDisposalDaysCalculation::legacy.MSG_INVALID_REFERENCE"
    If Not ex_Core.fn_TryGetWorkbookConfigValue(key, MSG_INVALID_REFERENCE) Then
        GoTo MissingKey
    End If
    key = "PersonnelAtDisposalDaysCalculation::legacy.MSG_PARAMETER_TABLE_MISSING"
    If Not ex_Core.fn_TryGetWorkbookConfigValue(key, MSG_PARAMETER_TABLE_MISSING) Then
        GoTo MissingKey
    End If
    key = "PersonnelAtDisposalDaysCalculation::event.alwaysExcluded"
    If Not ex_Core.fn_TryGetWorkbookConfigValue(key, EXCLUDED_EVENTS) Then
        GoTo MissingKey
    End If
    key = "PersonnelAtDisposalDaysCalculation::event.optionalExcluded"
    If Not ex_Core.fn_TryGetWorkbookConfigValue(key, OPTIONAL_EXCLUDED_EVENTS) Then
        GoTo MissingKey
    End If
    key = "PersonnelAtDisposalDaysCalculation::calculation.thresholdDays"
    If Not ex_Core.fn_TryGetWorkbookConfigValue(key, thresholdText) Then
        GoTo MissingKey
    End If
    If Not VBA.IsNumeric(thresholdText) Then
        GoTo MissingKey
    End If
    COUNTED_DAYS_THRESHOLD = VBA.CLng(thresholdText)
    If COUNTED_DAYS_THRESHOLD < 1 Then
        GoTo MissingKey
    End If
    key = "PersonnelAtDisposalDaysCalculation::header.FullName"
    If Not ex_Core.fn_TryGetWorkbookConfigValue(key, TARGET_COL_NAME) Then
        GoTo MissingKey
    End If
    key = "PersonnelAtDisposalDaysCalculation::header.TaxId"
    If Not ex_Core.fn_TryGetWorkbookConfigValue(key, TARGET_COL_TAX_ID) Then
        GoTo MissingKey
    End If
    key = "PersonnelAtDisposalDaysCalculation::header.StartDate"
    If Not ex_Core.fn_TryGetWorkbookConfigValue(key, TARGET_COL_START) Then
        GoTo MissingKey
    End If
    key = "PersonnelAtDisposalDaysCalculation::header.BusinessTrips"
    If Not ex_Core.fn_TryGetWorkbookConfigValue(key, TARGET_COL_TRIPS) Then
        GoTo MissingKey
    End If
    key = "PersonnelAtDisposalDaysCalculation::header.Periods"
    If Not ex_Core.fn_TryGetWorkbookConfigValue(key, TARGET_COL_PERIODS) Then
        GoTo MissingKey
    End If
    key = "PersonnelAtDisposalDaysCalculation::header.CountedDays"
    If Not ex_Core.fn_TryGetWorkbookConfigValue(key, TARGET_COL_DAYS) Then
        GoTo MissingKey
    End If
    key = "PersonnelAtDisposalDaysCalculation::header.ThresholdDate"
    If Not ex_Core.fn_TryGetWorkbookConfigValue(key, TARGET_COL_THRESHOLD_DATE) Then
        GoTo MissingKey
    End If
    key = "PersonnelAtDisposalDaysCalculation::flag.includeTrips"
    If Not ex_Core.fn_TryGetWorkbookConfigValue(key, m_includeTripValues) Then
        GoTo MissingKey
    End If
    key = "PersonnelAtDisposalDaysCalculation::flag.excludeTrips"
    If Not ex_Core.fn_TryGetWorkbookConfigValue(key, m_excludeTripValues) Then
        GoTo MissingKey
    End If
    private_LoadConfiguration = True
    Exit Function
MissingKey:
    VBA.MsgBox "Required configuration key not found: " & key, vbExclamation
End Function

Private Function private_ReadFormValue(ByVal fieldName As String) As Variant
    Dim value As Variant
    Dim valueObject As Object
    Dim isObject As Boolean

    If Not m_pageBase.BindingContext.TryGetValue("Form", fieldName, value, valueObject, isObject) Then
        private_Fail "Form." & fieldName
    End If
    private_ReadFormValue = value
End Function

Private Sub private_PublishResults( _
    ByRef names() As Variant, _
    ByRef ids() As Variant, _
    ByRef starts() As Variant, _
    ByRef trips() As Variant, _
    ByRef periods() As Variant, _
    ByRef days() As Variant, _
    ByRef thresholds() As Variant, _
    ByVal count As Long, _
    ByVal failures As Collection _
)
    Dim values() As Variant
    Dim headers As Variant
    Dim errorValues() As Variant
    Dim errorHeaders As Variant
    Dim item As Variant
    Dim i As Long
    Dim resultTable As obj_UiRawTable
    Dim errorTable As obj_UiRawTable

    headers = VBA.Array(TARGET_COL_NAME, TARGET_COL_TAX_ID, TARGET_COL_START, _
        TARGET_COL_TRIPS, TARGET_COL_PERIODS, TARGET_COL_DAYS, TARGET_COL_THRESHOLD_DATE)
    Set resultTable = New obj_UiRawTable
    If count > 0 Then
        ReDim values(1 To count, 1 To 7)
        For i = 1 To count
            values(i, 1) = names(i, 1)
            values(i, 2) = ids(i, 1)
            values(i, 3) = VBA.Format$(starts(i, 1), FORMAT_DATE)
            values(i, 4) = trips(i, 1)
            values(i, 5) = periods(i, 1)
            values(i, 6) = days(i, 1)
            If Not VBA.IsEmpty(thresholds(i, 1)) Then
                values(i, 7) = VBA.Format$(thresholds(i, 1), FORMAT_DATE)
            End If
        Next i
        If Not resultTable.Initialize(values, headers) Then
            private_Fail MSG_TARGET_PREPARE_FAILED
        End If
    Else
        If Not resultTable.InitializeEmpty(headers) Then
            private_Fail MSG_TARGET_PREPARE_FAILED
        End If
    End If
    errorHeaders = VBA.Array(ERROR_COL_NAME, ERROR_COL_TAX_ID, ERROR_COL_DESCRIPTION)
    Set errorTable = New obj_UiRawTable
    If failures.Count > 0 Then
        ReDim errorValues(1 To failures.Count, 1 To 3)
        For i = 1 To failures.Count
            item = failures(i)
            errorValues(i, 1) = item(0)
            errorValues(i, 2) = item(1)
            errorValues(i, 3) = item(2)
        Next i
        If Not errorTable.Initialize(errorValues, errorHeaders) Then
            private_Fail MSG_TARGET_PREPARE_FAILED
        End If
    Else
        If Not errorTable.InitializeEmpty(errorHeaders) Then
            private_Fail MSG_TARGET_PREPARE_FAILED
        End If
    End If
    If Not m_pageBase.BindingContext.SetObject("Data", "CalculationRows", resultTable) Then
        private_Fail MSG_TARGET_PREPARE_FAILED
    End If
    If Not m_pageBase.BindingContext.SetObject("Data", "Errors", errorTable) Then
        private_Fail MSG_TARGET_PREPARE_FAILED
    End If
    If Not m_pageBase.Render() Then
        private_Fail MSG_TARGET_PREPARE_FAILED
    End If
End Sub