VERSION 1.0 CLASS
BEGIN
  MultiUse = -1
END
Attribute VB_Name = "obj_PADC_EventPreparation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private m_isInitialized As Boolean
Private m_isDisposed As Boolean

Private m_configuration As obj_PADC_Configuration
Private m_dataSource As obj_PADC_DataSource
Private m_validation As obj_PADC_Validation
Private m_runContext As obj_PADC_RunContext
Private Const DICTIONARY_PROG_ID As String = "Scripting.Dictionary"

' //
' // Lifecycle
' //
Private Sub Class_Initialize()
End Sub

Private Sub Class_Terminate()
    Me.Dispose
End Sub

' //
' // API
' //
Public Function Initialize( _
    ByVal configuration As obj_PADC_Configuration, _
    ByVal dataSource As obj_PADC_DataSource, _
    ByVal validation As obj_PADC_Validation, _
    ByVal runContext As obj_PADC_RunContext _
) As Boolean
    If m_isInitialized Or m_isDisposed Then
        Exit Function
    End If
    If configuration Is Nothing Then
        Exit Function
    End If
    Set m_configuration = configuration
    If dataSource Is Nothing Then
        Exit Function
    End If
    Set m_dataSource = dataSource
    If validation Is Nothing Then
        Exit Function
    End If
    Set m_validation = validation
    If runContext Is Nothing Then
        Exit Function
    End If
    Set m_runContext = runContext
    m_isInitialized = True
    Initialize = True
End Function

Public Sub Dispose()
    If m_isDisposed Then
        Exit Sub
    End If
    m_isInitialized = False
    m_isDisposed = True
    Set m_configuration = Nothing
    Set m_dataSource = Nothing
    Set m_validation = Nothing
    Set m_runContext = Nothing
End Sub

Public Function Prepare(ByVal parameters As obj_PADC_Parameters) As Collection
    Dim personnel As obj_DataTable
    Dim orders As obj_DataTable
    Dim enrollment As obj_DataTable
    Dim events As obj_DataTable
    Dim personnelValues As Variant
    Dim eventValues As Variant
    Dim dates As Object
    Dim rosterDates As Object
    Dim disposalDates As Object
    Dim dateOverrides As Object
    Dim dateLookupKeys As Object
    Dim seen As Object
    Dim eventIndex As obj_PADC_PersonIndex
    Dim prepared As Collection
    Dim preparedPerson As obj_PADC_PreparedPerson
    Dim matches As Collection
    Dim intervals As Collection
    Dim chronology As Collection
    Dim classifiedIntervals As Collection
    Dim interval As Variant
    Dim eventDecisions As Object
    Dim excluded As Boolean
    Dim matchRow As Variant
    Dim rowIndex As Long
    Dim cutoff As Long
    Dim startDay As Long
    Dim fallbackDay As Long
    Dim note As String
    Dim departure As Long
    Dim arrival As Long
    Dim latestDeparture As Long
    Dim fullName As String
    Dim taxId As String
    Dim rank As String
    Dim status As String
    Dim eventName As String
    Dim openStatus As String
    Dim errorText As String
    Dim matchError As String
    Dim disposalCode As String
    Dim rosterCode As String
    Dim context As String

    private_EnsureReady
    cutoff = parameters.EndDay - 1
    fallbackDay = m_validation.ReadDay(m_configuration.GetText("calculation.fallbackStartDate"), False, "calculation.fallbackStartDate")
    disposalCode = m_configuration.GetText("source.personnel.disposalCode")
    rosterCode = m_configuration.GetText("source.personnel.rosterCode")
    m_runContext.BeginStep m_configuration.GetText("progress.Personnel")
    Set personnel = m_dataSource.ReadConfiguredSource("source.personnel", parameters.InputPath, _
        VBA.Array("TaxId", "FullName", "Rank", "Code"), parameters.SheetName, parameters.RangeAddress)
    m_runContext.CompleteStep
    m_runContext.BeginStep m_configuration.GetText("progress.Orders")
    Set orders = m_dataSource.ReadConfiguredSource("source.orders", vbNullString, _
        VBA.Array("TaxId", "FullName", "Action", "Date"))
    m_runContext.CompleteStep
    m_runContext.BeginStep m_configuration.GetText("progress.Enrollment")
    Set enrollment = m_dataSource.ReadConfiguredSource("source.enrollment", _
        VBA.CreateObject("WScript.Shell").ExpandEnvironmentStrings(m_configuration.GetText("source.orders.path")), _
        VBA.Array("TaxId", "FullName", "Date"))
    m_runContext.CompleteStep
    m_runContext.BeginStep m_configuration.GetText("progress.Events")
    Set events = m_dataSource.ReadConfiguredSource("source.events", vbNullString, _
        VBA.Array("TaxId", "FullName", "Event", "Departure", "Arrival"))
    m_runContext.CompleteStep
    personnelValues = personnel.Values
    Set dateLookupKeys = private_DateLookupKeys(personnelValues, personnel.RowCount, disposalCode, rosterCode)
    Set disposalDates = private_IndexDates(orders, cutoff, False, dateLookupKeys)
    Set rosterDates = private_IndexDates(enrollment, cutoff, True, dateLookupKeys)
    m_runContext.BeginStep m_configuration.GetText("progress.DateOverrides")
    Set dateOverrides = private_ReadDateOverrides()
    m_runContext.CompleteStep
    Set seen = VBA.CreateObject(DICTIONARY_PROG_ID)
    Set prepared = New Collection
    eventValues = events.Values
    Set eventIndex = New obj_PADC_PersonIndex
    m_runContext.BeginStep m_configuration.GetText("progress.EventIndex"), events.RowCount
    If Not eventIndex.Initialize(eventValues, events.RowCount, 1, 2, m_configuration, m_validation, m_runContext) Then
        m_configuration.Fail m_configuration.GetText("text.TargetPrepareFailed")
    End If
    m_runContext.CompleteStep

    m_runContext.BeginStep m_configuration.GetText("progress.Preparation"), personnel.RowCount
    For rowIndex = 1 To personnel.RowCount
        status = m_validation.MatchText(personnelValues(rowIndex, 4))
        If VBA.StrComp(status, disposalCode, vbTextCompare) = 0 Or VBA.StrComp(status, rosterCode, vbTextCompare) = 0 Then
            context = parameters.SheetName & m_configuration.GetText("text.ParameterRow") & rowIndex
            taxId = m_validation.MatchText(personnelValues(rowIndex, 1))
            fullName = m_validation.MatchText(personnelValues(rowIndex, 2))
            If VBA.Len(taxId) = 0 And VBA.Len(fullName) = 0 Then
                GoTo NextPersonnelRow
            End If
            rank = m_validation.MatchText(personnelValues(rowIndex, 3))
            startDay = 0
            note = vbNullString
            latestDeparture = 0
            openStatus = vbNullString
            errorText = vbNullString
            Set intervals = New Collection
            Set chronology = New Collection
            If VBA.Len(taxId) = 0 Or VBA.Len(fullName) = 0 Then
                errorText = m_configuration.GetText("text.PersonIdentityMissing")
            ElseIf seen.Exists(taxId) Then
                m_configuration.Fail context & m_configuration.GetText("text.DuplicateTaxId") & taxId
            Else
                seen.Add taxId, True
                If VBA.StrComp(status, rosterCode, vbTextCompare) = 0 Then
                    Set dates = rosterDates
                Else
                    Set dates = disposalDates
                End If
                If dateOverrides.Exists(fullName) Then
                    startDay = dateOverrides(fullName)
                    note = m_configuration.GetText("text.OverrideDateUsed")
                ElseIf dates.Exists("I|" & taxId) Then
                    startDay = dates("I|" & taxId)
                ElseIf dates.Exists("N|" & fullName) Then
                    startDay = dates("N|" & fullName)
                End If
                If startDay <= 0 Then
                    startDay = fallbackDay
                    note = m_configuration.GetText("text.FallbackDateUsed")
                End If
                If startDay > cutoff Then
                    errorText = m_configuration.GetText("text.StartAfterEnd")
                End If

                Set matches = eventIndex.Find(taxId, fullName, matchError)
                If VBA.Len(matchError) > 0 Then
                    If VBA.Len(errorText) > 0 Then
                        errorText = errorText & " | "
                    End If
                    errorText = errorText & matchError
                Else
                    For Each matchRow In matches
                        m_runContext.CheckCancel VBA.CLng(matchRow)
                        context = m_configuration.GetText("source.events.sheet") & m_configuration.GetText("text.ParameterRow") & matchRow
                        departure = m_validation.ReadDay(eventValues(matchRow, 4), False, context & m_configuration.GetText("text.Departure"))
                        If departure <= cutoff Then
                            eventName = m_validation.RequiredText(eventValues(matchRow, 3), context & m_configuration.GetText("text.Event"))
                            arrival = parameters.EndDay
                            If VBA.IsError(eventValues(matchRow, 5)) Then
                                m_configuration.Fail context & m_configuration.GetText("text.ArrivalCellError")
                            End If
                            If VBA.Len(m_validation.MatchText(eventValues(matchRow, 5))) > 0 Then
                                arrival = m_validation.ReadDay(eventValues(matchRow, 5), False, context & m_configuration.GetText("text.Arrival"))
                                If arrival < departure Then
                                    m_configuration.Fail context & m_configuration.GetText("text.ArrivalBeforeDeparture")
                                End If
                            End If
                            If m_validation.MatchText(eventValues(matchRow, 1)) = taxId Then
                                chronology.Add VBA.Array(departure, arrival, eventName)
                                If arrival > cutoff And (departure > latestDeparture Or _
                                        (departure = latestDeparture And VBA.StrComp(eventName, openStatus, vbBinaryCompare) > 0)) Then
                                    latestDeparture = departure
                                    openStatus = eventName
                                End If
                            End If
                            If departure < startDay Then
                                departure = startDay
                            End If
                            If arrival > parameters.EndDay Then
                                arrival = parameters.EndDay
                            End If
                            If startDay > 0 And departure < arrival Then
                                intervals.Add VBA.Array(departure, arrival, eventName)
                            End If
                        End If
                    Next matchRow
                End If
            End If

            If Not m_validation.IsEligibleOpenStatus(openStatus) Then
                GoTo NextPersonnelRow
            End If
            Set eventDecisions = m_validation.ResolveEventRules(status, openStatus)
            Set classifiedIntervals = New Collection
            For Each interval In intervals
                excluded = False
                If eventDecisions.Exists(VBA.CStr(interval(2))) Then
                    excluded = eventDecisions(VBA.CStr(interval(2)))
                End If
                classifiedIntervals.Add VBA.Array(interval(0), interval(1), interval(2), excluded)
            Next interval
            openStatus = m_validation.DisplayOpenStatus(openStatus, latestDeparture, chronology)
            Set preparedPerson = New obj_PADC_PreparedPerson
            If Not preparedPerson.Initialize(fullName, taxId, startDay, classifiedIntervals, errorText, note, rank, status, openStatus) Then
                m_configuration.Fail m_configuration.GetText("text.TargetPrepareFailed")
            End If
            prepared.Add preparedPerson
        End If
NextPersonnelRow:
        m_runContext.ReportProgress rowIndex
    Next rowIndex
    m_runContext.CompleteStep
    If prepared.Count = 0 Then
        m_configuration.Fail m_configuration.GetText("text.ParametersEmpty")
    End If
    Set Prepare = prepared
End Function

' //
' // Private
' //
Private Function private_ReadDateOverrides() As Object
    Dim fileSystem As Object
    Dim workbookPath As String
    Dim data As obj_DataTable
    Dim overrides As Object
    Dim values As Variant
    Dim rowIndex As Long
    Dim fullName As String
    Dim startDay As Long
    Dim context As String

    Set overrides = VBA.CreateObject(DICTIONARY_PROG_ID)
    overrides.CompareMode = vbTextCompare
    Set private_ReadDateOverrides = overrides
    Set fileSystem = VBA.CreateObject("Scripting.FileSystemObject")
    workbookPath = fileSystem.BuildPath(ThisWorkbook.Path, m_configuration.GetText("source.dateOverrides.fileName"))
    If Not fileSystem.FileExists(workbookPath) Then
        Exit Function
    End If

    Set data = m_dataSource.ReadConfiguredSource("source.dateOverrides", workbookPath, VBA.Array("FullName", "Date"))
    values = data.Values
    For rowIndex = 1 To data.RowCount
        m_runContext.CheckCancel rowIndex
        fullName = m_validation.MatchText(values(rowIndex, 1))
        If VBA.Len(fullName) > 0 Or VBA.Len(m_validation.MatchText(values(rowIndex, 2))) > 0 _
                Or VBA.IsError(values(rowIndex, 1)) Or VBA.IsError(values(rowIndex, 2)) Then
            context = workbookPath & m_configuration.GetText("text.ParameterRow") & rowIndex
            fullName = m_validation.RequiredText(values(rowIndex, 1), context & m_configuration.GetText("text.FullName"))
            startDay = m_validation.ReadDay(values(rowIndex, 2), False, context & m_configuration.GetText("text.TimePoint"))
            If overrides.Exists(fullName) Then
                If overrides(fullName) <> startDay Then
                    m_configuration.Fail context & " | " & m_configuration.GetText("text.OverrideDateConflict") & " " & fullName
                End If
            Else
                overrides.Add fullName, startDay
            End If
        End If
    Next rowIndex
End Function

Private Function private_DateLookupKeys( _
    ByRef personnelValues As Variant, _
    ByVal rowCount As Long, _
    ByVal disposalCode As String, _
    ByVal rosterCode As String _
) As Object
    Dim keys As Object
    Dim rowIndex As Long
    Dim code As String
    Dim taxId As String
    Dim fullName As String

    Set keys = VBA.CreateObject(DICTIONARY_PROG_ID)
    keys.CompareMode = vbTextCompare
    For rowIndex = 1 To rowCount
        code = m_validation.MatchText(personnelValues(rowIndex, 4))
        If VBA.StrComp(code, disposalCode, vbTextCompare) = 0 Or VBA.StrComp(code, rosterCode, vbTextCompare) = 0 Then
            taxId = m_validation.MatchText(personnelValues(rowIndex, 1))
            fullName = m_validation.MatchText(personnelValues(rowIndex, 2))
            If VBA.Len(taxId) > 0 And VBA.Len(fullName) > 0 Then
                keys(code & "|I|" & taxId) = True
                keys(code & "|N|" & fullName) = True
            End If
        End If
    Next rowIndex
    Set private_DateLookupKeys = keys
End Function

Private Function private_IndexDates( _
    ByVal orders As obj_DataTable, _
    ByVal cutoff As Long, _
    ByVal isEnrollment As Boolean, _
    ByVal dateLookupKeys As Object _
) As Object
    Dim dates As Object
    Dim nameOwners As Object
    Dim values As Variant
    Dim actionName As String
    Dim taxId As String
    Dim fullName As String
    Dim day As Long
    Dim rowIndex As Long
    Dim datePattern As Object
    Dim stageKey As String
    Dim startedAt As Double
    Dim parsedRows As Long
    Dim targetCode As String

    Set dates = VBA.CreateObject(DICTIONARY_PROG_ID)
    dates.CompareMode = vbTextCompare
    Set nameOwners = VBA.CreateObject(DICTIONARY_PROG_ID)
    nameOwners.CompareMode = vbTextCompare
    actionName = m_configuration.GetText("source.orders.action")
    stageKey = "progress.OrderIndex"
    targetCode = m_configuration.GetText("source.personnel.disposalCode")
    If isEnrollment Then
        stageKey = "progress.EnrollmentIndex"
        targetCode = m_configuration.GetText("source.personnel.rosterCode")
        Set datePattern = VBA.CreateObject("VBScript.RegExp")
        datePattern.Global = True
        datePattern.Pattern = "(^|[^0-9])([0-9]{1,2})[./]([0-9]{1,2})[./]([0-9]{4})(?![0-9])"
    End If
    values = orders.Values
    startedAt = VBA.Timer
    m_runContext.BeginStep m_configuration.GetText(stageKey), orders.RowCount
    For rowIndex = 1 To orders.RowCount
        If isEnrollment Or VBA.StrComp(m_validation.MatchText(values(rowIndex, 3)), actionName, vbTextCompare) = 0 Then
            taxId = m_validation.MatchText(values(rowIndex, 1))
            fullName = m_validation.MatchText(values(rowIndex, 2))
            If Not dateLookupKeys.Exists(targetCode & "|I|" & taxId) _
                    And Not dateLookupKeys.Exists(targetCode & "|N|" & fullName) Then
                GoTo NextOrderRow
            End If
            parsedRows = parsedRows + 1
            If isEnrollment Then
                day = private_EnrollmentDay(values(rowIndex, 3), cutoff, datePattern)
            Else
                day = private_OrderDay(values(rowIndex, 4), cutoff)
            End If
            If day > 0 And VBA.Len(taxId) > 0 Then
                private_KeepLatest dates, "I|" & taxId, day
            End If
            If VBA.Len(fullName) > 0 Then
                If Not nameOwners.Exists(fullName) Then
                    nameOwners.Add fullName, taxId
                ElseIf nameOwners(fullName) <> taxId Then
                    dates("N|" & fullName) = -1
                End If
                If day > 0 Then
                    private_KeepLatest dates, "N|" & fullName, day
                End If
            End If
        End If
NextOrderRow:
        m_runContext.ReportProgress rowIndex
    Next rowIndex
    m_runContext.CompleteStep
    ex_Core.fn_Diagnostic_WritePerf "PADC." & stageKey & " | ParsedRows=" & parsedRows, startedAt
    Set private_IndexDates = dates
End Function

Private Function private_EnrollmentDay( _
    ByVal value As Variant, _
    ByVal cutoff As Long, _
    ByVal datePattern As Object _
) As Long
    Dim hit As Object
    Dim day As Long
    Dim dateText As String

    If VBA.VarType(value) <> vbString Then
        private_EnrollmentDay = private_OrderDay(value, cutoff)
        Exit Function
    End If
    For Each hit In datePattern.Execute(VBA.CStr(value))
        dateText = VBA.Right$("0" & hit.SubMatches(1), 2) & "." _
            & VBA.Right$("0" & hit.SubMatches(2), 2) & "." & hit.SubMatches(3)
        day = private_OrderDay(dateText, cutoff)
        If day > private_EnrollmentDay Then
            private_EnrollmentDay = day
        End If
    Next hit
End Function

Private Sub private_KeepLatest( _
    ByVal dates As Object, _
    ByVal key As String, _
    ByVal day As Long _
)
    If Not dates.Exists(key) Then
        dates.Add key, day
    ElseIf dates(key) >= 0 And dates(key) < day Then
        dates(key) = day
    End If
End Sub

Private Function private_OrderDay( _
    ByVal value As Variant, _
    ByVal cutoff As Long _
) As Long
    Dim day As Long

    If VBA.IsError(value) Or VBA.IsNull(value) Or VBA.IsEmpty(value) Then
        Exit Function
    End If
    On Error GoTo InvalidDate
    day = m_validation.ReadDay(value, False, vbNullString)
    If day <= cutoff Then
        private_OrderDay = day
    End If
    Exit Function
InvalidDate:
    ' A missing or malformed order date is reported against the affected person.
End Function

Private Sub private_EnsureReady()
    If Not m_isInitialized Or m_isDisposed Then
        VBA.Err.Raise vbObjectError + 2100, "obj_PADC_EventPreparation", "Service is not initialized."
    End If
End Sub