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
    Dim events As obj_DataTable
    Dim personnelValues As Variant
    Dim eventValues As Variant
    Dim dates As Object
    Dim seen As Object
    Dim eventIndex As obj_PADC_PersonIndex
    Dim prepared As Collection
    Dim preparedPerson As obj_PADC_PreparedPerson
    Dim matches As Collection
    Dim intervals As Collection
    Dim classifiedIntervals As Collection
    Dim interval As Variant
    Dim eventDecisions As Object
    Dim excluded As Boolean
    Dim matchRow As Variant
    Dim rowIndex As Long
    Dim cutoff As Long
    Dim startDay As Long
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
    m_runContext.BeginStep m_configuration.GetText("progress.Events")
    Set events = m_dataSource.ReadConfiguredSource("source.events", vbNullString, _
        VBA.Array("TaxId", "FullName", "Event", "Departure", "Arrival"))
    m_runContext.CompleteStep
    Set dates = private_IndexDates(orders, cutoff)
    Set seen = VBA.CreateObject(DICTIONARY_PROG_ID)
    Set prepared = New Collection
    personnelValues = personnel.Values
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
            latestDeparture = 0
            openStatus = vbNullString
            errorText = vbNullString
            Set intervals = New Collection
            If VBA.Len(taxId) = 0 Or VBA.Len(fullName) = 0 Then
                errorText = m_configuration.GetText("text.PersonIdentityMissing")
            ElseIf seen.Exists(taxId) Then
                m_configuration.Fail context & m_configuration.GetText("text.DuplicateTaxId") & taxId
            Else
                seen.Add taxId, True
                If dates.Exists("I|" & taxId) Then
                    startDay = dates("I|" & taxId)
                ElseIf dates.Exists("N|" & fullName) Then
                    startDay = dates("N|" & fullName)
                End If
                If startDay <= 0 Then
                    startDay = 0
                    errorText = m_configuration.GetText("text.DisposalDateMissing")
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
                            If arrival > cutoff And m_validation.MatchText(eventValues(matchRow, 1)) = taxId Then
                                If departure > latestDeparture Or (departure = latestDeparture And VBA.StrComp(eventName, openStatus, vbBinaryCompare) > 0) Then
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
            Set preparedPerson = New obj_PADC_PreparedPerson
            If Not preparedPerson.Initialize(fullName, taxId, startDay, classifiedIntervals, errorText, rank, status, openStatus) Then
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
Private Function private_IndexDates( _
    ByVal orders As obj_DataTable, _
    ByVal cutoff As Long _
) As Object
    Dim dates As Object
    Dim nameOwners As Object
    Dim values As Variant
    Dim actionName As String
    Dim taxId As String
    Dim fullName As String
    Dim day As Long
    Dim rowIndex As Long

    Set dates = VBA.CreateObject(DICTIONARY_PROG_ID)
    dates.CompareMode = vbTextCompare
    Set nameOwners = VBA.CreateObject(DICTIONARY_PROG_ID)
    nameOwners.CompareMode = vbTextCompare
    actionName = m_configuration.GetText("source.orders.action")
    values = orders.Values
    m_runContext.BeginStep m_configuration.GetText("progress.OrderIndex"), orders.RowCount
    For rowIndex = 1 To orders.RowCount
        If VBA.StrComp(m_validation.MatchText(values(rowIndex, 3)), actionName, vbTextCompare) = 0 Then
            taxId = m_validation.MatchText(values(rowIndex, 1))
            fullName = m_validation.MatchText(values(rowIndex, 2))
            day = private_OrderDay(values(rowIndex, 4), cutoff)
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
        m_runContext.ReportProgress rowIndex
    Next rowIndex
    m_runContext.CompleteStep
    Set private_IndexDates = dates
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