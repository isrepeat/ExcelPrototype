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

Public Function Prepare( _
    ByVal params As ListObject, _
    ByVal lastDay As Long _
) As Collection
    Dim source As ListObject
    Dim roster As ListObject
    Dim people As Object
    Dim intervals As Object
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
    Dim firstDay As Long
    Dim arrival As Long
    Dim taxId As String
    Dim eventName As String
    Dim context As String
    Dim eventPersonIndex As obj_PADC_PersonIndex
    Dim rosterPersonIndex As obj_PADC_PersonIndex
    Dim prepared As Collection
    Dim preparedPerson As obj_PADC_PreparedPerson

    private_EnsureReady
    Set source = m_dataSource.FindSource()
    Set roster = m_dataSource.RequireTable(m_configuration.GetText("legacy.ROSTER_SHEET_NAME"), m_configuration.GetText("legacy.ROSTER_TABLE_NAME"), source.Parent.Parent)
    rosterTaxCol = m_dataSource.ColumnIndex(roster, m_configuration.GetText("legacy.ROSTER_COL_TAX_ID"))
    rosterNameCol = m_dataSource.ColumnIndex(roster, m_configuration.GetText("legacy.ROSTER_COL_NAME"))
    rosterData = m_dataSource.ReadQueryTable(roster)
    m_runContext.CheckCancel 0

    m_runContext.LogDebug m_configuration.GetText("legacy.MSG_RUKH_SOURCE") & source.Parent.Parent.Name & " / " & source.Parent.Name & " / " & source.Name

    pTax = m_dataSource.ColumnIndex(params, m_configuration.GetText("legacy.PARAM_COL_TAX_ID"))
    pName = m_dataSource.ColumnIndex(params, m_configuration.GetText("legacy.PARAM_COL_NAME"))
    pStart = m_dataSource.ColumnIndex(params, m_configuration.GetText("legacy.PARAM_COL_START"))
    pTrip = m_dataSource.ColumnIndex(params, m_configuration.GetText("legacy.PARAM_COL_TRIPS"))
    sourceNameCol = m_dataSource.ColumnIndex(source, m_configuration.GetText("legacy.SOURCE_COL_NAME"))
    taxCol = m_dataSource.ColumnIndex(source, m_configuration.GetText("legacy.SOURCE_COL_TAX_ID"))
    eventCol = m_dataSource.ColumnIndex(source, m_configuration.GetText("legacy.SOURCE_COL_EVENT"))
    fromCol = m_dataSource.ColumnIndex(source, m_configuration.GetText("legacy.SOURCE_COL_PERIOD_FROM"))
    toCol = m_dataSource.ColumnIndex(source, m_configuration.GetText("legacy.SOURCE_COL_PERIOD_TO"))
    If params.DataBodyRange Is Nothing Then
        m_configuration.Fail m_configuration.GetText("legacy.MSG_PARAMETERS_EMPTY")
    End If
    p = m_dataSource.ReadQueryTable(params)
    s = m_dataSource.ReadQueryTable(source)
    Set eventPersonIndex = New obj_PADC_PersonIndex
    If Not eventPersonIndex.Initialize(s, source.ListRows.Count, taxCol, sourceNameCol, _
            m_configuration, m_validation, m_runContext) Then
        m_configuration.Fail m_configuration.GetText("legacy.MSG_TARGET_PREPARE_FAILED")
    End If
    Set rosterPersonIndex = New obj_PADC_PersonIndex
    If Not rosterPersonIndex.Initialize(rosterData, roster.ListRows.Count, rosterTaxCol, rosterNameCol, _
            m_configuration, m_validation, m_runContext) Then
        m_configuration.Fail m_configuration.GetText("legacy.MSG_TARGET_PREPARE_FAILED")
    End If
    count = m_dataSource.CompactParameterRows(p, parameterRows)
    If count = 0 Then
        m_configuration.Fail m_configuration.GetText("legacy.MSG_PARAMETERS_EMPTY")
    End If
    m_runContext.LogDebug m_configuration.GetText("legacy.MSG_PEOPLE") & count & m_configuration.GetText("legacy.MSG_RUKH_EVENTS") & source.ListRows.Count
    ReDim personErrors(1 To count)
    ReDim starts(1 To count)
    ReDim trips(1 To count)
    ReDim resultNames(1 To count, 1 To 1)
    ReDim resultTaxIds(1 To count, 1 To 1)
    ReDim resultTrips(1 To count, 1 To 1)
    Set people = VBA.CreateObject(DICTIONARY_PROG_ID)
    Set intervals = VBA.CreateObject(DICTIONARY_PROG_ID)
    For i = 1 To count
        context = params.Parent.Name & m_configuration.GetText("legacy.MSG_WORKSHEET_ROW") & params.DataBodyRange.Row + parameterRows(i) - 1
        taxId = m_validation.RequiredText(p(i, pTax), context & m_configuration.GetText("legacy.MSG_TAX_ID"))
        If people.Exists(taxId) Then
            m_configuration.Fail context & m_configuration.GetText("legacy.MSG_DUPLICATE_TAX_ID") & taxId
        End If
        people.Add taxId, i
        starts(i) = m_validation.ReadDay(p(i, pStart), params.Parent.Parent.Date1904, context & m_configuration.GetText("legacy.MSG_TIME_POINT"))
        If starts(i) > lastDay Then
            m_configuration.Fail context & m_configuration.GetText("legacy.MSG_START_AFTER_END")
        End If
        trips(i) = m_validation.ReadFlag(p(i, pTrip), context)
        resultTrips(i, 1) = p(i, pTrip)
        resultTaxIds(i, 1) = taxId
        resultNames(i, 1) = m_validation.RequiredText(p(i, pName), context & m_configuration.GetText("legacy.MSG_FULL_NAME"))
        Set rows = New Collection
        intervals.Add taxId, rows
        m_runContext.CheckCancel i
    Next i
    For person = 1 To count
        m_runContext.CheckCancel 0
        taxId = resultTaxIds(person, 1)
        Set matches = eventPersonIndex.Find( _
            taxId, VBA.CStr(resultNames(person, 1)), matchError)
        If VBA.Len(matchError) > 0 Then
            personErrors(person) = matchError
            GoTo NextMatchedPerson
        End If
        If matches.Count = 0 Then
            Set rosterMatches = rosterPersonIndex.Find( _
                taxId, VBA.CStr(resultNames(person, 1)), matchError)
            If VBA.Len(matchError) > 0 Then
                personErrors(person) = matchError
                GoTo NextMatchedPerson
            End If
            If rosterMatches.Count = 0 Then
                personErrors(person) = m_configuration.GetText("legacy.MSG_PERSON_UNVERIFIED") & resultNames(person, 1) & " / " & taxId
                GoTo NextMatchedPerson
            End If
            resolvedTaxId = m_validation.MatchText(rosterData(rosterMatches(1), rosterTaxCol))
            If VBA.Len(resolvedTaxId) > 0 Then
                Set matches = eventPersonIndex.Find( _
            resolvedTaxId, VBA.CStr(resultNames(person, 1)), matchError)
                If VBA.Len(matchError) > 0 Then
                    personErrors(person) = matchError
                    GoTo NextMatchedPerson
                End If
            End If
        End If
        For Each rowNumber In matches
            m_runContext.CheckCancel VBA.CLng(rowNumber)
            context = source.Parent.Parent.Name & " / " & m_configuration.GetText("legacy.SOURCE_SHEET_NAME") & _
                m_configuration.GetText("legacy.MSG_WORKSHEET_ROW") & source.DataBodyRange.Row + VBA.CLng(rowNumber) - 1
            eventName = m_validation.RequiredText(s(VBA.CLng(rowNumber), eventCol), context & m_configuration.GetText("legacy.MSG_EVENT"))
            firstDay = m_validation.ReadDay(s(VBA.CLng(rowNumber), fromCol), source.Parent.Parent.Date1904, context & m_configuration.GetText("legacy.MSG_DEPARTURE"))
            If VBA.IsError(s(VBA.CLng(rowNumber), toCol)) Then
                m_configuration.Fail context & m_configuration.GetText("legacy.MSG_ARRIVAL_CELL_ERROR")
            End If
            If VBA.Len(VBA.Trim$(VBA.CStr(s(VBA.CLng(rowNumber), toCol)))) = 0 Then
                arrival = lastDay
            Else
                arrival = m_validation.ReadDay(s(VBA.CLng(rowNumber), toCol), source.Parent.Parent.Date1904, context & m_configuration.GetText("legacy.MSG_ARRIVAL"))
                If arrival < firstDay Then
                    m_configuration.Fail context & m_configuration.GetText("legacy.MSG_ARRIVAL_BEFORE_DEPARTURE")
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
                rows.Add VBA.Array(firstDay, arrival, eventName, m_validation.IsExcluded(eventName, Not trips(person)))
            End If
        Next rowNumber
NextMatchedPerson:
    Next person
    Set prepared = New Collection
    For i = 1 To count
        m_runContext.CheckCancel i
        Set preparedPerson = New obj_PADC_PreparedPerson
        Set rows = intervals(VBA.CStr(resultTaxIds(i, 1)))
        If Not preparedPerson.Initialize(VBA.CStr(resultNames(i, 1)), VBA.CStr(resultTaxIds(i, 1)), _
                starts(i), resultTrips(i, 1), rows, personErrors(i)) Then
            m_configuration.Fail m_configuration.GetText("legacy.MSG_TARGET_PREPARE_FAILED")
        End If
        prepared.Add preparedPerson
    Next i
    eventPersonIndex.Dispose
    rosterPersonIndex.Dispose
    Set Prepare = prepared
End Function

' //
' // Private
' //
Private Sub private_EnsureReady()
    If Not m_isInitialized Or m_isDisposed Then
        VBA.Err.Raise vbObjectError + 2100, "obj_PADC_EventPreparation", "Service is not initialized."
    End If
End Sub