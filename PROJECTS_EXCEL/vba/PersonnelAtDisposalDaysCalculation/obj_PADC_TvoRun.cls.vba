VERSION 1.0 CLASS
BEGIN
  MultiUse = -1
END
Attribute VB_Name = "obj_PADC_TvoRun"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private m_isInitialized As Boolean
Private m_isDisposed As Boolean

Private m_pageBase As obj_PageBase
Private m_configuration As obj_PADC_Configuration
Private m_runContext As obj_PADC_RunContext
Private m_dataSource As obj_PADC_DataSource
Private m_validation As obj_PADC_Validation
Private m_resultTable As obj_UiRawTable
Private m_values As Variant
Private m_headers As Variant
Private m_endDay As Long
Private m_running As Boolean
Private m_accessDates As Object

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
Public Function Initialize(ByVal pageBase As obj_PageBase) As Boolean
    If m_isInitialized Or m_isDisposed Or pageBase Is Nothing Then
        Exit Function
    End If
    Set m_pageBase = pageBase
    m_isInitialized = True
    Initialize = True
End Function

Public Sub Dispose()
    If m_isDisposed Then
        Exit Sub
    End If
    If m_running Then
        Me.RequestCancel
        Exit Sub
    End If
    private_DisposeServices
    If Not m_resultTable Is Nothing Then
        m_resultTable.Dispose
    End If
    Set m_resultTable = Nothing
    Set m_pageBase = Nothing
    If Not m_configuration Is Nothing Then
        m_configuration.Dispose
    End If
    Set m_configuration = Nothing
    m_values = Empty
    m_isInitialized = False
    m_isDisposed = True
End Sub

Public Sub Calculate()
    Dim diagnostic As String
    Dim data As obj_DataTable
    Dim events As obj_DataTable
    Dim assignments As Variant
    Dim absences As Object
    Dim rows As Collection
    Dim result As Variant
    Dim values() As Variant
    Dim rowIndex As Long
    Dim columnIndex As Long
    Dim cutoff As Long
    Dim value As Variant
    Dim boundObject As Object
    Dim isObject As Boolean
    Dim errorText As String
    Dim previousStatusBar As Variant

    private_EnsureReady
    If m_running Then
        Exit Sub
    End If
    On Error GoTo Failed
    previousStatusBar = Application.StatusBar
    m_running = True
    m_values = Empty
    private_DisposeServices
    If Not m_configuration Is Nothing Then
        m_configuration.Dispose
    End If
    Set m_configuration = New obj_PADC_Configuration
    If Not m_configuration.Initialize(diagnostic, True) Then
        VBA.Err.Raise vbObjectError + 2100, "obj_PADC_TvoRun", diagnostic
    End If
    Set m_validation = New obj_PADC_Validation
    Set m_runContext = New obj_PADC_RunContext
    Set m_dataSource = New obj_PADC_DataSource
    If Not m_validation.Initialize(m_configuration) Then
        GoTo InitializationFailed
    End If
    If Not m_runContext.Initialize(m_configuration, 5) Then
        GoTo InitializationFailed
    End If
    If Not m_dataSource.Initialize(m_configuration, m_runContext) Then
        GoTo InitializationFailed
    End If
    If Not m_pageBase.BindingContext.TryGetValue("Form", "EndDate", value, boundObject, isObject) Or isObject Then
        GoTo InitializationFailed
    End If
    cutoff = m_validation.ReadDay(value, ThisWorkbook.Date1904, "Form.EndDate") - 1
    m_endDay = cutoff + 1
    If Not VBA.IsNumeric(m_configuration.GetText("tvo.days")) Then
        private_Invalid "tvo.days"
    End If
    If VBA.CDbl(m_configuration.GetText("tvo.days")) <> 30 Then
        private_Invalid "tvo.days"
    End If

    ' Read both sources from one cached snapshot of the Movement workbook.
    m_runContext.BeginStep m_configuration.GetText("progress.Tvo")
    Set data = m_dataSource.ReadConfiguredSource("source.tvo", _
        VBA.CreateObject("WScript.Shell").ExpandEnvironmentStrings(m_configuration.GetText("source.events.path")), _
        VBA.Array("Rank", "FullName", "TaxId", "Position", "StartOrder", "StartDate", "EndDate", "OwnPosition", "AccessOrder", "Duration", "EndOrder", "PlannedEndDate"))
    m_runContext.CompleteStep
    m_runContext.BeginStep m_configuration.GetText("progress.Events")
    Set events = m_dataSource.ReadConfiguredSource("source.events", vbNullString, _
        VBA.Array("TaxId", "FullName", "Event", "Departure", "Arrival"))
    m_runContext.CompleteStep
    m_runContext.BeginStep m_configuration.GetText("progress.EventIndex"), events.RowCount
    Set absences = private_IndexAbsences(events)
    m_runContext.CompleteStep
    assignments = data.Values
    Set rows = New Collection
    m_runContext.BeginStep m_configuration.GetText("progress.Calculation"), data.RowCount
    For rowIndex = 1 To data.RowCount
        If VBA.Len(private_Text(assignments(rowIndex, 5))) > 0 Then
            result = private_CalculateRow(assignments, rowIndex, absences, cutoff)
            rows.Add result
        End If
        m_runContext.ReportProgress rowIndex
    Next rowIndex
    m_runContext.CompleteStep
    m_runContext.BeginStep m_configuration.GetText("progress.Publish")
    m_headers = VBA.Array(m_configuration.GetText("tvo.header.Rank"), m_configuration.GetText("header.FullName"), _
        m_configuration.GetText("header.TaxId"), m_configuration.GetText("tvo.header.OwnPosition"), m_configuration.GetText("tvo.header.Position"), _
        m_configuration.GetText("tvo.header.Duration"), m_configuration.GetText("tvo.header.AccessOrder"), m_configuration.GetText("tvo.header.StartOrder"), _
        m_configuration.GetText("tvo.header.StartDate"), m_configuration.GetText("tvo.header.StreakStart"), _
        m_configuration.GetText("tvo.header.ThresholdDate"), m_configuration.GetText("tvo.header.Days"), _
        m_configuration.GetText("tvo.header.PlannedEndDate"), m_configuration.GetText("tvo.header.EndDate"), _
        m_configuration.GetText("tvo.header.EndOrder"), m_configuration.GetText("header.Periods"))
    Set m_resultTable = New obj_UiRawTable
    If rows.Count = 0 Then
        If Not m_resultTable.InitializeEmpty(m_headers) Then
            GoTo InitializationFailed
        End If
    Else
        ReDim values(1 To rows.Count, 1 To 16)
        For rowIndex = 1 To rows.Count
            result = rows(rowIndex)
            For columnIndex = 1 To 16
                values(rowIndex, columnIndex) = result(columnIndex - 1)
            Next columnIndex
        Next rowIndex
        If Not m_resultTable.Initialize(values, m_headers) Then
            GoTo InitializationFailed
        End If
        m_values = values
    End If
    If Not m_pageBase.BindingContext.SetObject("Data", "CalculationRows", m_resultTable) Then
        GoTo InitializationFailed
    End If
    If Not m_pageBase.UpdatePage() Then
        GoTo InitializationFailed
    End If
    m_runContext.CompleteStep
    private_Status VBA.Replace$(m_configuration.GetText("tvo.text.Complete"), "{count}", VBA.CStr(rows.Count)), False
Cleanup:
    If Not data Is Nothing Then
        data.Dispose
    End If
    If Not events Is Nothing Then
        events.Dispose
    End If
    private_DisposeServices
    Application.StatusBar = previousStatusBar
    m_running = False
    Exit Sub
InitializationFailed:
    m_configuration.Fail m_configuration.GetText("text.TargetPrepareFailed")
Failed:
    errorText = VBA.Err.Description
    m_values = Empty
    private_Status errorText, True
    ex_WindowsUi.fn_ShowMessage errorText, VBA.vbExclamation
    Resume Cleanup
End Sub

Public Sub ExportResult()
    Dim exporter As obj_PADC_ResultExport
    Dim diagnostic As String
    Dim errorText As String

    private_EnsureReady
    If m_running Then
        Exit Sub
    End If
    On Error GoTo Failed
    If m_configuration Is Nothing Then
        Set m_configuration = New obj_PADC_Configuration
        If Not m_configuration.Initialize(diagnostic, True) Then
            VBA.Err.Raise vbObjectError + 2100, "obj_PADC_TvoRun", diagnostic
        End If
    End If
    If VBA.IsEmpty(m_values) Then
        m_configuration.Fail m_configuration.GetText("text.ExportEmpty")
    End If
    Set exporter = New obj_PADC_ResultExport
    If Not exporter.Initialize(m_configuration) Then
        m_configuration.Fail m_configuration.GetText("text.TargetPrepareFailed")
    End If
    exporter.CreateTvoWorkbook m_values, m_headers, m_endDay
Cleanup:
    If Not exporter Is Nothing Then
        exporter.Dispose
    End If
    Exit Sub
Failed:
    errorText = VBA.Err.Description
    private_Status errorText, True
    ex_WindowsUi.fn_ShowMessage errorText, VBA.vbExclamation
    Resume Cleanup
End Sub

Public Sub RequestCancel()
    private_EnsureReady
    If Not m_runContext Is Nothing Then
        m_runContext.RequestCancel
    End If
End Sub

' //
' // Private
' //
Private Function private_IndexAbsences(ByVal data As obj_DataTable) As Object
    Dim index As Object
    Dim intervals As Collection
    Dim values As Variant
    Dim taxId As String
    Dim departure As Long
    Dim arrival As Long
    Dim rowIndex As Long
    Dim position As Long
    Dim inserted As Boolean
    Dim interval As Variant
    Dim ignoredEvents As Object
    Dim ignoredEvent As Variant
    Dim eventName As String

    Set index = VBA.CreateObject("Scripting.Dictionary")
    Set ignoredEvents = VBA.CreateObject("Scripting.Dictionary")
    ignoredEvents.CompareMode = vbTextCompare
    For Each ignoredEvent In VBA.Split(m_configuration.GetText("tvo.ignoredEvents"), "|")
        eventName = private_Text(ignoredEvent)
        If VBA.Len(eventName) = 0 Then
            private_Invalid "tvo.ignoredEvents"
        End If
        ignoredEvents(eventName) = True
    Next ignoredEvent
    values = data.Values
    For rowIndex = 1 To data.RowCount
        m_runContext.ReportProgress rowIndex
        eventName = private_Text(values(rowIndex, 3))
        If ignoredEvents.Exists(eventName) Then
            GoTo NextAbsence
        End If
        taxId = private_Text(values(rowIndex, 1))
        If VBA.Len(taxId) > 0 Or VBA.Len(private_Text(values(rowIndex, 2))) > 0 Then
            taxId = m_validation.RequiredText(values(rowIndex, 1), "source.events #" & rowIndex)
            departure = m_validation.ReadDay(values(rowIndex, 4), False, "source.events #" & rowIndex)
            arrival = 2958466
            If VBA.Len(private_Text(values(rowIndex, 5))) > 0 Then
                arrival = m_validation.ReadDay(values(rowIndex, 5), False, "source.events #" & rowIndex)
                If arrival < departure Then
                    private_Invalid "source.events #" & rowIndex
                End If
            End If
            If Not index.Exists(taxId) Then
                Set intervals = New Collection
                index.Add taxId, intervals
            End If
            Set intervals = index(taxId)
            inserted = False
            For position = 1 To intervals.Count
                interval = intervals(position)
                If departure < interval(0) Then
                    intervals.Add VBA.Array(departure, arrival, eventName), Before:=position
                    inserted = True
                    Exit For
                End If
            Next position
            If Not inserted Then
                intervals.Add VBA.Array(departure, arrival, eventName)
            End If
        End If
NextAbsence:
    Next rowIndex
    Set private_IndexAbsences = index
End Function

Private Function private_CalculateRow( _
    ByRef assignments As Variant, _
    ByVal rowIndex As Long, _
    ByVal absences As Object, _
    ByVal cutoff As Long _
) As Variant
    Dim taxId As String
    Dim fullName As String
    Dim position As String
    Dim startDay As Long
    Dim endDay As Long
    Dim limit As Long
    Dim cursor As Long
    Dim threshold As Long
    Dim streakDays As Long
    Dim interval As Variant
    Dim intervals As Collection
    Dim result(0 To 15) As Variant
    Dim absenceDays As Long
    Dim coveredThrough As Long
    Dim absenceStart As Long
    Dim absenceEnd As Long
    Dim plannedEnd As Long
    Dim periodStart As Long
    Dim periodEnd As Long
    Dim periodText As String
    Dim periods As String

    taxId = m_validation.RequiredText(assignments(rowIndex, 3), "source.tvo #" & rowIndex)
    fullName = m_validation.RequiredText(assignments(rowIndex, 2), "source.tvo #" & rowIndex)
    position = m_validation.RequiredText(assignments(rowIndex, 4), "source.tvo #" & rowIndex)
    startDay = m_validation.ReadDay(assignments(rowIndex, 6), False, "source.tvo #" & rowIndex)
    endDay = 2958466
    If VBA.Len(private_Text(assignments(rowIndex, 7))) > 0 Then
        endDay = m_validation.ReadDay(assignments(rowIndex, 7), False, "source.tvo #" & rowIndex)
        If endDay < startDay Then
            private_Invalid "source.tvo #" & rowIndex
        End If
    End If
    limit = cutoff + 1
    If endDay < limit Then
        limit = endDay
    End If
    cursor = startDay
    coveredThrough = startDay

    ' Sorted overlapping absences advance one cursor and never double-count a reset.
    If absences.Exists(taxId) And limit > startDay Then
        Set intervals = absences(taxId)
        For Each interval In intervals
            m_runContext.CheckCancel rowIndex
            If interval(0) >= limit Then
                Exit For
            End If
            absenceStart = interval(0)
            If absenceStart < coveredThrough Then
                absenceStart = coveredThrough
            End If
            absenceEnd = interval(1)
            If absenceEnd > limit Then
                absenceEnd = limit
            End If
            If absenceEnd > absenceStart Then
                absenceDays = absenceDays + absenceEnd - absenceStart
                coveredThrough = absenceEnd
            End If
            If interval(1) > startDay Or (interval(0) = interval(1) And interval(0) >= startDay) Then
                periodStart = interval(0)
                If periodStart < startDay Then
                    periodStart = startDay
                End If
                periodEnd = interval(1) - 1
                If periodEnd >= limit Then
                    periodEnd = limit - 1
                End If
                If periodEnd < periodStart Then
                    periodEnd = periodStart
                End If
                periodText = m_configuration.GetText("tvo.text.AbsencePeriod")
                periodText = VBA.Replace$(periodText, "{event}", VBA.CStr(interval(2)))
                periodText = VBA.Replace$(periodText, "{from}", VBA.Format$(VBA.CDate(periodStart), m_configuration.GetText("format.date")))
                periodText = VBA.Replace$(periodText, "{through}", VBA.Format$(VBA.CDate(periodEnd), m_configuration.GetText("format.date")))
                If VBA.Len(periods) > 0 Then
                    periods = periods & vbLf
                End If
                periods = periods & periodText
            End If
            If interval(1) > cursor Then
                If interval(0) > cursor And threshold = 0 Then
                    If interval(0) - cursor >= 30 Then
                        threshold = cursor + 29
                    End If
                End If
                cursor = interval(1)
            ElseIf interval(0) = interval(1) And interval(0) >= cursor Then
                ' A same-day departure still resets the counter on the return day.
                If threshold = 0 And interval(0) - cursor >= 30 Then
                    threshold = cursor + 29
                End If
                cursor = interval(1)
            End If
        Next interval
    End If
    If cursor < limit Then
        streakDays = limit - cursor
        result(9) = VBA.Format$(VBA.CDate(cursor), m_configuration.GetText("format.date"))
        If threshold = 0 And streakDays >= 30 Then
            threshold = cursor + 29
        End If
    End If
    If threshold = 0 Then
        If cursor < limit And endDay > cursor + 29 And cursor + 29 <= 2958465 Then
            threshold = cursor + 29
        End If
    End If
    result(0) = private_Text(assignments(rowIndex, 1))
    result(1) = fullName
    result(2) = taxId
    result(3) = m_validation.RequiredText(assignments(rowIndex, 8), "source.tvo.columns.OwnPosition #" & rowIndex)
    result(4) = position
    result(5) = private_Text(assignments(rowIndex, 10))
    result(6) = private_Text(assignments(rowIndex, 9))
    result(7) = private_Text(assignments(rowIndex, 5))
    result(8) = VBA.Format$(VBA.CDate(startDay), m_configuration.GetText("format.date"))
    If endDay < 2958466 Then
        result(13) = VBA.Format$(VBA.CDate(endDay), m_configuration.GetText("format.date"))
    End If
    result(11) = streakDays
    If threshold > 0 Then
        result(10) = VBA.Format$(VBA.CDate(threshold), m_configuration.GetText("format.date"))
    End If
    If VBA.Len(private_Text(assignments(rowIndex, 12))) > 0 Then
        plannedEnd = m_validation.ReadDay(assignments(rowIndex, 12), False, "source.tvo.columns.PlannedEndDate #" & rowIndex)
        If plannedEnd > 2958465 - absenceDays Then
            private_Invalid "source.tvo.columns.PlannedEndDate #" & rowIndex
        End If
        result(12) = VBA.Format$(VBA.CDate(plannedEnd + absenceDays), m_configuration.GetText("format.date"))
    End If
    result(14) = private_Text(assignments(rowIndex, 11))
    result(15) = periods
    private_CalculateRow = result
End Function

Private Function private_Text(ByVal value As Variant) As String
    If VBA.IsError(value) Then
        private_Invalid "source cell"
    End If
    If Not VBA.IsNull(value) And Not VBA.IsEmpty(value) Then
        private_Text = VBA.Trim$(VBA.Replace$(VBA.CStr(value), VBA.ChrW(160), " "))
    End If
End Function

Private Sub private_Invalid(ByVal context As String)
    m_configuration.Fail VBA.Replace$(m_configuration.GetText("tvo.text.InvalidInterval"), "{context}", context)
End Sub

Private Sub private_Status( _
    ByVal text As String, _
    ByVal isError As Boolean _
)
    Dim style As String

    style = "statusNormal"
    If isError Then
        style = "statusError"
    End If
    m_pageBase.BindingContext.SetValue "Form", "ActionStatusStyle", style
    m_pageBase.BindingContext.SetValue "Form", "ActionStatus", text
End Sub

Private Sub private_DisposeServices()
    Set m_accessDates = Nothing
    If Not m_dataSource Is Nothing Then
        m_dataSource.Dispose
    End If
    If Not m_validation Is Nothing Then
        m_validation.Dispose
    End If
    If Not m_runContext Is Nothing Then
        m_runContext.Dispose
    End If
    Set m_dataSource = Nothing
    Set m_validation = Nothing
    Set m_runContext = Nothing
End Sub

Private Sub private_ReadAccessOrders()
    Dim data As obj_DataTable
    Dim values As Variant
    Dim rowIndex As Long
    Dim number As String
    Dim day As Long
    Dim errorNumber As Long
    Dim errorText As String

    On Error GoTo Failed
    Set m_accessDates = VBA.CreateObject("Scripting.Dictionary")
    m_accessDates.CompareMode = vbTextCompare
    Set data = m_dataSource.ReadConfiguredSource("source.accessOrders", vbNullString, VBA.Array("Number", "Date"))
    values = data.Values
    For rowIndex = 1 To data.RowCount
        m_runContext.CheckCancel rowIndex
        number = private_OrderNumber(values(rowIndex, 1))
        If VBA.Len(number) > 0 Then
            day = m_validation.ReadDay(values(rowIndex, 2), False, "source.accessOrders #" & rowIndex)
            number = VBA.CStr(VBA.Year(VBA.CDate(day))) & "|" & number
            If m_accessDates.Exists(number) Then
                If m_accessDates(number) <> day Then
                    m_accessDates(number) = -1
                End If
            Else
                m_accessDates.Add number, day
            End If
        End If
    Next rowIndex
    data.Dispose
    Exit Sub
Failed:
    errorNumber = VBA.Err.Number
    errorText = VBA.Err.Description
    If Not data Is Nothing Then
        data.Dispose
    End If
    VBA.Err.Raise errorNumber, "obj_PADC_TvoRun.private_ReadAccessOrders", errorText
End Sub

Private Function private_OrderNumber(ByVal value As Variant) As String
    Dim text As String

    text = VBA.Replace$(private_Text(value), " ", vbNullString)
    If VBA.Len(text) > 0 And VBA.IsNumeric(text) Then
        text = VBA.CStr(VBA.CLng(text))
    End If
    private_OrderNumber = text
End Function

Private Sub private_EnsureReady()
    If Not m_isInitialized Or m_isDisposed Then
        VBA.Err.Raise vbObjectError + 2100, "obj_PADC_TvoRun", "Service is not initialized."
    End If
End Sub