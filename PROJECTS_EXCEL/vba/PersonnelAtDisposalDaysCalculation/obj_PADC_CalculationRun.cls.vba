VERSION 1.0 CLASS
BEGIN
  MultiUse = -1
END
Attribute VB_Name = "obj_PADC_CalculationRun"
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
Private m_validation As obj_PADC_Validation
Private m_dataSource As obj_PADC_DataSource
Private m_parameters As obj_PADC_Parameters
Private m_eventPreparation As obj_PADC_EventPreparation
Private m_resultWriter As obj_PADC_ResultWriter
Private m_calculationService As obj_PADC_CalculationService
Private m_running As Boolean
Private m_configurationReady As Boolean
Private m_disposeRequested As Boolean
Private m_exportValues As Variant
Private m_exportEndDay As Long
Private Const CANCEL_ERROR As Long = vbObjectError + 2101

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
    ByVal pageBase As obj_PageBase _
) As Boolean
    If m_isInitialized Or m_isDisposed Then
        Exit Function
    End If
    If pageBase Is Nothing Then
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
        m_disposeRequested = True
        Me.RequestCancel
        Exit Sub
    End If
    m_isInitialized = False
    m_isDisposed = True
    m_exportValues = Empty
    private_DisposeServices
    Set m_pageBase = Nothing
End Sub

Public Sub Calculate()
    Dim parameterSheet As Worksheet
    Dim prepared As Collection
    Dim errorText As String
    Dim errorNumber As Long

    private_EnsureReady
    If m_running Then
        private_SetActionStatus m_configuration.GetText("text.AlreadyRunning"), True
        Exit Sub
    End If
    On Error GoTo Failed
    m_running = True
    m_exportValues = Empty
    If Not private_InitializeServices() Then
        GoTo Cleanup
    End If
    m_runContext.ClearLog
    private_SetActionStatus m_configuration.GetText("status.Loading")
    m_runContext.BeginStep m_configuration.GetText("progress.Parameters")
    m_runContext.LogDebug m_configuration.GetText("text.ParameterLoadingStarted")
    If Not TypeOf Application.ActiveSheet Is Worksheet Then
        m_configuration.Fail m_configuration.GetText("text.SelectToolSheet")
    End If
    Set parameterSheet = Application.ActiveSheet
    If Not parameterSheet.Parent Is ThisWorkbook Then
        m_configuration.Fail m_configuration.GetText("text.WrongWorkbook")
    End If
    m_runContext.CheckCancel 0
    m_runContext.LogDebug m_configuration.GetText("text.ParameterWorksheet") & parameterSheet.Name
    Set m_parameters = New obj_PADC_Parameters
    If Not m_parameters.Initialize(m_pageBase, m_configuration, m_validation, m_runContext, parameterSheet) Then
        m_configuration.Fail m_configuration.GetText("text.TargetPrepareFailed")
    End If
    m_runContext.CheckCancel 0
    m_runContext.CompleteStep
    m_runContext.LogDebug m_configuration.GetText("text.ParameterLoadingCompleted")
    m_runContext.LogStage "PrepareEvents.Started"
    Set prepared = m_eventPreparation.Prepare(m_parameters)
    m_runContext.LogStage "PrepareEvents.Completed"
    m_runContext.LogStage "CalculateAndPublish.Started"
    private_SetActionStatus m_configuration.GetText("status.Calculating")
    private_CalculatePreparedPeople prepared, m_parameters.EndDay, parameterSheet
    m_runContext.LogStage "CalculateAndPublish.Completed"
Cleanup:
    ex_Core.fn_Diagnostic_WriteLog "PADC_STAGE | Name=RunCleanup.Started"
    ex_Core.fn_Diagnostic_Flush
    private_DisposeServices
    ex_Core.fn_Diagnostic_WriteLog "PADC_STAGE | Name=RunCleanup.Completed"
    ex_Core.fn_Diagnostic_Flush
    m_running = False
    If m_disposeRequested Then
        Me.Dispose
    End If
    Exit Sub
Failed:
    errorText = VBA.Err.Description
    errorNumber = VBA.Err.Number
    If Not m_runContext Is Nothing Then
        m_runContext.LogFailure errorNumber, errorText
    Else
        ex_Core.fn_Diagnostic_WriteLog "PADC_FAILED | Number=" & errorNumber & " | " & errorText
    End If
    If errorNumber = CANCEL_ERROR Then
        private_SetActionStatus m_configuration.GetText("status.Cancelled")
    ElseIf m_configurationReady Then
        private_SetActionStatus m_configuration.GetText("text.OperationStopped") & errorText, True
    Else
        private_SetActionStatus errorText, True
    End If
    Resume Cleanup
End Sub

Public Sub RequestCancel()
    private_EnsureReady
    If m_running And Not m_runContext Is Nothing Then
        m_runContext.RequestCancel
        private_SetActionStatus m_configuration.GetText("status.Cancelling")
    End If
End Sub

Public Sub ExportResult()
    Dim configuration As obj_PADC_Configuration
    Dim exporter As obj_PADC_ResultExport
    Dim errorText As String

    private_EnsureReady
    If m_running Then
        Exit Sub
    End If
    Set configuration = New obj_PADC_Configuration
    If Not configuration.Initialize(errorText) Then
        private_SetActionStatus errorText, True
        Exit Sub
    End If
    On Error GoTo Failed
    If VBA.IsEmpty(m_exportValues) Then
        configuration.Fail configuration.GetText("text.ExportEmpty")
    End If
    Set exporter = New obj_PADC_ResultExport
    If Not exporter.Initialize(configuration) Then
        configuration.Fail configuration.GetText("text.TargetPrepareFailed")
    End If
    exporter.CreateWorkbook m_exportValues, m_exportEndDay
    private_SetActionStatus configuration.GetText("status.Exported")
Cleanup:
    If Not exporter Is Nothing Then
        exporter.Dispose
    End If
    configuration.Dispose
    Exit Sub
Failed:
    errorText = VBA.Err.Description
    private_SetActionStatus configuration.GetText("text.ExportFailed") & errorText, True
    Resume Cleanup
End Sub

' //
' // Private
' //
Private Sub private_SetActionStatus( _
    ByVal message As String, _
    Optional ByVal isError As Boolean = False _
)
    Dim updates As Object
    Dim diagnostic As String

    ex_Core.fn_StatusBar_Show message
    Set updates = VBA.CreateObject("Scripting.Dictionary")
    updates("Form.ActionStatus") = message
    updates("Form.ActionStatusStyle") = "statusNormal"
    If isError Then
        updates("Form.ActionStatusStyle") = "statusError"
    End If
    If Not m_pageBase.BindingContext.TryApplyValues(updates, diagnostic) Then
        ex_Core.fn_Diagnostic_WriteLog diagnostic
    End If
End Sub

Private Function private_InitializeServices() As Boolean
    Dim diagnostic As String

    Set m_configuration = New obj_PADC_Configuration
    If Not m_configuration.Initialize(diagnostic) Then
        private_SetActionStatus diagnostic, True
        Exit Function
    End If
    m_configurationReady = True
    Set m_runContext = New obj_PADC_RunContext
    If Not m_runContext.Initialize(m_configuration) Then
        m_configuration.Fail m_configuration.GetText("text.TargetPrepareFailed")
    End If
    Set m_validation = New obj_PADC_Validation
    If Not m_validation.Initialize(m_configuration) Then
        m_configuration.Fail m_configuration.GetText("text.TargetPrepareFailed")
    End If
    Set m_dataSource = New obj_PADC_DataSource
    If Not m_dataSource.Initialize(m_configuration, m_runContext) Then
        m_configuration.Fail m_configuration.GetText("text.TargetPrepareFailed")
    End If
    Set m_eventPreparation = New obj_PADC_EventPreparation
    If Not m_eventPreparation.Initialize(m_configuration, m_dataSource, m_validation, m_runContext) Then
        m_configuration.Fail m_configuration.GetText("text.TargetPrepareFailed")
    End If
    Set m_resultWriter = New obj_PADC_ResultWriter
    If Not m_resultWriter.Initialize(m_pageBase, m_configuration) Then
        m_configuration.Fail m_configuration.GetText("text.TargetPrepareFailed")
    End If
    Set m_calculationService = New obj_PADC_CalculationService
    If Not m_calculationService.Initialize(m_configuration.ThresholdDays, _
            m_configuration.GetText("event.present"), _
            m_configuration.GetText("format.periodRangeSeparator")) Then
        m_configuration.Fail m_configuration.GetText("text.TargetPrepareFailed")
    End If
    private_InitializeServices = True
End Function

Private Sub private_DisposeServices()
    m_configurationReady = False
    If Not m_calculationService Is Nothing Then
        m_calculationService.Dispose
    End If
    Set m_calculationService = Nothing
    If Not m_resultWriter Is Nothing Then
        m_resultWriter.Dispose
    End If
    Set m_resultWriter = Nothing
    If Not m_eventPreparation Is Nothing Then
        m_eventPreparation.Dispose
    End If
    Set m_eventPreparation = Nothing
    If Not m_parameters Is Nothing Then
        m_parameters.Dispose
    End If
    Set m_parameters = Nothing
    If Not m_dataSource Is Nothing Then
        m_dataSource.Dispose
    End If
    Set m_dataSource = Nothing
    If Not m_validation Is Nothing Then
        m_validation.Dispose
    End If
    Set m_validation = Nothing
    If Not m_runContext Is Nothing Then
        m_runContext.Dispose
    End If
    Set m_runContext = Nothing
    If Not m_configuration Is Nothing Then
        m_configuration.Dispose
    End If
    Set m_configuration = Nothing
End Sub

Private Sub private_CalculatePreparedPeople( _
    ByVal prepared As Collection, _
    ByVal lastDay As Long, _
    ByVal parameterSheet As Worksheet _
)
    Dim names() As Variant
    Dim ids() As Variant
    Dim starts() As Variant
    Dim periods() As Variant
    Dim days() As Variant
    Dim thresholds() As Variant
    Dim notes() As Variant
    Dim note As String
    Dim issueCount As Long
    Dim preparedPerson As obj_PADC_PreparedPerson
    Dim rows As Collection
    Dim periodsText As String
    Dim thresholdDate As Variant
    Dim count As Long
    Dim i As Long
    Dim exportValues() As Variant

    ReDim names(1 To prepared.Count, 1 To 1)
    ReDim ids(1 To prepared.Count, 1 To 1)
    ReDim starts(1 To prepared.Count, 1 To 1)
    ReDim periods(1 To prepared.Count, 1 To 1)
    ReDim days(1 To prepared.Count, 1 To 1)
    ReDim thresholds(1 To prepared.Count, 1 To 1)
    ReDim notes(1 To prepared.Count, 1 To 1)
    ReDim exportValues(1 To prepared.Count, 1 To 15)
    m_runContext.BeginStep m_configuration.GetText("progress.Calculation"), prepared.Count
    For i = 1 To prepared.Count
        Set preparedPerson = prepared(i)
        exportValues(i, 1) = preparedPerson.Rank
        exportValues(i, 2) = preparedPerson.FullName
        exportValues(i, 3) = preparedPerson.TaxId
        exportValues(i, 4) = preparedPerson.Status
        exportValues(i, 5) = preparedPerson.OpenStatus
        If preparedPerson.StartDay > 0 Then
            exportValues(i, 6) = VBA.CDate(preparedPerson.StartDay)
        End If
        names(i, 1) = preparedPerson.FullName
        ids(i, 1) = preparedPerson.TaxId
        starts(i, 1) = exportValues(i, 6)
        note = preparedPerson.Note
        If VBA.Len(preparedPerson.ValidationError) > 0 Then
            If VBA.Len(note) > 0 Then
                note = note & " "
            End If
            note = note & preparedPerson.ValidationError
        Else
            Set rows = preparedPerson.Intervals
            count = count + 1
            days(i, 1) = m_calculationService.Calculate(rows, VBA.CLng(preparedPerson.StartDay), lastDay, _
                periodsText, thresholdDate, m_runContext)
            periods(i, 1) = periodsText
            thresholds(i, 1) = thresholdDate
            exportValues(i, 7) = days(i, 1)
            exportValues(i, 8) = thresholdDate
        End If
        notes(i, 1) = note
        exportValues(i, 15) = note
        If VBA.Len(note) > 0 Then
            issueCount = issueCount + 1
            m_runContext.LogWarning note
        End If
        m_runContext.ReportProgress i
    Next i
    m_runContext.CompleteStep
    m_runContext.CheckCancel 0
    m_runContext.LogDebug m_configuration.GetText("text.WritingResult") & count & _
        m_configuration.GetText("text.Rows")
    m_runContext.BeginStep m_configuration.GetText("progress.Publish")
    m_runContext.LogStage "PublishResults.Started", "Rows=" & count & " | Failures=" & issueCount
    m_resultWriter.Publish names, ids, starts, periods, days, thresholds, notes, prepared.Count, lastDay, parameterSheet
    m_exportValues = exportValues
    m_exportEndDay = lastDay
    m_runContext.CompleteStep
    m_runContext.LogStage "PublishResults.Completed"
    m_runContext.LogDebug m_configuration.GetText("text.CalculationCompletedPeople") & count & _
        m_configuration.GetText("text.SkippedPeople") & issueCount
    private_SetActionStatus VBA.Replace$(VBA.Replace$(m_configuration.GetText("status.Completed"), _
        "{count}", VBA.CStr(count)), "{errors}", VBA.CStr(issueCount)), issueCount > 0
End Sub

Private Sub private_EnsureReady()
    If Not m_isInitialized Or m_isDisposed Then
        VBA.Err.Raise vbObjectError + 2100, "obj_PADC_CalculationRun", "Service is not initialized."
    End If
End Sub