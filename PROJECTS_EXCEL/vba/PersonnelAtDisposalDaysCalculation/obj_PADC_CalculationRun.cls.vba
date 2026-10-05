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
    private_DisposeServices
    Set m_pageBase = Nothing
End Sub

Public Sub Calculate()
    Dim parameterSheet As Worksheet
    Dim parameterTable As ListObject
    Dim prepared As Collection
    Dim oldStatus As Variant
    Dim errorText As String
    Dim errorNumber As Long

    private_EnsureReady
    If m_running Then
        ex_WindowsUi.fn_ShowMessage m_configuration.GetText("legacy.MSG_ALREADY_RUNNING"), vbExclamation
        Exit Sub
    End If
    oldStatus = Application.StatusBar
    On Error GoTo Failed
    m_running = True
    If Not private_InitializeServices() Then
        GoTo Cleanup
    End If
    m_runContext.ClearLog
    m_runContext.LogDebug m_configuration.GetText("legacy.MSG_PARAMETER_LOADING_STARTED")
    If Not TypeOf Application.ActiveSheet Is Worksheet Then
        m_configuration.Fail m_configuration.GetText("legacy.MSG_SELECT_TOOL_SHEET")
    End If
    Set parameterSheet = Application.ActiveSheet
    If Not parameterSheet.Parent Is ThisWorkbook Then
        m_configuration.Fail m_configuration.GetText("legacy.MSG_WRONG_WORKBOOK")
    End If
    m_runContext.CheckCancel 0
    m_runContext.LogDebug m_configuration.GetText("legacy.MSG_PARAMETER_WORKSHEET") & parameterSheet.Name
    Set m_parameters = New obj_PADC_Parameters
    If Not m_parameters.Initialize(m_pageBase, m_configuration, m_validation, m_runContext, parameterSheet) Then
        m_configuration.Fail m_configuration.GetText("legacy.MSG_TARGET_PREPARE_FAILED")
    End If
    m_runContext.CheckCancel 0
    m_runContext.LogDebug m_configuration.GetText("legacy.MSG_PARAMETER_LOADING_COMPLETED")
    m_runContext.LogStage "OpenInput.Started"
    Set parameterTable = m_dataSource.OpenParameterTable(m_parameters)
    m_runContext.LogStage "OpenInput.Completed"
    m_runContext.LogStage "PrepareEvents.Started"
    Set prepared = m_eventPreparation.Prepare(parameterTable, m_parameters.EndDay)
    m_runContext.LogStage "PrepareEvents.Completed"
    m_runContext.LogStage "CalculateAndPublish.Started"
    private_CalculatePreparedPeople prepared, m_parameters.EndDay
    m_runContext.LogStage "CalculateAndPublish.Completed"
Cleanup:
    ex_Core.fn_Diagnostic_WriteLog "PADC_STAGE | Name=RunCleanup.Started"
    ex_Core.fn_Diagnostic_Flush
    private_DisposeServices
    ex_Core.fn_Diagnostic_WriteLog "PADC_STAGE | Name=RunCleanup.Completed"
    ex_Core.fn_Diagnostic_Flush
    Application.StatusBar = oldStatus
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
        ex_WindowsUi.fn_ShowMessage m_configuration.GetText("legacy.MSG_OPERATION_CANCELLED"), vbInformation
    ElseIf m_configurationReady Then
        ex_WindowsUi.fn_ShowMessage m_configuration.GetText("legacy.MSG_OPERATION_STOPPED") & errorText, vbExclamation
    Else
        ex_WindowsUi.fn_ShowMessage errorText, vbExclamation
    End If
    Resume Cleanup
End Sub

Public Sub RequestCancel()
    private_EnsureReady
    If m_running And Not m_runContext Is Nothing Then
        m_runContext.RequestCancel
    End If
End Sub

' //
' // Private
' //
Private Function private_InitializeServices() As Boolean
    Set m_configuration = New obj_PADC_Configuration
    If Not m_configuration.Initialize() Then
        Exit Function
    End If
    m_configurationReady = True
    Set m_runContext = New obj_PADC_RunContext
    If Not m_runContext.Initialize(m_configuration) Then
        m_configuration.Fail m_configuration.GetText("legacy.MSG_TARGET_PREPARE_FAILED")
    End If
    Set m_validation = New obj_PADC_Validation
    If Not m_validation.Initialize(m_configuration) Then
        m_configuration.Fail m_configuration.GetText("legacy.MSG_TARGET_PREPARE_FAILED")
    End If
    Set m_dataSource = New obj_PADC_DataSource
    If Not m_dataSource.Initialize(m_configuration, m_runContext) Then
        m_configuration.Fail m_configuration.GetText("legacy.MSG_TARGET_PREPARE_FAILED")
    End If
    Set m_eventPreparation = New obj_PADC_EventPreparation
    If Not m_eventPreparation.Initialize(m_configuration, m_dataSource, m_validation, m_runContext) Then
        m_configuration.Fail m_configuration.GetText("legacy.MSG_TARGET_PREPARE_FAILED")
    End If
    Set m_resultWriter = New obj_PADC_ResultWriter
    If Not m_resultWriter.Initialize(m_pageBase, m_configuration) Then
        m_configuration.Fail m_configuration.GetText("legacy.MSG_TARGET_PREPARE_FAILED")
    End If
    Set m_calculationService = New obj_PADC_CalculationService
    If Not m_calculationService.Initialize(m_configuration.ThresholdDays, _
            m_configuration.GetText("legacy.PRESENT_EVENT_NAME"), _
            m_configuration.GetText("format.periodRangeSeparator")) Then
        m_configuration.Fail m_configuration.GetText("legacy.MSG_TARGET_PREPARE_FAILED")
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
    ByVal lastDay As Long _
)
    Dim names() As Variant
    Dim ids() As Variant
    Dim starts() As Variant
    Dim trips() As Variant
    Dim periods() As Variant
    Dim days() As Variant
    Dim thresholds() As Variant
    Dim failures As Collection
    Dim preparedPerson As obj_PADC_PreparedPerson
    Dim rows As Collection
    Dim periodsText As String
    Dim thresholdDate As Variant
    Dim count As Long
    Dim i As Long

    ReDim names(1 To prepared.Count, 1 To 1)
    ReDim ids(1 To prepared.Count, 1 To 1)
    ReDim starts(1 To prepared.Count, 1 To 1)
    ReDim trips(1 To prepared.Count, 1 To 1)
    ReDim periods(1 To prepared.Count, 1 To 1)
    ReDim days(1 To prepared.Count, 1 To 1)
    ReDim thresholds(1 To prepared.Count, 1 To 1)
    Set failures = New Collection
    For i = 1 To prepared.Count
        m_runContext.CheckCancel i
        Set preparedPerson = prepared(i)
        If VBA.Len(preparedPerson.ValidationError) > 0 Then
            failures.Add VBA.Array(preparedPerson.FullName, preparedPerson.TaxId, preparedPerson.ValidationError)
            m_runContext.LogWarning VBA.CStr(preparedPerson.ValidationError)
        Else
            Set rows = preparedPerson.Intervals
            count = count + 1
            days(count, 1) = m_calculationService.Calculate(rows, VBA.CLng(preparedPerson.StartDay), lastDay, _
                periodsText, thresholdDate, m_runContext)
            names(count, 1) = preparedPerson.FullName
            ids(count, 1) = preparedPerson.TaxId
            starts(count, 1) = VBA.CDate(preparedPerson.StartDay)
            trips(count, 1) = preparedPerson.TripFlag
            periods(count, 1) = periodsText
            thresholds(count, 1) = thresholdDate
        End If
    Next i
    m_runContext.CheckCancel 0
    m_runContext.LogDebug m_configuration.GetText("legacy.MSG_WRITING_RESULT") & count & _
        m_configuration.GetText("legacy.MSG_ROWS")
    Application.StatusBar = m_configuration.GetText("legacy.MSG_WRITING_RESULTS")
    m_runContext.LogStage "PublishResults.Started", "Rows=" & count & " | Failures=" & failures.Count
    m_resultWriter.Publish names, ids, starts, trips, periods, days, thresholds, count, failures
    m_runContext.LogStage "PublishResults.Completed"
    m_runContext.LogDebug m_configuration.GetText("legacy.MSG_CALCULATION_COMPLETED_PEOPLE") & count & _
        m_configuration.GetText("legacy.MSG_SKIPPED_PEOPLE") & failures.Count
    ex_WindowsUi.fn_ShowMessage m_configuration.GetText("legacy.MSG_CALCULATION_COMPLETED_PEOPLE") & count & _
        m_configuration.GetText("legacy.MSG_SKIPPED_PEOPLE") & failures.Count, vbInformation
End Sub

Private Sub private_EnsureReady()
    If Not m_isInitialized Or m_isDisposed Then
        VBA.Err.Raise vbObjectError + 2100, "obj_PADC_CalculationRun", "Service is not initialized."
    End If
End Sub