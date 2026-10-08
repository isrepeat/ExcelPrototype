VERSION 1.0 CLASS
BEGIN
  MultiUse = -1
END
Attribute VB_Name = "obj_PG_CalculationRun"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private m_isInitialized As Boolean
Private m_isDisposed As Boolean

Private m_pageBase As obj_PageBase
Private m_configuration As obj_PG_Configuration
Private m_runContext As obj_PG_RunContext
Private m_parameters As obj_PG_Parameters
Private m_dataSource As obj_PG_DataSource
Private m_calculationService As obj_PG_CalculationService
Private m_resultWriter As obj_PG_ResultWriter
Private m_running As Boolean
Private m_disposeRequested As Boolean

' //
' // Жизненный цикл
' //
Private Sub Class_Initialize()
End Sub

Private Sub Class_Terminate()
    Me.Dispose
End Sub

' //
' // Свойства
' //
Public Property Get IsRunning() As Boolean
    private_EnsureReady
    IsRunning = m_running
End Property

' //
' // API
' //
Public Function Initialize( _
    ByVal pageBase As obj_PageBase, _
    ByVal configuration As obj_PG_Configuration _
) As Boolean
    If m_isInitialized Or m_isDisposed Then
        Exit Function
    End If
    If pageBase Is Nothing Or configuration Is Nothing Then
        Exit Function
    End If
    Set m_pageBase = pageBase
    Set m_configuration = configuration
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
    Set m_configuration = Nothing
End Sub

Public Function Calculate() As Boolean
    Dim sourceTable As ListObject
    Dim values As Variant
    Dim count As Long
    Dim oldStatus As Variant
    Dim errorNumber As Long
    Dim errorText As String
    Dim runtimeContext As Object

    private_EnsureReady
    If m_running Then
        ex_WindowsUi.fn_ShowMessage m_configuration.GetText("message.AlreadyRunning"), vbExclamation
        Exit Function
    End If
    If Not ex_RuntimeLifecycle.fn_TryEnter(runtimeContext) Then
        Exit Function
    End If
    oldStatus = Application.StatusBar
    On Error GoTo Failed
    m_running = True
    Set m_runContext = New obj_PG_RunContext
    If Not m_runContext.Initialize(m_configuration) Then
        GoTo ServiceFailed
    End If
    Set m_parameters = New obj_PG_Parameters
    If Not m_parameters.Initialize(m_pageBase, m_configuration) Then
        GoTo Cleanup
    End If
    Set m_dataSource = New obj_PG_DataSource
    If Not m_dataSource.Initialize(m_configuration, m_runContext) Then
        GoTo ServiceFailed
    End If
    Set m_calculationService = New obj_PG_CalculationService
    If Not m_calculationService.Initialize(m_configuration, m_runContext) Then
        GoTo ServiceFailed
    End If
    Set m_resultWriter = New obj_PG_ResultWriter
    If Not m_resultWriter.Initialize(m_pageBase, m_configuration) Then
        GoTo ServiceFailed
    End If
    Set sourceTable = m_dataSource.FindSourceTable()
    count = m_calculationService.Calculate(sourceTable, m_parameters.StartDate, m_parameters.EndDate, values)
    m_runContext.CheckCancel 0
    Application.StatusBar = m_configuration.GetText("status.Writing")
    m_resultWriter.Publish values, count
    m_runContext.LogDebug "Completed | Count=" & count
    ex_WindowsUi.fn_ShowMessage VBA.Replace(m_configuration.GetText("message.Done"), "{count}", VBA.CStr(count)), vbInformation
    Calculate = True
Cleanup:
    Set sourceTable = Nothing
    private_DisposeServices
    Application.StatusBar = oldStatus
    m_running = False
    ex_RuntimeLifecycle.fn_Leave runtimeContext
    If m_disposeRequested Then
        Me.Dispose
    End If
    Exit Function
ServiceFailed:
    m_configuration.Fail m_configuration.GetText("message.PublishFailed")
Failed:
    errorNumber = VBA.Err.Number
    errorText = VBA.Err.Description
    ex_Core.fn_Diagnostic_WriteLog "PG_FAILED | Number=" & errorNumber & " | " & errorText
    If errorNumber = vbObjectError + 2301 Then
        ex_WindowsUi.fn_ShowMessage m_configuration.GetText("message.Cancelled"), vbInformation
    Else
        ex_WindowsUi.fn_ShowMessage VBA.Replace(m_configuration.GetText("message.Failed"), "{error}", errorText), vbExclamation
    End If
    Resume Cleanup
End Function

Public Sub RequestCancel()
    private_EnsureReady
    If m_running And Not m_runContext Is Nothing Then
        m_runContext.RequestCancel
    End If
End Sub

' //
' // Вспомогательные методы
' //
Private Sub private_DisposeServices()
    If Not m_resultWriter Is Nothing Then
        m_resultWriter.Dispose
    End If
    Set m_resultWriter = Nothing
    If Not m_calculationService Is Nothing Then
        m_calculationService.Dispose
    End If
    Set m_calculationService = Nothing
    If Not m_dataSource Is Nothing Then
        m_dataSource.Dispose
    End If
    Set m_dataSource = Nothing
    If Not m_parameters Is Nothing Then
        m_parameters.Dispose
    End If
    Set m_parameters = Nothing
    If Not m_runContext Is Nothing Then
        m_runContext.Dispose
    End If
    Set m_runContext = Nothing
End Sub

Private Sub private_EnsureReady()
    If Not m_isInitialized Or m_isDisposed Then
        VBA.Err.Raise vbObjectError + 2300, "obj_PG_CalculationRun", "Service is not initialized."
    End If
End Sub