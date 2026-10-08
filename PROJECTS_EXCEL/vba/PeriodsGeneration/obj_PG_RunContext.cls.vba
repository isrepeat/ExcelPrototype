VERSION 1.0 CLASS
BEGIN
  MultiUse = -1
END
Attribute VB_Name = "obj_PG_RunContext"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private m_isInitialized As Boolean
Private m_isDisposed As Boolean

Private m_cancelRequested As Boolean
Private m_runtimeContext As Object
Private m_interval As Long

' //
' // Жизненный цикл
' //
Private Sub Class_Initialize()
End Sub

Private Sub Class_Terminate()
    Me.Dispose
End Sub

' //
' // API
' //

Public Function Initialize(ByVal configuration As obj_PG_Configuration) As Boolean
    If m_isInitialized Or m_isDisposed Then
        Exit Function
    End If
    If configuration Is Nothing Then
        Exit Function
    End If
    m_interval = VBA.CLng(configuration.GetText("legacy.UI_YIELD_INTERVAL"))
    Set m_runtimeContext = ex_RuntimeLifecycle.fn_Context()
    m_isInitialized = True
    Initialize = True
End Function

Public Sub Dispose()
    If m_isDisposed Then
        Exit Sub
    End If
    m_isInitialized = False
    m_isDisposed = True
    Set m_runtimeContext = Nothing
End Sub

Public Sub RequestCancel()
    private_EnsureReady
    m_cancelRequested = True
End Sub

Public Sub CheckCancel(ByVal index As Long)
    private_EnsureReady
    If index Mod m_interval = 0 Then
        DoEvents
    End If
    If m_cancelRequested Or ex_RuntimeLifecycle.fn_StopRequested(m_runtimeContext) Then
        VBA.Err.Raise vbObjectError + 2301, "PeriodsGeneration", "Cancelled"
    End If
End Sub

Public Sub LogWarning(ByVal message As String)
    private_EnsureReady
    ex_Core.fn_Diagnostic_WriteLog "PG_WARNING | " & message
End Sub

Public Sub LogDebug(ByVal message As String)
    private_EnsureReady
    ex_Core.fn_Diagnostic_WriteLog "PG | " & message
End Sub

' //
' // Вспомогательные методы
' //
Private Sub private_EnsureReady()
    If Not m_isInitialized Or m_isDisposed Then
        VBA.Err.Raise vbObjectError + 2300, "obj_PG_RunContext", "Service is not initialized."
    End If
End Sub