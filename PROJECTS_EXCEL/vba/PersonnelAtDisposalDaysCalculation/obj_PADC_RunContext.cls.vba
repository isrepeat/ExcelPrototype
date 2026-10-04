VERSION 1.0 CLASS
BEGIN
  MultiUse = -1
END
Attribute VB_Name = "obj_PADC_RunContext"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private m_isInitialized As Boolean
Private m_isDisposed As Boolean

Private m_configuration As obj_PADC_Configuration
Private m_cancelled As Boolean
Private m_lastUiYield As Single
Private Const UI_YIELD_INTERVAL As Long = 100
Private Const UI_YIELD_SECONDS As Single = 0.1
Private Const CANCEL_ERROR As Long = vbObjectError + 2101
Private Const ERROR_SOURCE As String = "PersonnelAtDisposalDays"

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
    ByVal configuration As obj_PADC_Configuration _
) As Boolean
    If m_isInitialized Or m_isDisposed Then
        Exit Function
    End If
    If configuration Is Nothing Then
        Exit Function
    End If
    Set m_configuration = configuration
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
End Sub

Public Sub RequestCancel()
    private_EnsureReady
    m_cancelled = True
    Application.StatusBar = m_configuration.GetText("legacy.MSG_CANCELLING_OPERATION")
End Sub

Public Sub CheckCancel(ByVal index As Long)
    Dim currentTime As Single

    private_EnsureReady
    If m_cancelled Then
        VBA.Err.Raise CANCEL_ERROR, ERROR_SOURCE, m_configuration.GetText("legacy.MSG_CANCELLED")
    End If
    If index Mod UI_YIELD_INTERVAL <> 0 Then
        Exit Sub
    End If
    currentTime = VBA.Timer
    If index <> 0 And currentTime >= m_lastUiYield Then
        If currentTime - m_lastUiYield < UI_YIELD_SECONDS Then
            Exit Sub
        End If
    End If
    m_lastUiYield = currentTime
    Application.StatusBar = m_configuration.GetText("legacy.MSG_CALCULATING_DAYS") & index
    VBA.DoEvents
    If m_cancelled Then
        VBA.Err.Raise CANCEL_ERROR, ERROR_SOURCE, m_configuration.GetText("legacy.MSG_CANCELLED")
    End If
End Sub

Public Sub ClearLog()
    private_EnsureReady
    ex_Core.fn_Diagnostic_WriteLog "PADC_CALCULATION_STARTED"
End Sub

Public Sub LogDebug(ByVal message As String)
    private_EnsureReady
    ex_Core.fn_Diagnostic_WriteLog "PADC | " & message
End Sub

Public Sub LogWarning(ByVal message As String)
    private_EnsureReady
    ex_Core.fn_Diagnostic_WriteLog "PADC_WARNING | " & message
End Sub

Public Sub LogError(ByVal message As String)
    private_EnsureReady
    ex_Core.fn_Diagnostic_WriteLog "PADC_ERROR | " & message
End Sub

Public Sub LogFailure( _
    ByVal errorNumber As Long, _
    ByVal errorText As String _
)
    private_EnsureReady
    ex_Core.fn_Diagnostic_WriteLog "PADC_FAILED | Number=" & errorNumber & " | " & errorText
End Sub

' //
' // Private
' //
Private Sub private_EnsureReady()
    If Not m_isInitialized Or m_isDisposed Then
        VBA.Err.Raise vbObjectError + 2100, "obj_PADC_RunContext", "Service is not initialized."
    End If
End Sub