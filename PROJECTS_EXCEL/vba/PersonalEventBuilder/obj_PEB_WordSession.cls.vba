VERSION 1.0 CLASS
BEGIN
  MultiUse = -1
END
Attribute VB_Name = "obj_PEB_WordSession"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private m_isInitialized As Boolean
Private m_isDisposed As Boolean

Private m_application As Object

' //
' // Lifecycle
' //
Private Sub Class_Initialize()
End Sub

Private Sub Class_Terminate()
    On Error Resume Next
    Me.Dispose
End Sub

' //
' // API
' //
Public Function Initialize() As Boolean
    If m_isInitialized Or m_isDisposed Then Exit Function
    m_isInitialized = True
    Initialize = True
End Function

Public Sub Dispose()
    Dim errorNumber As Long
    Dim errorSource As String
    Dim errorDescription As String

    If m_isDisposed Then Exit Sub
    m_isDisposed = True
    m_isInitialized = False
    If Not m_application Is Nothing Then
        ' A disconnected Word instance is already closed; always release its proxy.
        On Error Resume Next
        VBA.Err.Clear
        m_application.Quit 0
        errorNumber = VBA.Err.Number
        errorSource = VBA.Err.Source
        errorDescription = VBA.Err.Description
        VBA.Err.Clear
        On Error GoTo 0
        Set m_application = Nothing
        If errorNumber <> 0 Then
            If errorNumber = 462 Or errorNumber = -2147417848 Then
                ex_Core.fn_Diagnostic_WriteLog "WORD_SESSION_ALREADY_CLOSED | Stage=CloseWord.Quit" & _
                    " | Number=" & VBA.CStr(errorNumber) & " | Source=" & errorSource & _
                    " | Description=" & errorDescription
            Else
                VBA.Err.Raise errorNumber, errorSource, errorDescription
            End If
        End If
    End If
End Sub

Public Function Acquire() As Object
    Dim documentCount As Long
    Dim disconnected As Boolean

    If Not m_isInitialized Or m_isDisposed Then
        Err.Raise 5, , "Word session is not initialized."
    End If
    If Not m_application Is Nothing Then
        On Error Resume Next
        documentCount = m_application.Documents.Count
        disconnected = (VBA.Err.Number <> 0)
        VBA.Err.Clear
        On Error GoTo 0
        If disconnected Then Set m_application = Nothing
    End If
    If m_application Is Nothing Then
        Set m_application = VBA.CreateObject("Word.Application")
    End If
    m_application.Visible = False
    m_application.ScreenUpdating = True
    m_application.DisplayAlerts = 0
    m_application.AutomationSecurity = 3
    Set Acquire = m_application
End Function