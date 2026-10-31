Option Explicit

Private m_application As Object

' --------------------------------------
' namespace Lifecycle {
' --------------------------------------
Public Sub fn_Module_Dispose()
    Dim wordApplication As Object
    Dim failure As String

    Set wordApplication = m_application
    Set m_application = Nothing
    If wordApplication Is Nothing Then Exit Sub
    On Error Resume Next
    wordApplication.Quit 0
    If VBA.Err.Number <> 0 Then failure = VBA.Err.Description
    Set wordApplication = Nothing
    If VBA.Len(failure) > 0 Then ex_Core.fn_Diagnostic_WriteLog "TDB_WORD_SESSION_DISPOSE_FAILED | " & failure
    On Error GoTo 0
End Sub
' --------------------------------------
' } // namespace Lifecycle
' --------------------------------------

' --------------------------------------
' namespace API {
' --------------------------------------
Public Function fn_Acquire() As Object
    Dim documentCount As Long
    Dim disconnected As Boolean

    ' Retain the private hidden process between exports, as PEB does with its Word session.
    If Not m_application Is Nothing Then
        On Error Resume Next
        documentCount = m_application.Documents.Count
        disconnected = (VBA.Err.Number <> 0)
        VBA.Err.Clear
        On Error GoTo 0
        If disconnected Then Set m_application = Nothing
    End If
    If m_application Is Nothing Then Set m_application = VBA.CreateObject("Word.Application")
    m_application.Visible = False
    m_application.ScreenUpdating = False
    m_application.DisplayAlerts = 0
    m_application.AutomationSecurity = 3
    Set fn_Acquire = m_application
End Function
' --------------------------------------
' } // namespace API
' --------------------------------------