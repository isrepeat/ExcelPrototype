Option Explicit

#Const ENABLE_LOGGING = True

Private Const DIAGNOSTIC_FOLDER_NAME As String = "4. PersonalEventBuilder"
Private Const DIAGNOSTIC_FILE_NAME As String = "diagnostic.log"
Private m_diagnosticSessionStarted As Boolean

' --------------------------------------
' namespace Lifecycle {
' --------------------------------------
Public Sub fn_Module_Dispose()
    m_diagnosticSessionStarted = False
End Sub
' --------------------------------------
' } // namespace Lifecycle
' --------------------------------------

' --------------------------------------
' namespace Diagnostic {
' --------------------------------------
Public Sub fn_Diagnostic_WriteLog(ByVal messageText As String)
#If ENABLE_LOGGING Then
    private_Diagnostic_WriteSessionHeader
    private_Diagnostic_WriteRawLine VBA.Format$(VBA.Now, "yyyy-mm-dd hh:nn:ss") & _
        " | " & messageText
#End If
End Sub
' --------------------------------------
' } // namespace Diagnostic
' --------------------------------------

' Writes a prepared diagnostic line to the session log.
Private Sub private_Diagnostic_WriteRawLine(ByVal lineText As String)
    Const FOR_APPENDING As Long = 8
    Const TRISTATE_TRUE As Long = -1
    Dim fileSystem As Object
    Dim logFile As Object
    Dim tempPath As String
    Dim logFolderPath As String

    On Error Resume Next
    tempPath = VBA.Environ$("TEMP")
    If VBA.Len(tempPath) = 0 Then Exit Sub
    logFolderPath = tempPath & "\" & DIAGNOSTIC_FOLDER_NAME
    Set fileSystem = VBA.CreateObject("Scripting.FileSystemObject")
    If Not fileSystem.FolderExists(logFolderPath) Then fileSystem.CreateFolder logFolderPath
    Set logFile = fileSystem.OpenTextFile( _
        logFolderPath & "\" & DIAGNOSTIC_FILE_NAME, FOR_APPENDING, True, TRISTATE_TRUE)
    logFile.WriteLine lineText
    logFile.Close
End Sub

' Adds a visible boundary before the first diagnostic record of this VBA session.
Private Sub private_Diagnostic_WriteSessionHeader()
    If m_diagnosticSessionStarted Then Exit Sub

    private_Diagnostic_WriteRawLine String$(96, "=")
    private_Diagnostic_WriteRawLine "New session started | " & _
        VBA.Format$(VBA.Now, "yyyy-mm-dd hh:nn:ss")
    private_Diagnostic_WriteRawLine String$(96, "=")
    m_diagnosticSessionStarted = True
End Sub