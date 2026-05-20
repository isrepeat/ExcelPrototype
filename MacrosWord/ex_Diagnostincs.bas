Option Explicit
#Const LOGGING_DEBUG_ENABLED = True
#Const LOGGING_VERBOSE_ENABLED = False

Private Const DIAGNOSTICS_LOG_FILE_REL_PATH As String = "Logs\\word_diagnostics.log"

Public Sub fn_Diagnostic_LogInfo(ByVal messageText As String)
#If LOGGING_DEBUG_ENABLED Then
    private_Diagnostic_LogEvent "info", messageText
#End If
End Sub

Public Sub fn_Diagnostic_LogWarning(ByVal messageText As String)
    messageText = Trim$(CStr(messageText))
    If Len(messageText) = 0 Then Exit Sub
#If LOGGING_DEBUG_ENABLED Then
    private_Diagnostic_LogEvent "warning", messageText
#End If
End Sub

Public Sub fn_Diagnostic_LogError(ByVal messageText As String)
    messageText = Trim$(CStr(messageText))
    If Len(messageText) = 0 Then Exit Sub
#If LOGGING_DEBUG_ENABLED Then
    private_Diagnostic_LogEvent "error", messageText
#End If
End Sub

Public Sub fn_Diagnostic_LogVerbose(ByVal messageText As String)
#If LOGGING_VERBOSE_ENABLED Then
    private_Diagnostic_LogEvent "verbose", messageText
#End If
End Sub

Public Sub fn_Diagnostic_LogException(ByVal sourceName As String, ByVal errNumber As Long, ByVal errDescription As String)
    sourceName = Trim$(CStr(sourceName))
    If Len(sourceName) = 0 Then sourceName = "unknown-source"
    fn_Diagnostic_LogError sourceName & ": #" & CStr(errNumber) & " " & CStr(errDescription)
End Sub

Public Sub fn_Diagnostic_LogStatusBarMessage( _
    ByVal actionName As String, _
    ByVal messageText As String, _
    Optional ByVal timeoutSeconds As Long = 0 _
)
    Dim logLine As String

    actionName = Trim$(CStr(actionName))
    If Len(actionName) = 0 Then actionName = "event"

    logLine = "status-bar-" & actionName
    If timeoutSeconds > 0 Then
        logLine = logLine & ": timeout=" & CStr(timeoutSeconds)
    End If

    messageText = Trim$(CStr(messageText))
    If Len(messageText) > 0 Then
        logLine = logLine & " message='" & private_Diagnostic_EscapeValue(messageText) & "'"
    End If

#If LOGGING_DEBUG_ENABLED Then
    private_Diagnostic_LogEvent "info", logLine
#End If
End Sub

Public Sub fn_Diagnostic_ClearLog()
    private_Diagnostic_ClearLogFile
End Sub

Private Sub private_Diagnostic_LogEvent(ByVal levelName As String, ByVal messageText As String)
    Dim logPath As String
    Dim folderPath As String
    Dim fso As Object
    Dim stream As Object
    Dim lineText As String

    messageText = Trim$(CStr(messageText))
    If Len(messageText) = 0 Then Exit Sub
    If Not private_Diagnostic_TryResolveLogPath(logPath, folderPath) Then Exit Sub

    On Error Resume Next
    Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso Is Nothing Then
        If Len(folderPath) > 0 Then
            If Not fso.FolderExists(folderPath) Then fso.CreateFolder folderPath
        End If
    End If

    lineText = Format$(Now, "yyyy-mm-dd hh:nn:ss") & " | " & CStr(levelName) & ": " & messageText
    Set stream = fso.OpenTextFile(logPath, 8, True) ' ForAppending
    If Not stream Is Nothing Then
        stream.WriteLine lineText
        stream.Close
    End If
    Err.Clear
    On Error GoTo 0
End Sub

Private Sub private_Diagnostic_ClearLogFile()
    Dim logPath As String
    Dim folderPath As String
    Dim fso As Object
    Dim stream As Object

    If Not private_Diagnostic_TryResolveLogPath(logPath, folderPath) Then Exit Sub

    On Error Resume Next
    Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso Is Nothing Then
        If Len(folderPath) > 0 Then
            If Not fso.FolderExists(folderPath) Then fso.CreateFolder folderPath
        End If
    End If

    Set stream = fso.OpenTextFile(logPath, 2, True) ' ForWriting
    If Not stream Is Nothing Then stream.Close
    Err.Clear
    On Error GoTo 0
End Sub

Private Function private_Diagnostic_TryResolveLogPath(ByRef outLogPath As String, ByRef outFolderPath As String) As Boolean
    Dim basePath As String

    On Error Resume Next
    If Documents.Count > 0 Then
        basePath = Trim$(ActiveDocument.Path)
    End If
    On Error GoTo 0

    If Len(basePath) = 0 Then Exit Function

    outLogPath = basePath & "\" & DIAGNOSTICS_LOG_FILE_REL_PATH
    outFolderPath = Left$(outLogPath, InStrRev(outLogPath, "\") - 1)
    private_Diagnostic_TryResolveLogPath = True
End Function

Private Function private_Diagnostic_EscapeValue(ByVal valueText As String) As String
    private_Diagnostic_EscapeValue = Replace$(CStr(valueText), "'", "''")
End Function
