Option Explicit

#Const ENABLE_LOGGING = True

Private Const DIAGNOSTIC_FOLDER_NAME As String = "PROJECTS_EXCEL"
Private Const DIAGNOSTIC_FILE_NAME As String = "diagnostic.log"
Private Const DIAGNOSTIC_MODE_IMMEDIATE As String = "Immediate"
Private Const DIAGNOSTIC_MODE_BUFFERED As String = "Buffered"
Private Const DIAGNOSTIC_MODE_CONFIG_KEY As String = "ThisWorkbook::logMode"
Private m_diagnosticSessionStarted As Boolean
Private m_diagnosticConfigurationLoaded As Boolean
Private m_diagnosticMode As String
Private m_diagnosticBuffer As Collection

' --------------------------------------
' namespace Lifecycle {
' --------------------------------------
Public Sub fn_Module_Dispose()
    If fn_Diagnostic_Flush() Then
        m_diagnosticSessionStarted = False
        m_diagnosticConfigurationLoaded = False
        m_diagnosticMode = VBA.vbNullString
        Set m_diagnosticBuffer = Nothing
    End If
End Sub
' --------------------------------------
' } // namespace Lifecycle
' --------------------------------------

' --------------------------------------
' namespace Diagnostic {
' --------------------------------------
Public Sub fn_Diagnostic_WriteLog(ByVal messageText As String)
#If ENABLE_LOGGING Then
    private_Diagnostic_EnsureConfiguration
    private_Diagnostic_WriteSessionHeader
    private_Diagnostic_WriteLine VBA.Format$(VBA.Now, "yyyy-mm-dd hh:nn:ss") & _
        " | " & messageText
#End If
End Sub

Public Function fn_Diagnostic_SetMode(ByVal modeName As String) As Boolean
    modeName = VBA.Trim$(modeName)
    If VBA.StrComp(modeName, DIAGNOSTIC_MODE_IMMEDIATE, VBA.vbTextCompare) <> 0 And _
       VBA.StrComp(modeName, DIAGNOSTIC_MODE_BUFFERED, VBA.vbTextCompare) <> 0 Then Exit Function

    private_Diagnostic_EnsureConfiguration
    If VBA.StrComp(m_diagnosticMode, DIAGNOSTIC_MODE_BUFFERED, VBA.vbTextCompare) = 0 And _
       VBA.StrComp(modeName, DIAGNOSTIC_MODE_IMMEDIATE, VBA.vbTextCompare) = 0 Then
        If Not fn_Diagnostic_Flush() Then Exit Function
    End If
    m_diagnosticMode = modeName
    fn_Diagnostic_SetMode = True
End Function

Public Function fn_Diagnostic_GetMode() As String
    private_Diagnostic_EnsureConfiguration
    fn_Diagnostic_GetMode = m_diagnosticMode
End Function

Public Function fn_Diagnostic_Flush() As Boolean
#If ENABLE_LOGGING Then
    Const FOR_APPENDING As Long = 8
    Const TRISTATE_TRUE As Long = -1
    Dim fileSystem As Object
    Dim logFile As Object
    Dim logFolderPath As String
    Dim lineText As String

    private_Diagnostic_EnsureConfiguration
    If VBA.StrComp(m_diagnosticMode, DIAGNOSTIC_MODE_BUFFERED, VBA.vbTextCompare) <> 0 Then
        fn_Diagnostic_Flush = True
        Exit Function
    End If
    If m_diagnosticBuffer Is Nothing Then
        fn_Diagnostic_Flush = True
        Exit Function
    End If
    If m_diagnosticBuffer.Count = 0 Then
        fn_Diagnostic_Flush = True
        Exit Function
    End If

    On Error GoTo EH
    Set fileSystem = VBA.CreateObject("Scripting.FileSystemObject")
    logFolderPath = private_Diagnostic_GetLogFolderPath(fileSystem)
    If VBA.Len(logFolderPath) = 0 Then GoTo EH
    Set logFile = fileSystem.OpenTextFile( _
        logFolderPath & "\" & DIAGNOSTIC_FILE_NAME, FOR_APPENDING, True, TRISTATE_TRUE)
    Do While m_diagnosticBuffer.Count > 0
        lineText = VBA.CStr(m_diagnosticBuffer(1))
        logFile.WriteLine lineText
        m_diagnosticBuffer.Remove 1
    Loop
    logFile.Close
    Set logFile = Nothing
    fn_Diagnostic_Flush = True
    Exit Function
EH:
    On Error Resume Next
    If Not logFile Is Nothing Then logFile.Close
#Else
    fn_Diagnostic_Flush = True
#End If
End Function

Public Sub fn_Diagnostic_WritePerf(ByVal stageName As String, ByVal startedAt As Double)
    Dim elapsedMilliseconds As Double

    elapsedMilliseconds = VBA.Timer - startedAt
    If elapsedMilliseconds < 0 Then elapsedMilliseconds = elapsedMilliseconds + 86400#
    fn_Diagnostic_WriteLog "PERF | Stage=" & stageName & _
        " | ElapsedMs=" & VBA.Format$(elapsedMilliseconds * 1000#, "0.0")
End Sub

Public Function fn_TryGetWorkbookProfileId(ByRef outProfileId As String) As Boolean
    Const PROFILE_ID_KEY As String = "ThisWorkbook::id"
    If fn_TryGetWorkbookConfigValue(PROFILE_ID_KEY, outProfileId) Then
        fn_TryGetWorkbookProfileId = True
        Exit Function
    End If
    VBA.MsgBox "Configuration key '" & PROFILE_ID_KEY & _
        "' must contain a profile ID in the next column.", _
        VBA.vbExclamation, "Workbook profile"
End Function

Public Function fn_TryGetWorkbookConfigValue( _
    ByVal keyName As String, _
    ByRef outValue As String _
) As Boolean
    Const CONFIG_TABLE_NAME As String = "tbConfig"
    Const CONFIG_KEY_COLUMN_NAME As String = "Key"
    Dim targetWorksheet As Worksheet
    Dim configTable As ListObject
    Dim configRow As ListRow
    Dim keyColumnIndex As Long

    outValue = VBA.vbNullString
    keyName = VBA.Trim$(keyName)
    If VBA.Len(keyName) = 0 Then Exit Function
    On Error GoTo CleanExit
    For Each targetWorksheet In ThisWorkbook.Worksheets
        For Each configTable In targetWorksheet.ListObjects
            If VBA.StrComp(configTable.Name, CONFIG_TABLE_NAME, VBA.vbTextCompare) = 0 Then
                keyColumnIndex = configTable.ListColumns(CONFIG_KEY_COLUMN_NAME).Index
                If keyColumnIndex = configTable.ListColumns.Count Then GoTo ContinueTable
                For Each configRow In configTable.ListRows
                    If VBA.StrComp( _
                            VBA.CStr(configRow.Range.Cells(1, keyColumnIndex).Value2), _
                            keyName, VBA.vbTextCompare) = 0 Then
                        outValue = VBA.Trim$(VBA.CStr( _
                            configRow.Range.Cells(1, keyColumnIndex + 1).Value2))
                        If VBA.Len(outValue) > 0 Then
                            fn_TryGetWorkbookConfigValue = True
                            Exit Function
                        End If
                    End If
                Next configRow
            End If
ContinueTable:
        Next configTable
    Next targetWorksheet
CleanExit:
End Function
Public Function fn_TryGetWorkbookConfigFolder( _
    ByVal keyName As String, _
    ByRef outFolderPath As String _
) As Boolean
    Dim configuredPath As String
    Dim fileSystem As Object

    outFolderPath = VBA.vbNullString
    If Not fn_TryGetWorkbookConfigValue(keyName, configuredPath) Then Exit Function
    configuredPath = VBA.Replace$(VBA.Trim$(configuredPath), "/", "\")
    If VBA.Len(configuredPath) = 0 Or VBA.InStr(configuredPath, ":") > 0 Or _
       VBA.Left$(configuredPath, 1) = "\" Then Exit Function
    Set fileSystem = VBA.CreateObject("Scripting.FileSystemObject")
    outFolderPath = fileSystem.GetAbsolutePathName( _
        ThisWorkbook.Path & Application.PathSeparator & configuredPath)
    If Not fileSystem.FolderExists(outFolderPath) Then
        outFolderPath = VBA.vbNullString
        Exit Function
    End If
    fn_TryGetWorkbookConfigFolder = True
End Function
' --------------------------------------
' } // namespace Diagnostic
' --------------------------------------

' //
' // Private
' //
Private Sub private_Diagnostic_EnsureConfiguration()
    Dim configuredMode As String

    If m_diagnosticConfigurationLoaded Then Exit Sub
    m_diagnosticConfigurationLoaded = True
    m_diagnosticMode = DIAGNOSTIC_MODE_IMMEDIATE
    If Not fn_TryGetWorkbookConfigValue(DIAGNOSTIC_MODE_CONFIG_KEY, configuredMode) Then Exit Sub
    configuredMode = VBA.Trim$(configuredMode)
    If VBA.StrComp(configuredMode, DIAGNOSTIC_MODE_BUFFERED, VBA.vbTextCompare) = 0 Then
        m_diagnosticMode = DIAGNOSTIC_MODE_BUFFERED
    ElseIf VBA.StrComp(configuredMode, DIAGNOSTIC_MODE_IMMEDIATE, VBA.vbTextCompare) = 0 Then
        m_diagnosticMode = DIAGNOSTIC_MODE_IMMEDIATE
    End If
End Sub

Private Sub private_Diagnostic_WriteLine(ByVal lineText As String)
    If VBA.StrComp(m_diagnosticMode, DIAGNOSTIC_MODE_BUFFERED, VBA.vbTextCompare) = 0 Then
        If m_diagnosticBuffer Is Nothing Then Set m_diagnosticBuffer = New Collection
        m_diagnosticBuffer.Add lineText
    Else
        private_Diagnostic_WriteRawLine lineText
    End If
End Sub

Private Function private_Diagnostic_WriteRawLine(ByVal lineText As String) As Boolean
    Const FOR_APPENDING As Long = 8
    Const TRISTATE_TRUE As Long = -1
    Dim fileSystem As Object
    Dim logFile As Object
    Dim logFolderPath As String

    On Error GoTo EH
    Set fileSystem = VBA.CreateObject("Scripting.FileSystemObject")
    logFolderPath = private_Diagnostic_GetLogFolderPath(fileSystem)
    If VBA.Len(logFolderPath) = 0 Then Exit Function
    Set logFile = fileSystem.OpenTextFile( _
        logFolderPath & "\" & DIAGNOSTIC_FILE_NAME, FOR_APPENDING, True, TRISTATE_TRUE)
    logFile.WriteLine lineText
    logFile.Close
    private_Diagnostic_WriteRawLine = True
    Exit Function
EH:
    On Error Resume Next
    If Not logFile Is Nothing Then logFile.Close
End Function

Private Function private_Diagnostic_GetLogFolderPath(ByVal fileSystem As Object) As String
    Const LOG_PATH_KEY As String = "ThisWorkbook::logPath"
    Dim configuredPath As String
    Dim fallbackPath As String
    Dim shellObject As Object

    fallbackPath = VBA.Environ$("TEMP") & "\" & DIAGNOSTIC_FOLDER_NAME
    If fn_TryGetWorkbookConfigValue(LOG_PATH_KEY, configuredPath) Then
        Set shellObject = VBA.CreateObject("WScript.Shell")
        configuredPath = shellObject.ExpandEnvironmentStrings(configuredPath)
        If private_Diagnostic_TryEnsureFolder(fileSystem, configuredPath) Then
            private_Diagnostic_GetLogFolderPath = configuredPath
            Exit Function
        End If
    End If
    If private_Diagnostic_TryEnsureFolder(fileSystem, fallbackPath) Then _
        private_Diagnostic_GetLogFolderPath = fallbackPath
End Function

Private Function private_Diagnostic_TryEnsureFolder( _
    ByVal fileSystem As Object, _
    ByVal folderPath As String _
) As Boolean
    Dim parentFolderPath As String

    folderPath = VBA.Trim$(folderPath)
    If VBA.Len(folderPath) = 0 Then Exit Function
    If fileSystem.FolderExists(folderPath) Then
        private_Diagnostic_TryEnsureFolder = True
        Exit Function
    End If
    parentFolderPath = fileSystem.GetParentFolderName(folderPath)
    If VBA.Len(parentFolderPath) = 0 Then Exit Function
    If Not private_Diagnostic_TryEnsureFolder(fileSystem, parentFolderPath) Then Exit Function
    fileSystem.CreateFolder folderPath
    private_Diagnostic_TryEnsureFolder = fileSystem.FolderExists(folderPath)
End Function

' Добавляет границу перед первой диагностической записью сессии.
Private Sub private_Diagnostic_WriteSessionHeader()
    If m_diagnosticSessionStarted Then Exit Sub

    private_Diagnostic_WriteLine String$(96, "=")
    private_Diagnostic_WriteLine "New session started | " & _
        VBA.Format$(VBA.Now, "yyyy-mm-dd hh:nn:ss")
    private_Diagnostic_WriteLine String$(96, "=")
    m_diagnosticSessionStarted = True
End Sub
' --------------------------------------
' } // namespace Private
' --------------------------------------