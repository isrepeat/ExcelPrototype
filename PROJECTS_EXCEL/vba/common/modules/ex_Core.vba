Option Explicit

#Const ENABLE_LOGGING = True

Private Const DIAGNOSTIC_FOLDER_NAME As String = "PROJECTS_EXCEL"
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

' Writes a prepared diagnostic line to the session log.
Private Sub private_Diagnostic_WriteRawLine(ByVal lineText As String)
    Const FOR_APPENDING As Long = 8
    Const TRISTATE_TRUE As Long = -1
    Dim fileSystem As Object
    Dim logFile As Object
    Dim logFolderPath As String

    On Error Resume Next
    Set fileSystem = VBA.CreateObject("Scripting.FileSystemObject")
    logFolderPath = private_Diagnostic_GetLogFolderPath(fileSystem)
    If VBA.Len(logFolderPath) = 0 Then Exit Sub
    Set logFile = fileSystem.OpenTextFile( _
        logFolderPath & "\" & DIAGNOSTIC_FILE_NAME, FOR_APPENDING, True, TRISTATE_TRUE)
    logFile.WriteLine lineText
    logFile.Close
End Sub

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

' Adds a visible boundary before the first diagnostic record of this VBA session.
Private Sub private_Diagnostic_WriteSessionHeader()
    If m_diagnosticSessionStarted Then Exit Sub

    private_Diagnostic_WriteRawLine String$(96, "=")
    private_Diagnostic_WriteRawLine "New session started | " & _
        VBA.Format$(VBA.Now, "yyyy-mm-dd hh:nn:ss")
    private_Diagnostic_WriteRawLine String$(96, "=")
    m_diagnosticSessionStarted = True
End Sub