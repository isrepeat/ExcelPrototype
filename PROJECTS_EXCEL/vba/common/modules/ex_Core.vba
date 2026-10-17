Option Explicit

#If VBA7 Then
Private Declare PtrSafe Function private_GetUserDefaultLocaleName Lib "kernel32" Alias "GetUserDefaultLocaleName" (ByVal buffer As LongPtr, ByVal count As Long) As Long
Private Declare PtrSafe Function private_GetSystemDefaultLocaleName Lib "kernel32" Alias "GetSystemDefaultLocaleName" (ByVal buffer As LongPtr, ByVal count As Long) As Long
Private Declare PtrSafe Function private_GetACP Lib "kernel32" Alias "GetACP" () As Long
Private Declare PtrSafe Function private_GetThreadLocale Lib "kernel32" Alias "GetThreadLocale" () As Long
#Else
Private Declare Function private_GetUserDefaultLocaleName Lib "kernel32" Alias "GetUserDefaultLocaleName" (ByVal buffer As Long, ByVal count As Long) As Long
Private Declare Function private_GetSystemDefaultLocaleName Lib "kernel32" Alias "GetSystemDefaultLocaleName" (ByVal buffer As Long, ByVal count As Long) As Long
Private Declare Function private_GetACP Lib "kernel32" Alias "GetACP" () As Long
Private Declare Function private_GetThreadLocale Lib "kernel32" Alias "GetThreadLocale" () As Long
#End If

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
Private m_statusBarMessage As String
Private m_statusBarOwned As Boolean
Private m_statusBarPrevious As Variant

' --------------------------------------
' namespace Lifecycle {
' --------------------------------------
Public Sub fn_Module_Dispose()
    fn_StatusBar_Clear
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
' namespace StatusBar {
' --------------------------------------
Public Sub fn_StatusBar_Show(ByVal messageText As String)
    Dim currentMessage As Variant
    Dim displayText As String

    On Error GoTo EH
    If VBA.Len(VBA.Trim$(messageText)) = 0 Then
        fn_StatusBar_Clear
        Exit Sub
    End If
    currentMessage = Application.StatusBar
    If Not m_statusBarOwned Then
        m_statusBarPrevious = currentMessage
    ElseIf VBA.CStr(currentMessage) <> m_statusBarMessage Then
        m_statusBarPrevious = currentMessage
    End If
    ' Excel rejects StatusBar strings longer than 255 characters.
    displayText = VBA.Left$(messageText, 255)
    Application.DisplayStatusBar = True
    Application.StatusBar = displayText
    m_statusBarMessage = displayText
    m_statusBarOwned = True
    Exit Sub
EH:
    private_StatusBarError "Show", VBA.Err.Number, VBA.Err.Source, VBA.Err.Description, VBA.Len(messageText)
End Sub

Public Sub fn_StatusBar_Clear()
    Dim currentMessage As Variant

    On Error GoTo EH
    If Not m_statusBarOwned Then Exit Sub
    currentMessage = Application.StatusBar
    ' Excel shares the status bar: do not clear a newer message from another owner.
    If VBA.VarType(currentMessage) = VBA.vbString Then
        If VBA.StrComp(VBA.CStr(currentMessage), m_statusBarMessage, VBA.vbBinaryCompare) = 0 Then
            Application.StatusBar = m_statusBarPrevious
        End If
    End If
    m_statusBarOwned = False
    m_statusBarMessage = VBA.vbNullString
    m_statusBarPrevious = Empty
    Exit Sub
EH:
    private_StatusBarError "Clear", VBA.Err.Number, VBA.Err.Source, VBA.Err.Description, VBA.Len(m_statusBarMessage)
    m_statusBarOwned = False
    m_statusBarMessage = VBA.vbNullString
    m_statusBarPrevious = Empty
End Sub
' --------------------------------------
' } // namespace StatusBar
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
    If Not fn_TrySetWorkbookConfigValue(DIAGNOSTIC_MODE_CONFIG_KEY, modeName) Then Exit Function
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

Public Sub fn_Diagnostic_WriteError( _
    ByVal stageName As String, _
    ByVal errorNumber As Long, _
    ByVal errorSource As String, _
    ByVal errorDescription As String _
)
    ' Реквизиты передаются до очистки Err и вызова обработчиков восстановления.
    On Error GoTo EH
    fn_Diagnostic_WriteLog "ERROR | Stage=" & stageName & " | Number=" & VBA.CStr(errorNumber) & _
        " | Source=" & errorSource & " | Description=" & errorDescription
    Exit Sub
EH:
    Debug.Print "Diagnostic error recording failed: " & VBA.Err.Description
End Sub

' --------------------------------------
' } // namespace Diagnostic
' --------------------------------------

' --------------------------------------
' namespace Configuration {
' --------------------------------------
Public Function fn_TryGetWorkbookProfileId(ByRef outProfileId As String) As Boolean
    Const PROFILE_ID_KEY As String = "ThisWorkbook::id"
    If fn_TryGetWorkbookConfigValue(PROFILE_ID_KEY, outProfileId) Then
        fn_TryGetWorkbookProfileId = True
        Exit Function
    End If
    ex_WindowsUi.fn_ShowMessage "Configuration key '" & PROFILE_ID_KEY & _
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
Public Function fn_TrySetWorkbookConfigValue( _
    ByVal keyName As String, _
    ByVal newValue As String _
) As Boolean
    Const CONFIG_TABLE_NAME As String = "tbConfig"
    Const CONFIG_KEY_COLUMN_NAME As String = "Key"
    Dim targetWorksheet As Worksheet
    Dim candidateTable As ListObject
    Dim configTable As ListObject
    Dim configRow As ListRow
    Dim matchedRow As ListRow
    Dim addedRow As ListRow
    Dim keyColumnIndex As Long
    Dim previousEnableEvents As Boolean
    Dim eventsCaptured As Boolean
    Dim errorDescription As String

    On Error GoTo EH
    keyName = VBA.Trim$(keyName)
    If VBA.Len(keyName) = 0 Then _
        VBA.Err.Raise VBA.vbObjectError + 2210, _
            "fn_TrySetWorkbookConfigValue", "Configuration key must not be empty."
    For Each targetWorksheet In ThisWorkbook.Worksheets
        For Each candidateTable In targetWorksheet.ListObjects
            If VBA.StrComp(candidateTable.Name, CONFIG_TABLE_NAME, VBA.vbTextCompare) = 0 Then
                If Not configTable Is Nothing Then _
                    VBA.Err.Raise VBA.vbObjectError + 2211, _
            "fn_TrySetWorkbookConfigValue", "Multiple tbConfig tables were found."
                Set configTable = candidateTable
            End If
        Next candidateTable
    Next targetWorksheet
    If configTable Is Nothing Then _
        VBA.Err.Raise VBA.vbObjectError + 2212, _
            "fn_TrySetWorkbookConfigValue", "Configuration table tbConfig was not found."
    keyColumnIndex = configTable.ListColumns(CONFIG_KEY_COLUMN_NAME).Index
    If keyColumnIndex = configTable.ListColumns.Count Then _
        VBA.Err.Raise VBA.vbObjectError + 2213, _
            "fn_TrySetWorkbookConfigValue", "tbConfig must have a value column after Key."
    For Each configRow In configTable.ListRows
        If VBA.StrComp(VBA.CStr(configRow.Range.Cells(1, keyColumnIndex).Value2), _
                keyName, VBA.vbTextCompare) = 0 Then
            If Not matchedRow Is Nothing Then _
                VBA.Err.Raise VBA.vbObjectError + 2214, _
            "fn_TrySetWorkbookConfigValue", "Duplicate configuration key: " & keyName
            Set matchedRow = configRow
        End If
    Next configRow

    ' Disable events only while writing and preserve the original Excel state.
    previousEnableEvents = Application.EnableEvents
    eventsCaptured = True
    Application.EnableEvents = False
    If matchedRow Is Nothing Then
        Set addedRow = configTable.ListRows.Add
        addedRow.Range.Cells(1, keyColumnIndex).Value2 = keyName
        Set matchedRow = addedRow
    End If
    matchedRow.Range.Cells(1, keyColumnIndex + 1).Value2 = newValue
    Application.EnableEvents = previousEnableEvents
    eventsCaptured = False
    fn_TrySetWorkbookConfigValue = True
    Exit Function
EH:
    errorDescription = VBA.Err.Description
    On Error Resume Next
    If Not addedRow Is Nothing Then addedRow.Delete
    If eventsCaptured Then Application.EnableEvents = previousEnableEvents
    On Error GoTo 0
    ex_WindowsUi.fn_ShowMessage "Could not write configuration key '" & keyName & "': " & errorDescription, _
        VBA.vbExclamation, "Workbook configuration"
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
' } // namespace Configuration
' --------------------------------------

Private Sub private_StatusBarError( _
    ByVal operation As String, _
    ByVal errorNumber As Long, _
    ByVal errorSource As String, _
    ByVal errorDescription As String, _
    ByVal textLength As Long _
)
    ' Status display and its diagnostics must never interrupt the caller's operation.
    On Error GoTo EH
    fn_Diagnostic_WriteError _
        "StatusBar." & operation & " | TextLength=" & VBA.CStr(textLength), _
        errorNumber, _
        errorSource, _
        errorDescription
    Exit Sub
EH:
    Debug.Print "StatusBar diagnostic failed: " & errorDescription
End Sub

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

' Adds a separator before the first diagnostic entry of the session.
Private Sub private_Diagnostic_WriteSessionHeader()
    If m_diagnosticSessionStarted Then Exit Sub

    private_Diagnostic_WriteLine String$(96, "=")
    private_Diagnostic_WriteLine "New session started | " & _
        VBA.Format$(VBA.Now, "yyyy-mm-dd hh:nn:ss")
    private_Diagnostic_WriteLine String$(96, "=")
    m_diagnosticSessionStarted = True
    private_Diagnostic_WriteEnvironment
End Sub

Private Sub private_Diagnostic_WriteEnvironment()
    Dim profileId As String

    ' Снимок пишется один раз; ошибка диагностики не прерывает запуск приложения.
    On Error GoTo EH
    private_Diagnostic_WriteLine "ENV_WINDOWS | UserLocale=" & private_Diagnostic_LocaleName(False) & _
        " | SystemLocale=" & private_Diagnostic_LocaleName(True) & _
        " | ThreadLCID=" & VBA.CStr(private_GetThreadLocale()) & " | ACP=" & VBA.CStr(private_GetACP())
    private_Diagnostic_WriteLine "ENV_HOST | OS=" & Application.OperatingSystem & _
        " | ExcelVersion=" & Application.Version & " | ExcelBuild=" & VBA.CStr(Application.Build)
#If Win64 Then
    private_Diagnostic_WriteLine "ENV_OFFICE | ProcessBits=64"
#Else
    private_Diagnostic_WriteLine "ENV_OFFICE | ProcessBits=32"
#End If
    private_Diagnostic_WriteLine "ENV_EXCEL | UILanguageID=" & VBA.CStr(Application.LanguageSettings.LanguageID(2)) & _
        " | CountrySetting=" & VBA.CStr(Application.International(xlCountrySetting)) & _
        " | UseSystemSeparators=" & VBA.CStr(Application.UseSystemSeparators) & _
        " | DecimalSeparator=[" & Application.DecimalSeparator & "]" & _
        " | ThousandsSeparator=[" & Application.ThousandsSeparator & "]"
    private_Diagnostic_WriteLine "ENV_FORMAT | ListSeparator=[" & VBA.CStr(Application.International(xlListSeparator)) & _
        "] | DateSeparator=[" & VBA.CStr(Application.International(xlDateSeparator)) & _
        "] | DateOrder=" & VBA.CStr(Application.International(xlDateOrder)) & _
        " | VBA_NumberSample=[" & VBA.CStr(1234.5) & "] | VBA_DateSample=[" & VBA.CStr(VBA.DateSerial(2026, 4, 3)) & "]"
    If Not fn_TryGetWorkbookConfigValue("ThisWorkbook::id", profileId) Then profileId = "<missing>"
    private_Diagnostic_WriteLine "ENV_WORKBOOK | FullName=" & ThisWorkbook.FullName & _
        " | Profile=" & profileId & " | Date1904=" & VBA.CStr(ThisWorkbook.Date1904) & _
        " | ReadOnly=" & VBA.CStr(ThisWorkbook.ReadOnly) & " | LogMode=" & m_diagnosticMode
    Exit Sub
EH:
    private_Diagnostic_WriteLine "ENV_CAPTURE_FAILED | Number=" & VBA.CStr(VBA.Err.Number) & _
        " | Source=" & VBA.Err.Source & " | Description=" & VBA.Err.Description
End Sub

Private Function private_Diagnostic_LocaleName(ByVal systemLocale As Boolean) As String
    Dim buffer As String
    Dim count As Long
    Dim lastError As Long

    buffer = VBA.String$(85, VBA.vbNullChar)
    If systemLocale Then
        count = private_GetSystemDefaultLocaleName(VBA.StrPtr(buffer), 85)
    Else
        count = private_GetUserDefaultLocaleName(VBA.StrPtr(buffer), 85)
    End If
    If count = 0 Then
        lastError = VBA.Err.LastDllError
        private_Diagnostic_LocaleName = "<unavailable: Win32=" & VBA.CStr(lastError) & ">"
    Else
        private_Diagnostic_LocaleName = VBA.Left$(buffer, count - 1)
    End If
End Function