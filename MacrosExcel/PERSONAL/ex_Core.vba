Option Explicit

#Const ENABLE_LOGGING = True

Private Const DIAGNOSTIC_FOLDER_NAME As String = "PERSONAL.EXCEL"
Private Const DIAGNOSTIC_FILE_NAME As String = "diagnostic.log"
Private m_diagnosticSessionStarted As Boolean
Private m_hotkeyBrokerBindingsByWorkbook As Object
Private m_hotkeyBrokerLegacyCleanupCompleted As Boolean

' --------------------------------------
' namespace API {
' --------------------------------------
' Recreates PERSONAL.XLSB runtime state after source modules are reloaded.
Public Sub fn_ReloadPersonalRuntime()
    Dim errorNumber As Long
    Dim errorDescription As String

    fn_Diagnostic_WriteLog "PERSONAL_RUNTIME_RELOAD_REQUESTED"
    On Error GoTo EH
    ThisWorkbook.fn_ReloadRuntime
    fn_Diagnostic_WriteLog "PERSONAL_RUNTIME_RELOAD_REQUEST_COMPLETED"
    Exit Sub
EH:
    errorNumber = VBA.Err.Number
    errorDescription = VBA.Err.Description
    fn_Diagnostic_WriteLog "PERSONAL_RUNTIME_RELOAD_REQUEST_ERROR | Number=" & _
        VBA.CStr(errorNumber) & " | Description=" & errorDescription
    Err.Raise errorNumber, "ex_Core.fn_ReloadPersonalRuntime", errorDescription
End Sub

Public Sub fn_RestorePersonalHotkeys()
    Dim errorNumber As Long
    Dim errorDescription As String

    fn_Diagnostic_WriteLog "PERSONAL_HOTKEY_RESTORE_REQUESTED"
    On Error GoTo EH
    ThisWorkbook.BindKeys
    fn_Diagnostic_WriteLog "PERSONAL_HOTKEY_RESTORE_REQUEST_COMPLETED"
    Exit Sub
EH:
    errorNumber = VBA.Err.Number
    errorDescription = VBA.Err.Description
    fn_Diagnostic_WriteLog "PERSONAL_HOTKEY_RESTORE_REQUEST_ERROR | Number=" & _
        VBA.CStr(errorNumber) & " | Description=" & errorDescription
    Err.Raise errorNumber, "ex_Core.fn_RestorePersonalHotkeys", errorDescription
End Sub

Public Function fn_HotkeyBroker_Activate( _
    ByVal workbookFullName As String, _
    ByVal bindingsText As String _
) As Boolean
    Dim targetWorkbook As Workbook
    Dim bindingsByKey As Object

    On Error GoTo EH
    Set targetWorkbook = private_HotkeyBroker_FindWorkbook(workbookFullName)
    If targetWorkbook Is Nothing Then Exit Function
    If Not private_HotkeyBroker_TryParseBindings(bindingsText, bindingsByKey) Then Exit Function
    private_HotkeyBroker_EnsureRegistry
    private_HotkeyBroker_ClearLegacyBindings
    private_HotkeyBroker_UnbindAll
    ThisWorkbook.BindKeys
    private_HotkeyBroker_BindWorkbook targetWorkbook, bindingsByKey
    Set m_hotkeyBrokerBindingsByWorkbook(workbookFullName) = bindingsByKey
    fn_Diagnostic_WriteLog "PERSONAL_HOTKEY_BROKER_ACTIVATED | Workbook=" & _
        targetWorkbook.Name & " | KeyCount=" & VBA.CStr(bindingsByKey.Count)
    fn_HotkeyBroker_Activate = True
    Exit Function
EH:
    fn_Diagnostic_WriteLog "PERSONAL_HOTKEY_BROKER_ACTIVATE_ERROR | Number=" & _
        VBA.CStr(VBA.Err.Number) & " | Description=" & VBA.Err.Description
    private_HotkeyBroker_UnbindAll
    ThisWorkbook.BindKeys
End Function

Public Function fn_HotkeyBroker_Deactivate( _
    ByVal workbookFullName As String _
) As Boolean
    On Error GoTo EH
    private_HotkeyBroker_EnsureRegistry
    private_HotkeyBroker_UnbindWorkbook workbookFullName
    ThisWorkbook.BindKeys
    fn_Diagnostic_WriteLog "PERSONAL_HOTKEY_BROKER_DEACTIVATED | Workbook=" & _
        workbookFullName
    fn_HotkeyBroker_Deactivate = True
    Exit Function
EH:
    fn_Diagnostic_WriteLog "PERSONAL_HOTKEY_BROKER_DEACTIVATE_ERROR | Number=" & _
        VBA.CStr(VBA.Err.Number) & " | Description=" & VBA.Err.Description
End Function

Public Sub fn_ReloadActiveWorkbookVba()
    private_VbaReload_RequestSafeReload
End Sub

' Совместимое имя для прежнего сочетания клавиш.
Public Sub fn_ReloadActiveWorkbookVbaDeferred()
    private_VbaReload_RequestSafeReload
End Sub

Private Sub private_VbaReload_ReloadActiveWorkbookVba( _
    ByVal deferInitialization As Boolean _
)
    Const VBEXT_CT_STD_MODULE As Long = 1
    Const VBEXT_CT_CLASS_MODULE As Long = 2
    Const VBEXT_CT_MS_FORM As Long = 3

    Dim targetWorkbook As Workbook
    Dim vbProject As Object
    Dim vbComponent As Object
    Dim componentNames As Collection
    Dim importFiles As Collection
    Dim documentImportFiles As Collection
    Dim vbaFolderPath As String
    Dim uiFolderPath As String
    Dim targetWorkbookName As String
    Dim importedCount As Long
    Dim previousEnableEvents As Boolean
    Dim previousScreenUpdating As Boolean
    Dim applicationStateChanged As Boolean
    Dim reloadSucceeded As Boolean
    Dim componentName As Variant
    Dim importFile As Variant
    Dim startedAt As Double

    startedAt = VBA.Timer
    On Error GoTo EH
    fn_Diagnostic_WriteLog "VBA_RELOAD_ENTRY | DeferredInitialization=" & _
        VBA.CStr(deferInitialization) & " | ExcelVersion=" & Application.Version & _
        " | Workbooks=" & VBA.CStr(Application.Workbooks.Count)

    If Application.ActiveWorkbook Is Nothing Then
        fn_Diagnostic_WriteLog "VBA_RELOAD_ABORTED | Reason=NoActiveWorkbook"
        VBA.MsgBox "There is no active workbook whose VBA modules can be reloaded.", _
            VBA.vbExclamation, "Reload VBA"
        Exit Sub
    End If

    Set targetWorkbook = Application.ActiveWorkbook
    targetWorkbookName = targetWorkbook.Name
    fn_Diagnostic_WriteLog "VBA_RELOAD_STARTED | Workbook=" & targetWorkbookName
    fn_Diagnostic_WriteLog "VBA_RELOAD_CONTEXT | DeferredInitialization=" & _
        VBA.CStr(deferInitialization) & " | Path=" & targetWorkbook.FullName
    If targetWorkbook Is ThisWorkbook Then
        fn_Diagnostic_WriteLog "VBA_RELOAD_ABORTED | Reason=TargetIsPersonalWorkbook"
        VBA.MsgBox _
            "The workbook containing the shortcut handler cannot reload itself. " & _
            "Activate the target workbook and press Ctrl+Alt+R again.", _
            VBA.vbExclamation, "Reload VBA"
        Exit Sub
    End If

    If VBA.Len(targetWorkbook.Path) = 0 Then
        fn_Diagnostic_WriteLog "VBA_RELOAD_ABORTED | Reason=TargetWorkbookUnsaved"
        VBA.MsgBox _
            "Save the active workbook first. The vba folder must be located " & _
            "next to the workbook file.", _
            VBA.vbExclamation, "Reload VBA"
        Exit Sub
    End If

    fn_Diagnostic_WriteLog "VBA_RELOAD_STAGE_STARTED | Name=ResolveSourceFolders"
    If Not private_VbaReload_TryResolveConfiguredFolder(targetWorkbook, "ThisWorkbook::vbaPath", vbaFolderPath) Then
        fn_Diagnostic_WriteLog "VBA_RELOAD_ABORTED | Stage=ResolveSourceFolders | Key=ThisWorkbook::vbaPath"
        Exit Sub
    End If
    If Not private_VbaReload_TryResolveConfiguredFolder(targetWorkbook, "ThisWorkbook::uiPath", uiFolderPath) Then
        fn_Diagnostic_WriteLog "VBA_RELOAD_ABORTED | Stage=ResolveSourceFolders | Key=ThisWorkbook::uiPath"
        Exit Sub
    End If
    fn_Diagnostic_WriteLog "VBA_RELOAD_STAGE_COMPLETED | Name=ResolveSourceFolders" & _
        " | VbaPath=" & vbaFolderPath & " | UiPath=" & uiFolderPath

    Set importFiles = New Collection
    Set documentImportFiles = New Collection
    fn_Diagnostic_WriteLog "VBA_RELOAD_STAGE_STARTED | Name=CollectSourceFiles"
    If Not private_VbaReload_TryCollectConfiguredVbaFiles( _
            vbaFolderPath, targetWorkbook, importFiles, documentImportFiles) Then
        fn_Diagnostic_WriteLog "VBA_RELOAD_ABORTED | Stage=CollectSourceFiles"
        Exit Sub
    End If
    fn_Diagnostic_WriteLog "VBA_RELOAD_STAGE_COMPLETED | Name=CollectSourceFiles" & _
        " | ModuleFiles=" & VBA.CStr(importFiles.Count) & _
        " | DocumentFiles=" & VBA.CStr(documentImportFiles.Count)
    If importFiles.Count = 0 And documentImportFiles.Count = 0 Then
        fn_Diagnostic_WriteLog "VBA_RELOAD_ABORTED | Reason=NoSourceFiles"
        VBA.MsgBox _
            "No .bas, .frm, .vba, or .utf8.vba files were found in: " & vbaFolderPath, _
            VBA.vbExclamation, "Reload VBA"
        Exit Sub
    End If

    fn_Diagnostic_WriteLog "VBA_RELOAD_STAGE_STARTED | Name=ReadExistingComponents"
    Set vbProject = targetWorkbook.VBProject
    Set componentNames = New Collection
    For Each vbComponent In vbProject.VBComponents
        Select Case CLng(vbComponent.Type)
            Case VBEXT_CT_STD_MODULE, VBEXT_CT_CLASS_MODULE, VBEXT_CT_MS_FORM
                componentNames.Add VBA.CStr(vbComponent.Name)
        End Select
    Next vbComponent
    fn_Diagnostic_WriteLog "VBA_RELOAD_STAGE_COMPLETED | Name=ReadExistingComponents" & _
        " | RemovableComponents=" & VBA.CStr(componentNames.Count)

    previousScreenUpdating = Application.ScreenUpdating
    previousEnableEvents = Application.EnableEvents
    applicationStateChanged = True
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.StatusBar = "Reloading VBA modules..."
    fn_Diagnostic_WriteLog "VBA_RELOAD_STAGE_STARTED | Name=RemoveExistingComponents"

    For Each componentName In componentNames
        fn_Diagnostic_WriteLog "VBA_RELOAD_COMPONENT_REMOVE_STARTED | Name=" & _
            VBA.CStr(componentName)
        vbProject.VBComponents.Remove vbProject.VBComponents(VBA.CStr(componentName))
        fn_Diagnostic_WriteLog "VBA_RELOAD_COMPONENT_REMOVE_COMPLETED | Name=" & _
            VBA.CStr(componentName)
    Next componentName
    fn_Diagnostic_WriteLog "VBA_RELOAD_STAGE_COMPLETED | Name=RemoveExistingComponents"

    ' The profile defines the complete source set. Clear workbook and worksheet
    ' code so that event handlers from the previous profile do not stay active.
    fn_Diagnostic_WriteLog "VBA_RELOAD_STAGE_STARTED | Name=ClearDocumentModules"
    private_VbaReload_ClearDocumentModules targetWorkbook
    fn_Diagnostic_WriteLog "VBA_RELOAD_STAGE_COMPLETED | Name=ClearDocumentModules"

    For Each importFile In importFiles
        fn_Diagnostic_WriteLog "VBA_RELOAD_MODULE_IMPORT_STARTED | File=" & _
            VBA.CStr(importFile)
        private_VbaReload_ImportVbaFile vbProject, VBA.CStr(importFile)
        importedCount = importedCount + 1
        fn_Diagnostic_WriteLog "VBA_RELOAD_MODULE_IMPORT_COMPLETED | File=" & _
            VBA.CStr(importFile)
    Next importFile

    ' Document modules cannot be imported as standard modules because Excel
    ' would create a separate module and Workbook/Worksheet events would not run.
    For Each importFile In documentImportFiles
        fn_Diagnostic_WriteLog "VBA_RELOAD_DOCUMENT_IMPORT_STARTED | File=" & _
            VBA.CStr(importFile)
        private_VbaReload_ImportDocumentVbaFile targetWorkbook, VBA.CStr(importFile)
        importedCount = importedCount + 1
        fn_Diagnostic_WriteLog "VBA_RELOAD_DOCUMENT_IMPORT_COMPLETED | File=" & _
            VBA.CStr(importFile)
    Next importFile
    fn_Diagnostic_WriteLog "VBA_RELOAD_STAGE_STARTED | Name=SetRuntimePaths"
    private_VbaReload_SetRuntimePaths targetWorkbook, uiFolderPath
    fn_Diagnostic_WriteLog "VBA_RELOAD_STAGE_COMPLETED | Name=SetRuntimePaths"
    fn_Diagnostic_WriteLog "VBA_RELOAD_STAGE_STARTED | Name=InitializeReloadedWorkbook"
    private_VbaReload_InitializeReloadedWorkbook targetWorkbook, deferInitialization
    fn_Diagnostic_WriteLog "VBA_RELOAD_STAGE_COMPLETED | Name=InitializeReloadedWorkbook"

    Application.StatusBar = "Imported VBA modules: " & _
        VBA.CStr(importedCount) & "; imported at: " & _
        VBA.Format$(VBA.Now, "dd.mm.yyyy HH:nn:ss")
    reloadSucceeded = True

CleanExit:
    If applicationStateChanged Then
        Application.EnableEvents = previousEnableEvents
        Application.ScreenUpdating = previousScreenUpdating
    End If
    If reloadSucceeded Then
        fn_Diagnostic_WriteLog "VBA_RELOAD_COMPLETED | Workbook=" & _
            targetWorkbookName & " | ModuleCount=" & VBA.CStr(importedCount) & _
            " | ElapsedMs=" & private_VbaReload_FormatElapsedMilliseconds(startedAt)
    End If
    Exit Sub

EH:
    reloadSucceeded = False
    Application.StatusBar = False
    fn_Diagnostic_WriteLog "VBA_RELOAD_ERROR | Workbook=" & targetWorkbookName & _
        " | Number=" & VBA.CStr(VBA.Err.Number) & _
        " | Description=" & VBA.Err.Description & _
        " | ElapsedMs=" & private_VbaReload_FormatElapsedMilliseconds(startedAt)
    VBA.MsgBox "Failed to reload VBA modules in workbook '" & _
        targetWorkbookName & _
        "': [" & VBA.CStr(VBA.Err.Number) & "] " & VBA.Err.Description & VBA.vbCrLf & _
        "Make sure 'Trust access to the VBA project object model' is enabled.", _
        VBA.vbCritical, "Reload VBA"
    Resume CleanExit
End Sub


' Removes standard modules and class modules from the active workbook.
Public Sub fn_ClearActiveWorkbookVba()
    private_VbaReload_RequestSafeReload "ex_WorkbookUpdater.fn_RequestClear"
End Sub
' --------------------------------------
' } // namespace API
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
    Set fileSystem = CreateObject("Scripting.FileSystemObject")
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
' --------------------------------------
' } // namespace Diagnostic
' --------------------------------------

' --------------------------------------
' namespace HotkeyBroker {
' --------------------------------------
Private Sub private_HotkeyBroker_EnsureRegistry()
    If Not m_hotkeyBrokerBindingsByWorkbook Is Nothing Then Exit Sub

    Set m_hotkeyBrokerBindingsByWorkbook = VBA.CreateObject("Scripting.Dictionary")
    m_hotkeyBrokerBindingsByWorkbook.CompareMode = VBA.vbTextCompare
End Sub


Private Sub private_HotkeyBroker_ClearLegacyBindings()
    If m_hotkeyBrokerLegacyCleanupCompleted Then Exit Sub

    On Error Resume Next
    Application.OnKey "^%l"
    On Error GoTo 0
    m_hotkeyBrokerLegacyCleanupCompleted = True
    fn_Diagnostic_WriteLog "PERSONAL_HOTKEY_BROKER_LEGACY_KEY_CLEARED | Key=^%l"
End Sub


Private Function private_HotkeyBroker_FindWorkbook( _
    ByVal workbookFullName As String _
) As Workbook
    Dim targetWorkbook As Workbook

    For Each targetWorkbook In Application.Workbooks
        If VBA.StrComp(targetWorkbook.FullName, workbookFullName, _
                VBA.vbTextCompare) = 0 Then
            Set private_HotkeyBroker_FindWorkbook = targetWorkbook
            Exit Function
        End If
    Next targetWorkbook
End Function


Private Function private_HotkeyBroker_TryParseBindings( _
    ByVal bindingsText As String, _
    ByRef outBindingsByKey As Object _
) As Boolean
    Dim entryText As Variant
    Dim entryParts As Variant
    Dim keySequence As String
    Dim macroName As String

    Set outBindingsByKey = VBA.CreateObject("Scripting.Dictionary")
    outBindingsByKey.CompareMode = VBA.vbTextCompare
    If VBA.Len(bindingsText) = 0 Then
        private_HotkeyBroker_TryParseBindings = True
        Exit Function
    End If
    For Each entryText In VBA.Split(bindingsText, VBA.ChrW$(31))
        entryParts = VBA.Split(VBA.CStr(entryText), VBA.ChrW$(30))
        If UBound(entryParts) <> 1 Then Exit Function
        keySequence = VBA.Trim$(VBA.CStr(entryParts(0)))
        macroName = VBA.Trim$(VBA.CStr(entryParts(1)))
        If VBA.Len(keySequence) = 0 Or VBA.Len(macroName) = 0 Then Exit Function
        outBindingsByKey(keySequence) = macroName
    Next entryText
    private_HotkeyBroker_TryParseBindings = True
End Function


Private Sub private_HotkeyBroker_BindWorkbook( _
    ByVal targetWorkbook As Workbook, _
    ByVal bindingsByKey As Object _
)
    Dim keySequence As Variant
    Dim macroReference As String

    For Each keySequence In bindingsByKey.Keys
        macroReference = "'" & VBA.Replace$(targetWorkbook.Name, "'", "''") & _
            "'!" & VBA.CStr(bindingsByKey(keySequence))
        Application.OnKey VBA.CStr(keySequence), macroReference
    Next keySequence
End Sub


Private Sub private_HotkeyBroker_UnbindAll()
    Dim workbookIdentity As Variant
    Dim workbookIdentities As Collection

    If m_hotkeyBrokerBindingsByWorkbook Is Nothing Then Exit Sub
    Set workbookIdentities = New Collection
    For Each workbookIdentity In m_hotkeyBrokerBindingsByWorkbook.Keys
        workbookIdentities.Add VBA.CStr(workbookIdentity)
    Next workbookIdentity
    For Each workbookIdentity In workbookIdentities
        private_HotkeyBroker_UnbindWorkbook VBA.CStr(workbookIdentity)
    Next workbookIdentity
End Sub


Private Sub private_HotkeyBroker_UnbindWorkbook(ByVal workbookFullName As String)
    Dim bindingsByKey As Object
    Dim keySequence As Variant

    If m_hotkeyBrokerBindingsByWorkbook Is Nothing Then Exit Sub
    If Not m_hotkeyBrokerBindingsByWorkbook.Exists(workbookFullName) Then Exit Sub
    Set bindingsByKey = m_hotkeyBrokerBindingsByWorkbook(workbookFullName)
    On Error Resume Next
    For Each keySequence In bindingsByKey.Keys
        Application.OnKey VBA.CStr(keySequence)
    Next keySequence
    On Error GoTo 0
    m_hotkeyBrokerBindingsByWorkbook.Remove workbookFullName
End Sub
' --------------------------------------
' } // namespace HotkeyBroker
' --------------------------------------

' --------------------------------------
' namespace VbaReload {
' --------------------------------------
' Перезагружает VBA-проект во внешнем процессе Excel.
Private Sub private_VbaReload_RequestSafeReload(Optional ByVal requestMethod As String = "ex_WorkbookUpdater.fn_RequestReload")
    Dim targetWorkbook As Workbook
    Dim updaterWorkbook As Workbook
    Dim vbaFolderPath As String
    Dim updaterPath As String
    Dim fileSystem As Object
    On Error GoTo EH
    Set targetWorkbook = Application.ActiveWorkbook
    If targetWorkbook Is Nothing Then Err.Raise vbObjectError + 2400, , "No active workbook."
    If targetWorkbook Is ThisWorkbook Then Err.Raise vbObjectError + 2401, , "Activate the target workbook first."
    If Not private_VbaReload_TryResolveConfiguredFolder( _
            targetWorkbook, "ThisWorkbook::vbaPath", vbaFolderPath) Then Exit Sub
    Set fileSystem = CreateObject("Scripting.FileSystemObject")
    updaterPath = fileSystem.GetParentFolderName(vbaFolderPath) & "\WorkbookUpdater\WorkbookUpdater.xlam"
    If Not fileSystem.FileExists(updaterPath) Then _
        Err.Raise vbObjectError + 2402, , "Build WorkbookUpdater.xlam first: " & updaterPath
    On Error Resume Next
    Set updaterWorkbook = Application.Workbooks("WorkbookUpdater.xlam")
    On Error GoTo EH
    If updaterWorkbook Is Nothing Then Set updaterWorkbook = Application.Workbooks.Open(updaterPath)
    If StrComp(updaterWorkbook.FullName, updaterPath, vbTextCompare) <> 0 Then _
        Err.Raise vbObjectError + 2403, , "A different WorkbookUpdater.xlam is already loaded."
    ' Передаём конкретную книгу: активное окно может измениться до OnTime.
    Application.Run "'WorkbookUpdater.xlam'!" & requestMethod, targetWorkbook
    Exit Sub
EH:
    MsgBox "Could not request VBA reload: " & Err.Description, vbExclamation, "Workbook updater"
End Sub


' Экранирует один аргумент командной строки Windows.
Private Function private_VbaReload_QuoteCommandArgument(ByVal valueText As String) As String
    private_VbaReload_QuoteCommandArgument = """" & _
        VBA.Replace$(valueText, """", """""") & """"
End Function


' Проверяет доступность файла без исключения для вызывающего кода.
Private Function private_VbaReload_FileExists(ByVal filePath As String) As Boolean
    Dim fileSystem As Object

    Set fileSystem = VBA.CreateObject("Scripting.FileSystemObject")
    private_VbaReload_FileExists = fileSystem.FileExists(filePath)
End Function


' Reinitializes runtime event handlers after a hot reload.
Private Sub private_VbaReload_InitializeReloadedWorkbook( _
    ByVal targetWorkbook As Workbook, _
    ByVal deferInitialization As Boolean _
)
    Dim lifecycleComponent As Object
    Dim initializerName As String
    Dim macroReference As String

    On Error Resume Next
    Set lifecycleComponent = targetWorkbook.VBProject.VBComponents( _
        "rt_Lifecycle")
    On Error GoTo EH
    If Not lifecycleComponent Is Nothing Then
        If lifecycleComponent.CodeModule.CountOfLines > 0 Then
            macroReference = "'" & VBA.Replace$(targetWorkbook.Name, "'", "''") & _
                "'!rt_Lifecycle.fn_InitializeRuntime"
            If deferInitialization Then
                Application.OnTime VBA.Now + VBA.TimeSerial(0, 0, 1), macroReference
                fn_Diagnostic_WriteLog "VBA_RELOAD_INITIALIZE_SCHEDULED | Macro=" & _
                    macroReference
            Else
                Application.Run macroReference, "source-reload"
                fn_Diagnostic_WriteLog "VBA_RELOAD_INITIALIZE_COMPLETED | Macro=" & _
                    macroReference
            End If
            Exit Sub
        End If
    End If
    If Not private_VbaReload_TryFindInitializerName(targetWorkbook, initializerName) Then Exit Sub
    macroReference = "'" & VBA.Replace$(targetWorkbook.Name, "'", "''") & _
        "'!" & initializerName & ".fn_Initialize"
    private_VbaReload_RunOrScheduleInitialization macroReference, deferInitialization
    Exit Sub
EH:
    VBA.Err.Raise VBA.Err.Number, "private_VbaReload_InitializeReloadedWorkbook", _
        "Failed to initialize reloaded workbook '" & targetWorkbook.Name & _
        "': " & VBA.Err.Description
End Sub

Private Function private_VbaReload_TryFindInitializerName( _
    ByVal targetWorkbook As Workbook, _
    ByRef outInitializerName As String _
) As Boolean
    Const VBEXT_CT_STD_MODULE As Long = 1
    Dim vbComponent As Object
    Dim sourceText As String

    outInitializerName = VBA.vbNullString
    For Each vbComponent In targetWorkbook.VBProject.VBComponents
        If CLng(vbComponent.Type) = VBEXT_CT_STD_MODULE Then
            If vbComponent.CodeModule.CountOfLines > 0 Then
                sourceText = vbComponent.CodeModule.Lines(1, vbComponent.CodeModule.CountOfLines)
                If VBA.InStr(1, sourceText, "Public Sub fn_Initialize", VBA.vbTextCompare) > 0 Or _
                   VBA.InStr(1, sourceText, "Public Function fn_Initialize", VBA.vbTextCompare) > 0 Then
                    If VBA.Len(outInitializerName) > 0 Then
                        VBA.MsgBox "More than one public fn_Initialize procedure was found.", _
                            VBA.vbExclamation, "Reload VBA"
                        Exit Function
                    End If
                    outInitializerName = VBA.CStr(vbComponent.Name)
                End If
            End If
        End If
    Next vbComponent
    private_VbaReload_TryFindInitializerName = VBA.Len(outInitializerName) > 0
End Function

Private Sub private_VbaReload_RunOrScheduleInitialization( _
    ByVal macroReference As String, _
    ByVal deferInitialization As Boolean _
)
    fn_Diagnostic_WriteLog "VBA_RELOAD_INITIALIZE | Macro=" & macroReference
    If deferInitialization Then
        Application.OnTime VBA.Now + VBA.TimeSerial(0, 0, 1), macroReference
        fn_Diagnostic_WriteLog "VBA_RELOAD_INITIALIZE_SCHEDULED | Macro=" & _
            macroReference
        Exit Sub
    End If
    Application.Run macroReference
    fn_Diagnostic_WriteLog "VBA_RELOAD_INITIALIZE_COMPLETED | Macro=" & macroReference
End Sub

' Workbook configuration resolution.
Private Function private_VbaReload_TryResolveConfiguredFolder( _
    ByVal targetWorkbook As Workbook, _
    ByVal keyName As String, _
    ByRef outFolderPath As String _
) As Boolean
    Dim configuredPath As String
    Dim fileSystem As Object

    outFolderPath = VBA.vbNullString
    If Not private_VbaReload_TryGetWorkbookConfigValue( _
            targetWorkbook, keyName, configuredPath) Then Exit Function
    configuredPath = VBA.Replace$(VBA.Trim$(configuredPath), "/", "\")
    If VBA.Len(configuredPath) = 0 Or VBA.InStr(configuredPath, ":") > 0 Or _
       VBA.Left$(configuredPath, 1) = "\" Then
        VBA.MsgBox "The configuration key '" & keyName & _
            "' must contain a relative folder path: " & configuredPath, _
            VBA.vbExclamation, "Reload VBA"
        Exit Function
    End If

    Set fileSystem = VBA.CreateObject("Scripting.FileSystemObject")
    outFolderPath = fileSystem.GetAbsolutePathName( _
        targetWorkbook.Path & Application.PathSeparator & configuredPath)
    If Not fileSystem.FolderExists(outFolderPath) Then
        VBA.MsgBox "The folder from configuration key '" & keyName & _
            "' was not found: " & outFolderPath, _
            VBA.vbExclamation, "Reload VBA"
        outFolderPath = VBA.vbNullString
        Exit Function
    End If
    private_VbaReload_TryResolveConfiguredFolder = True
End Function

Private Sub private_VbaReload_SetRuntimePaths( _
    ByVal targetWorkbook As Workbook, _
    ByVal uiFolderPath As String _
)
    Dim macroReference As String

    macroReference = "'" & VBA.Replace$(targetWorkbook.Name, "'", "''") & _
        "'!ex_RuntimePaths.fn_SetUiFolder"
    Application.Run macroReference, uiFolderPath
End Sub
Private Function private_VbaReload_TryCollectConfiguredVbaFiles( _
    ByVal vbaFolderPath As String, _
    ByVal targetWorkbook As Workbook, _
    ByRef outFiles As Collection, _
    ByRef outDocumentFiles As Collection _
) As Boolean
    Const CONFIG_FILE_NAME As String = "modules.json"

    Dim fileSystem As Object
    Dim configPath As String
    Dim profileName As String
    Dim configText As String
    Dim relativeFiles As Collection
    Dim relativeFile As Variant
    Dim sourcePaths As Collection
    Dim sourcePath As Variant
    Dim lowerName As String
    Dim importedPaths As Object

    Set fileSystem = VBA.CreateObject("Scripting.FileSystemObject")
    configPath = vbaFolderPath & Application.PathSeparator & CONFIG_FILE_NAME
    If Not fileSystem.FileExists(configPath) Then
        VBA.MsgBox "VBA import configuration was not found: " & configPath, _
            VBA.vbExclamation, "Reload VBA"
        Exit Function
    End If

    If Not private_VbaReload_TryGetWorkbookProfileId(targetWorkbook, profileName) Then Exit Function
    configText = private_VbaReload_ReadUtf8TextFile(configPath)
    Set relativeFiles = New Collection
    If Not private_VbaReload_TryReadJsonStringArray(configText, profileName, relativeFiles) Then
        VBA.MsgBox "No VBA import profile was found for workbook: " & _
            targetWorkbook.Name, VBA.vbExclamation, "Reload VBA"
        Exit Function
    End If
    If relativeFiles.Count = 0 Then
        VBA.MsgBox "The VBA import profile is empty for workbook: " & _
            targetWorkbook.Name, VBA.vbExclamation, "Reload VBA"
        Exit Function
    End If

    Set importedPaths = VBA.CreateObject("Scripting.Dictionary")
    importedPaths.CompareMode = VBA.vbTextCompare
    For Each relativeFile In relativeFiles
        Set sourcePaths = New Collection
        If Not private_VbaReload_TryExpandConfiguredSourcePaths( _
                vbaFolderPath, VBA.CStr(relativeFile), sourcePaths) Then Exit Function
        For Each sourcePath In sourcePaths
            If importedPaths.Exists(sourcePath) Then
                VBA.MsgBox "The VBA import profile contains a duplicate module: " & _
                    VBA.CStr(relativeFile), VBA.vbExclamation, "Reload VBA"
                Exit Function
            End If
            importedPaths.Add sourcePath, True
            lowerName = VBA.LCase$(fileSystem.GetFileName(sourcePath))
            If private_VbaReload_IsDocumentModuleSource(lowerName) Then
                outDocumentFiles.Add sourcePath
            Else
                outFiles.Add sourcePath
            End If
        Next sourcePath
    Next relativeFile
    private_VbaReload_TryCollectConfiguredVbaFiles = True
End Function

Private Function private_VbaReload_TryGetWorkbookProfileId( _
    ByVal targetWorkbook As Workbook, _
    ByRef outProfileId As String _
) As Boolean
    private_VbaReload_TryGetWorkbookProfileId = _
        private_VbaReload_TryGetWorkbookConfigValue( _
            targetWorkbook, "ThisWorkbook::id", outProfileId)
End Function

Private Function private_VbaReload_TryGetWorkbookConfigValue( _
    ByVal targetWorkbook As Workbook, _
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
    On Error GoTo MissingConfiguration
    For Each targetWorksheet In targetWorkbook.Worksheets
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
                        private_VbaReload_TryGetWorkbookConfigValue = VBA.Len(outValue) > 0
                        Exit Function
                    End If
                Next configRow
            End If
ContinueTable:
        Next configTable
    Next targetWorksheet
MissingConfiguration:
    VBA.MsgBox "Configuration table '" & CONFIG_TABLE_NAME & _
        "' must contain key '" & keyName & _
        "' with a value in the next column.", _
        VBA.vbExclamation, "Reload VBA"
End Function
' Expands * and ? patterns recursively under the workbook profile vba folder.
Private Function private_VbaReload_TryExpandConfiguredSourcePaths( _
    ByVal vbaFolderPath As String, _
    ByVal configuredPath As String, _
    ByRef outPaths As Collection _
) As Boolean
    Dim fileSystem As Object
    Dim normalizedPattern As String
    Dim matchedPaths As Object
    Dim sortedPaths() As String
    Dim pathIndex As Long
    Dim pathKey As Variant

    If VBA.InStr(configuredPath, "*") = 0 And VBA.InStr(configuredPath, "?") = 0 Then
        If Not private_VbaReload_TryResolveConfiguredSourcePath( _
                vbaFolderPath, configuredPath, configuredPath) Then Exit Function
        outPaths.Add configuredPath
        private_VbaReload_TryExpandConfiguredSourcePaths = True
        Exit Function
    End If

    Set fileSystem = VBA.CreateObject("Scripting.FileSystemObject")
    normalizedPattern = VBA.Replace$(VBA.Trim$(configuredPath), "/", "\")
    If Not private_VbaReload_IsValidRelativeModulePattern(normalizedPattern) Then
        VBA.MsgBox "Invalid relative VBA module pattern: " & configuredPath, _
            VBA.vbExclamation, "Reload VBA"
        Exit Function
    End If

    Set matchedPaths = VBA.CreateObject("Scripting.Dictionary")
    matchedPaths.CompareMode = VBA.vbTextCompare
    private_VbaReload_CollectMatchingModulePaths _
        fileSystem.GetAbsolutePathName(vbaFolderPath), _
        fileSystem.GetAbsolutePathName(vbaFolderPath), normalizedPattern, matchedPaths
    If matchedPaths.Count = 0 Then
        private_VbaReload_TryExpandConfiguredSourcePaths = True
        Exit Function
    End If

    ReDim sortedPaths(1 To matchedPaths.Count)
    For Each pathKey In matchedPaths.Keys
        pathIndex = pathIndex + 1
        sortedPaths(pathIndex) = VBA.CStr(pathKey)
    Next pathKey
    private_VbaReload_SortTextArray sortedPaths
    For pathIndex = LBound(sortedPaths) To UBound(sortedPaths)
        outPaths.Add sortedPaths(pathIndex)
    Next pathIndex
    private_VbaReload_TryExpandConfiguredSourcePaths = True
End Function


' Collects matching module files from a folder and every nested folder.
Private Sub private_VbaReload_CollectMatchingModulePaths( _
    ByVal rootFolderPath As String, _
    ByVal folderPath As String, _
    ByVal relativePattern As String, _
    ByVal matchedPaths As Object _
)
    Dim fileSystem As Object
    Dim folderObject As Object
    Dim fileObject As Object
    Dim childFolder As Object
    Dim relativePath As String
    Dim candidateText As String

    Set fileSystem = VBA.CreateObject("Scripting.FileSystemObject")
    Set folderObject = fileSystem.GetFolder(folderPath)
    For Each fileObject In folderObject.Files
        If private_VbaReload_IsVbaModulePath(VBA.CStr(fileObject.Path)) Then
            relativePath = VBA.Mid$(VBA.CStr(fileObject.Path), VBA.Len(rootFolderPath) + 2)
            candidateText = VBA.CStr(fileObject.Name)
            If VBA.InStr(relativePattern, "\") > 0 Then candidateText = relativePath
            If VBA.LCase$(candidateText) Like VBA.LCase$(relativePattern) Then _
                matchedPaths.Add VBA.CStr(fileObject.Path), True
        End If
    Next fileObject
    For Each childFolder In folderObject.SubFolders
        private_VbaReload_CollectMatchingModulePaths _
            rootFolderPath, VBA.CStr(childFolder.Path), relativePattern, matchedPaths
    Next childFolder
End Sub


' Validates a wildcard pattern before scanning the workbook profile folder.
Private Function private_VbaReload_IsValidRelativeModulePattern( _
    ByVal relativePattern As String _
) As Boolean
    If VBA.Len(relativePattern) = 0 Or VBA.InStr(relativePattern, "..") > 0 Or _
       VBA.InStr(relativePattern, ":") > 0 Or VBA.Left$(relativePattern, 1) = "\" Then Exit Function
    If Not private_VbaReload_IsVbaModulePath(relativePattern) Then Exit Function
    private_VbaReload_IsValidRelativeModulePattern = True
End Function


' Returns whether a file name or path uses a supported VBA module extension.
Private Function private_VbaReload_IsVbaModulePath(ByVal filePath As String) As Boolean
    Dim lowerPath As String

    lowerPath = VBA.LCase$(filePath)
    private_VbaReload_IsVbaModulePath = _
        VBA.Right$(lowerPath, 4) = ".bas" Or _
        VBA.Right$(lowerPath, 4) = ".frm" Or _
        VBA.Right$(lowerPath, 4) = ".vba"
End Function


' Sorts paths so wildcard imports produce the same component order every time.
Private Sub private_VbaReload_SortTextArray(ByRef values() As String)
    Dim leftIndex As Long
    Dim rightIndex As Long
    Dim temporaryValue As String

    For leftIndex = LBound(values) To UBound(values) - 1
        For rightIndex = leftIndex + 1 To UBound(values)
            If VBA.StrComp(values(leftIndex), values(rightIndex), VBA.vbTextCompare) > 0 Then
                temporaryValue = values(leftIndex)
                values(leftIndex) = values(rightIndex)
                values(rightIndex) = temporaryValue
            End If
        Next rightIndex
    Next leftIndex
End Sub


' Only relative paths inside the workbook profile vba folder are allowed.
Private Function private_VbaReload_TryResolveConfiguredSourcePath( _
    ByVal vbaFolderPath As String, _
    ByVal relativePath As String, _
    ByRef outSourcePath As String _
) As Boolean
    Dim fileSystem As Object
    Dim normalizedRoot As String
    Dim normalizedPath As String
    Dim lowerName As String

    Set fileSystem = VBA.CreateObject("Scripting.FileSystemObject")
    relativePath = VBA.Replace$(VBA.Trim$(relativePath), "/", "\")
    If VBA.Len(relativePath) = 0 Or VBA.InStr(relativePath, "..") > 0 Or _
       VBA.InStr(relativePath, ":") > 0 Or VBA.Left$(relativePath, 1) = "\" Then
        VBA.MsgBox "Invalid relative VBA module path: " & relativePath, _
            VBA.vbExclamation, "Reload VBA"
        Exit Function
    End If

    normalizedRoot = fileSystem.GetAbsolutePathName(vbaFolderPath)
    normalizedPath = fileSystem.GetAbsolutePathName( _
        normalizedRoot & Application.PathSeparator & relativePath)
    If VBA.StrComp(VBA.Left$(normalizedPath, VBA.Len(normalizedRoot) + 1), _
            normalizedRoot & Application.PathSeparator, VBA.vbTextCompare) <> 0 Then
        VBA.MsgBox "VBA module path is outside the vba folder: " & relativePath, _
            VBA.vbExclamation, "Reload VBA"
        Exit Function
    End If
    If Not fileSystem.FileExists(normalizedPath) Then
        VBA.MsgBox "Configured VBA module was not found: " & relativePath, _
            VBA.vbExclamation, "Reload VBA"
        Exit Function
    End If
    lowerName = VBA.LCase$(normalizedPath)
    If Not (VBA.Right$(lowerName, 4) = ".bas" Or _
            VBA.Right$(lowerName, 4) = ".frm" Or _
            VBA.Right$(lowerName, 4) = ".vba") Then
        VBA.MsgBox "Configured file is not a VBA module: " & relativePath, _
            VBA.vbExclamation, "Reload VBA"
        Exit Function
    End If
    outSourcePath = normalizedPath
    private_VbaReload_TryResolveConfiguredSourcePath = True
End Function


Private Sub private_VbaReload_ClearDocumentModules(ByVal targetWorkbook As Workbook)
    Const VBEXT_CT_DOCUMENT As Long = 100

    Dim vbComponent As Object

    For Each vbComponent In targetWorkbook.VBProject.VBComponents
        If CLng(vbComponent.Type) = VBEXT_CT_DOCUMENT Then
            If vbComponent.CodeModule.CountOfLines > 0 Then _
                vbComponent.CodeModule.DeleteLines 1, vbComponent.CodeModule.CountOfLines
        End If
    Next vbComponent
End Sub


' Minimal JSON parsing for modules.json.
Private Function private_VbaReload_TryReadJsonStringArray( _
    ByVal jsonText As String, _
    ByVal profileName As String, _
    ByRef outValues As Collection _
) As Boolean
    Dim keyToken As String
    Dim position As Long
    Dim currentChar As String
    Dim valueText As String

    keyToken = """" & profileName & """"
    position = VBA.InStr(1, jsonText, keyToken, VBA.vbBinaryCompare)
    If position = 0 Then Exit Function
    position = position + VBA.Len(keyToken)
    private_VbaReload_SkipJsonWhitespace jsonText, position
    If VBA.Mid$(jsonText, position, 1) <> ":" Then Exit Function
    position = position + 1
    private_VbaReload_SkipJsonWhitespace jsonText, position
    If VBA.Mid$(jsonText, position, 1) <> "[" Then Exit Function
    position = position + 1

    Do
        private_VbaReload_SkipJsonWhitespace jsonText, position
        currentChar = VBA.Mid$(jsonText, position, 1)
        If currentChar = "]" Then
            private_VbaReload_TryReadJsonStringArray = True
            Exit Function
        End If
        If outValues.Count > 0 Then
            If currentChar <> "," Then Exit Function
            position = position + 1
            private_VbaReload_SkipJsonWhitespace jsonText, position
        End If
        If Not private_VbaReload_TryReadJsonString(jsonText, position, valueText) Then Exit Function
        outValues.Add valueText
    Loop
End Function


Private Sub private_VbaReload_SkipJsonWhitespace(ByVal jsonText As String, ByRef position As Long)
    Do While position <= VBA.Len(jsonText)
        Select Case VBA.Mid$(jsonText, position, 1)
            Case " ", VBA.vbTab, VBA.vbCr, VBA.vbLf
                position = position + 1
            Case Else
                Exit Do
        End Select
    Loop
End Sub


' Reads one string property from a small JSON object.
Private Function private_VbaReload_TryReadJsonString( _
    ByVal jsonText As String, _
    ByRef position As Long, _
    ByRef outValue As String _
) As Boolean
    Dim currentChar As String
    Dim escapedChar As String

    outValue = VBA.vbNullString
    If VBA.Mid$(jsonText, position, 1) <> """" Then Exit Function
    position = position + 1
    Do While position <= VBA.Len(jsonText)
        currentChar = VBA.Mid$(jsonText, position, 1)
        If currentChar = """" Then
            position = position + 1
            private_VbaReload_TryReadJsonString = True
            Exit Function
        End If
        If currentChar = "\" Then
            position = position + 1
            escapedChar = VBA.Mid$(jsonText, position, 1)
            Select Case escapedChar
                Case """", "\", "/"
                    outValue = outValue & escapedChar
                Case "b"
                    outValue = outValue & VBA.ChrW$(8)
                Case "f"
                    outValue = outValue & VBA.ChrW$(12)
                Case "n"
                    outValue = outValue & VBA.vbLf
                Case "r"
                    outValue = outValue & VBA.vbCr
                Case "t"
                    outValue = outValue & VBA.vbTab
                Case Else
                    Exit Function
            End Select
        Else
            outValue = outValue & currentChar
        End If
        position = position + 1
    Loop
End Function


Private Function private_VbaReload_IsDocumentModuleSource( _
    ByVal lowerFileName As String _
) As Boolean
    private_VbaReload_IsDocumentModuleSource = ( _
        private_VbaReload_IsThisWorkbookModuleSource(lowerFileName) Or _
        (VBA.Left$(lowerFileName, 3) = "ws_" And _
         private_VbaReload_IsVbaSourceFile(lowerFileName)))
End Function


' VBA project import operations.
Private Sub private_VbaReload_ImportVbaFile( _
    ByVal vbProject As Object, _
    ByVal sourcePath As String _
)
    Dim lowerPath As String
    Dim componentType As Long
    Dim sourceText As String
    Dim componentName As String
    Dim vbComponent As Object

    On Error GoTo EH

    lowerPath = VBA.LCase$(sourcePath)
    If VBA.Right$(lowerPath, 4) <> ".vba" Then
        vbProject.VBComponents.Import sourcePath
        Exit Sub
    End If

    sourceText = private_VbaReload_ReadUtf8TextFile(sourcePath)
    componentName = private_VbaReload_GetComponentName(sourcePath, sourceText)
    componentType = private_VbaReload_GetVbaComponentType(lowerPath, sourceText)

    Set vbComponent = vbProject.VBComponents.Add(componentType)
    vbComponent.Name = componentName
    vbComponent.CodeModule.AddFromString private_VbaReload_PrepareSourceForVbe(sourceText)
    Exit Sub

EH:
    VBA.Err.Raise VBA.Err.Number, "private_VbaReload_ImportVbaFile", _
        "Failed to import '" & sourcePath & "': " & VBA.Err.Description
End Sub


' Replaces source code in an existing workbook or worksheet module.
Private Sub private_VbaReload_ImportDocumentVbaFile( _
    ByVal targetWorkbook As Workbook, _
    ByVal sourcePath As String _
)
    Dim fileSystem As Object
    Dim fileName As String
    Dim worksheetName As String
    Dim targetSheet As Worksheet
    Dim componentName As String
    Dim vbComponent As Object
    Dim sourceText As String

    On Error GoTo EH
    Set fileSystem = VBA.CreateObject("Scripting.FileSystemObject")
    fileName = VBA.CStr(fileSystem.GetFileName(sourcePath))
    sourceText = private_VbaReload_ReadUtf8TextFile(sourcePath)
    If VBA.Len(sourceText) = 0 Then
        VBA.Err.Raise VBA.vbObjectError + 1201, _
            "private_VbaReload_ImportDocumentVbaFile", "Source is empty: " & sourcePath
    End If

    If private_VbaReload_IsThisWorkbookModuleSource(fileName) Then
        componentName = targetWorkbook.CodeName
    Else
        worksheetName = private_VbaReload_GetDocumentModuleSourceName(fileName, "ws_")
        On Error Resume Next
        Set targetSheet = targetWorkbook.Worksheets(worksheetName)
        On Error GoTo EH
        If targetSheet Is Nothing Then
            VBA.Err.Raise VBA.vbObjectError + 1202, _
                "private_VbaReload_ImportDocumentVbaFile", _
                "Worksheet '" & worksheetName & "' was not found for " & sourcePath
        End If
        componentName = targetSheet.CodeName
    End If

    Set vbComponent = targetWorkbook.VBProject.VBComponents(componentName)
    With vbComponent.CodeModule
        If .CountOfLines > 0 Then .DeleteLines 1, .CountOfLines
        .AddFromString private_VbaReload_PrepareSourceForVbe(sourceText)
    End With
    Exit Sub
EH:
    VBA.Err.Raise VBA.Err.Number, "private_VbaReload_ImportDocumentVbaFile", _
        "Failed to import document module '" & sourcePath & "': " & _
        VBA.Err.Description
End Sub


Private Function private_VbaReload_GetVbaComponentType( _
    ByVal lowerPath As String, _
    ByVal sourceText As String _
) As Long
    Const VBEXT_CT_STD_MODULE As Long = 1
    Const VBEXT_CT_CLASS_MODULE As Long = 2
    Const VBEXT_CT_MS_FORM As Long = 3

    If VBA.Right$(lowerPath, 8) = ".cls.vba" Or _
       VBA.Right$(lowerPath, 13) = ".cls.utf8.vba" Then
        private_VbaReload_GetVbaComponentType = VBEXT_CT_CLASS_MODULE
    ElseIf VBA.Right$(lowerPath, 8) = ".frm.vba" Or _
           VBA.Right$(lowerPath, 13) = ".frm.utf8.vba" Or _
           VBA.InStr(1, sourceText, "BEGIN VB.Form", VBA.vbTextCompare) > 0 Then
        private_VbaReload_GetVbaComponentType = VBEXT_CT_MS_FORM
    Else
        private_VbaReload_GetVbaComponentType = VBEXT_CT_STD_MODULE
    End If
End Function


Private Function private_VbaReload_GetComponentName( _
    ByVal sourcePath As String, _
    ByVal sourceText As String _
) As String
    Dim attributePosition As Long
    Dim nameStart As Long
    Dim nameEnd As Long
    Dim fileSystem As Object
    Dim fileName As String

    attributePosition = VBA.InStr(1, sourceText, "Attribute VB_Name = """, VBA.vbTextCompare)
    If attributePosition > 0 Then
        nameStart = attributePosition + VBA.Len("Attribute VB_Name = """)
        nameEnd = VBA.InStr(nameStart, sourceText, """")
        If nameEnd > nameStart Then
            private_VbaReload_GetComponentName = VBA.Mid$(sourceText, nameStart, nameEnd - nameStart)
            Exit Function
        End If
    End If

    Set fileSystem = VBA.CreateObject("Scripting.FileSystemObject")
    fileName = fileSystem.GetBaseName(sourcePath)
    If VBA.Right$(VBA.LCase$(fileName), 5) = ".utf8" Then
        fileName = VBA.Left$(fileName, VBA.Len(fileName) - VBA.Len(".utf8"))
    End If
    If VBA.Right$(VBA.LCase$(fileName), 4) = ".cls" Or _
        VBA.Right$(VBA.LCase$(fileName), 4) = ".bas" Or _
        VBA.Right$(VBA.LCase$(fileName), 4) = ".frm" Then
        fileName = VBA.Left$(fileName, VBA.Len(fileName) - 4)
    End If
    private_VbaReload_GetComponentName = fileName
End Function


Private Function private_VbaReload_FormatElapsedMilliseconds( _
    ByVal startedAt As Double _
) As String
    Dim elapsedSeconds As Double

    elapsedSeconds = VBA.Timer - startedAt
    If elapsedSeconds < 0 Then elapsedSeconds = elapsedSeconds + 86400#
    private_VbaReload_FormatElapsedMilliseconds = _
        VBA.Format$(elapsedSeconds * 1000#, "0.0")
End Function


' Source preparation for VBE import.
Private Function private_VbaReload_RemoveExportMetadata(ByVal sourceText As String) As String
    Dim lines As Variant
    Dim line As Variant
    Dim result As String
    Dim insideHeaderBlock As Boolean
    Dim trimmedLine As String

    lines = VBA.Split(VBA.Replace$(sourceText, VBA.vbCrLf, VBA.vbLf), VBA.vbLf)
    For Each line In lines
        trimmedLine = VBA.Trim$(VBA.CStr(line))
        If VBA.StrComp(trimmedLine, "VERSION 1.0 CLASS", VBA.vbTextCompare) = 0 Then
            insideHeaderBlock = True
        ElseIf insideHeaderBlock And VBA.StrComp(trimmedLine, "END", VBA.vbTextCompare) = 0 Then
            insideHeaderBlock = False
        ElseIf Not insideHeaderBlock And _
            VBA.Left$(trimmedLine, 10) <> "Attribute " Then
            result = result & VBA.CStr(line) & VBA.vbCrLf
        End If
    Next line
    private_VbaReload_RemoveExportMetadata = result
End Function


Private Function private_VbaReload_ReadUtf8TextFile(ByVal filePath As String) As String
    Const AD_TYPE_TEXT As Long = 2
    Const AD_READ_ALL As Long = -1

    Dim textStream As Object

    ' Project sources are stored in UTF-8. OpenTextFile with the system ANSI
    ' encoding corrupts non-ASCII text.
    Set textStream = VBA.CreateObject("ADODB.Stream")
    textStream.Type = AD_TYPE_TEXT
    textStream.Charset = "utf-8"
    textStream.Open
    textStream.LoadFromFile filePath
    private_VbaReload_ReadUtf8TextFile = textStream.ReadText(AD_READ_ALL)
    textStream.Close
    If VBA.Left$(private_VbaReload_ReadUtf8TextFile, 1) = VBA.ChrW$(65279) Then
        private_VbaReload_ReadUtf8TextFile = VBA.Mid$(private_VbaReload_ReadUtf8TextFile, 2)
    End If
End Function


' A .utf8.vba source must be read as UTF-8 and passed to
' CodeModule.AddFromString as a Unicode String without VBComponents.Import.
Private Function private_VbaReload_IsVbaSourceFile(ByVal lowerFileName As String) As Boolean
    private_VbaReload_IsVbaSourceFile = (VBA.Right$(lowerFileName, 4) = ".vba")
End Function


Private Function private_VbaReload_IsThisWorkbookModuleSource(ByVal fileName As String) As Boolean
    private_VbaReload_IsThisWorkbookModuleSource = ( _
        VBA.StrComp(fileName, "ThisWorkbook.vba", VBA.vbTextCompare) = 0 Or _
        VBA.StrComp(fileName, "ThisWorkbook.utf8.vba", VBA.vbTextCompare) = 0)
End Function


Private Function private_VbaReload_GetDocumentModuleSourceName( _
    ByVal fileName As String, _
    ByVal requiredPrefix As String _
) As String
    Dim sourceStem As String

    sourceStem = fileName
    If VBA.Right$(VBA.LCase$(sourceStem), 9) = ".utf8.vba" Then
        sourceStem = VBA.Left$(sourceStem, VBA.Len(sourceStem) - VBA.Len(".utf8.vba"))
    ElseIf VBA.Right$(VBA.LCase$(sourceStem), 4) = ".vba" Then
        sourceStem = VBA.Left$(sourceStem, VBA.Len(sourceStem) - VBA.Len(".vba"))
    Else
        VBA.Err.Raise VBA.vbObjectError + 1203, _
            "private_VbaReload_GetDocumentModuleSourceName", _
            "Unsupported document module source extension: " & fileName
    End If
    If VBA.Left$(sourceStem, VBA.Len(requiredPrefix)) <> requiredPrefix Then
        VBA.Err.Raise VBA.vbObjectError + 1204, _
            "private_VbaReload_GetDocumentModuleSourceName", _
            "Document module source has invalid prefix: " & fileName
    End If
    private_VbaReload_GetDocumentModuleSourceName = VBA.Mid$(sourceStem, _
        VBA.Len(requiredPrefix) + 1)
End Function


Private Function private_VbaReload_FolderExists(ByVal folderPath As String) As Boolean
    Dim fileSystem As Object

    Set fileSystem = VBA.CreateObject("Scripting.FileSystemObject")
    private_VbaReload_FolderExists = fileSystem.FolderExists(folderPath)
End Function


' The classic VBE stores source code using the current system code page.
' Convert only non-ASCII string literals into ChrW$ expressions before adding
' code, so the imported module remains executable on every legacy code page.
Private Function private_VbaReload_PrepareSourceForVbe(ByVal sourceText As String) As String
    private_VbaReload_PrepareSourceForVbe = private_VbaReload_RemoveExportMetadata( _
        private_VbaReload_EncodeUnicodeStringLiterals(sourceText))
End Function


Private Function private_VbaReload_EncodeUnicodeStringLiterals( _
    ByVal sourceText As String _
) As String
    Dim sourceLines As Variant
    Dim sourceLine As Variant
    Dim resultText As String

    sourceLines = VBA.Split(VBA.Replace$(sourceText, VBA.vbCrLf, VBA.vbLf), _
        VBA.vbLf)
    For Each sourceLine In sourceLines
        resultText = resultText & private_VbaReload_EncodeUnicodeStringLiteralsOnLine( _
            VBA.CStr(sourceLine)) & VBA.vbCrLf
    Next sourceLine

    private_VbaReload_EncodeUnicodeStringLiterals = resultText
End Function


Private Function private_VbaReload_EncodeUnicodeStringLiteralsOnLine( _
    ByVal sourceLine As String _
) As String
    Dim charIndex As Long
    Dim literalStartIndex As Long
    Dim literalText As String
    Dim currentCharacter As String
    Dim resultText As String
    Dim literalClosed As Boolean

    charIndex = 1
    Do While charIndex <= VBA.Len(sourceLine)
        currentCharacter = VBA.Mid$(sourceLine, charIndex, 1)

        If currentCharacter = "'" Then
            ' A quote begins a comment outside a string literal.
            resultText = resultText & VBA.Mid$(sourceLine, charIndex)
            Exit Do
        End If

        If currentCharacter <> """" Then
            resultText = resultText & currentCharacter
            charIndex = charIndex + 1
        Else
            literalStartIndex = charIndex
            charIndex = charIndex + 1
            literalText = VBA.vbNullString
            literalClosed = False

            Do While charIndex <= VBA.Len(sourceLine)
                currentCharacter = VBA.Mid$(sourceLine, charIndex, 1)
                If currentCharacter <> """" Then
                    literalText = literalText & currentCharacter
                    charIndex = charIndex + 1
                ElseIf charIndex < VBA.Len(sourceLine) And _
                       VBA.Mid$(sourceLine, charIndex + 1, 1) = """" Then
                    literalText = literalText & """"
                    charIndex = charIndex + 2
                Else
                    charIndex = charIndex + 1
                    literalClosed = True
                    Exit Do
                End If
            Loop

            If Not literalClosed Then
                ' Keep an unterminated literal unchanged and let the VBE report it.
                resultText = resultText & VBA.Mid$(sourceLine, literalStartIndex)
                Exit Do
            End If

            resultText = resultText & private_VbaReload_EncodeUnicodeLiteral(literalText)
        End If
    Loop

    private_VbaReload_EncodeUnicodeStringLiteralsOnLine = resultText
End Function


Private Function private_VbaReload_EncodeUnicodeLiteral( _
    ByVal literalText As String _
) As String
    Dim charIndex As Long
    Dim characterCode As Long
    Dim asciiBuffer As String
    Dim expressionParts As Collection

    Set expressionParts = New Collection

    For charIndex = 1 To VBA.Len(literalText)
        characterCode = VBA.AscW(VBA.Mid$(literalText, charIndex, 1))
        If characterCode >= 0 And characterCode <= 127 Then
            asciiBuffer = asciiBuffer & VBA.Mid$(literalText, charIndex, 1)
        Else
            private_VbaReload_AppendAsciiLiteralPart expressionParts, asciiBuffer
            asciiBuffer = VBA.vbNullString
            expressionParts.Add "VBA.ChrW$(" & VBA.CStr(characterCode) & ")"
        End If
    Next charIndex
    private_VbaReload_AppendAsciiLiteralPart expressionParts, asciiBuffer

    private_VbaReload_EncodeUnicodeLiteral = private_VbaReload_JoinExpressionParts(expressionParts)
End Function


Private Sub private_VbaReload_AppendAsciiLiteralPart( _
    ByVal expressionParts As Collection, _
    ByVal asciiText As String _
)
    If VBA.Len(asciiText) = 0 Then Exit Sub

    expressionParts.Add """" & VBA.Replace$(asciiText, """", """""") & """"
End Sub


Private Function private_VbaReload_JoinExpressionParts( _
    ByVal expressionParts As Collection _
) As String
    Dim partIndex As Long
    Dim resultText As String

    For partIndex = 1 To expressionParts.Count
        If partIndex > 1 Then resultText = resultText & " & "
        resultText = resultText & VBA.CStr(expressionParts(partIndex))
    Next partIndex

    If expressionParts.Count = 0 Then resultText = """"
    private_VbaReload_JoinExpressionParts = resultText
End Function
' --------------------------------------
' } // namespace VbaReload
' --------------------------------------