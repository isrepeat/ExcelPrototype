Option Explicit

#Const ENABLE_LOGGING = True

Private Const DIAGNOSTIC_FOLDER_NAME As String = "PERSONAL.EXCEL"
Private Const DIAGNOSTIC_FILE_NAME As String = "diagnostic.log"
Private m_diagnosticSessionStarted As Boolean

' --------------------------------------
' namespace API {
' --------------------------------------
' Recreates PERSONAL.XLSB runtime state after source modules are reloaded.
Public Sub fn_ReloadPersonalRuntime()
    ThisWorkbook.fn_ReloadRuntime
End Sub

Public Sub fn_ReloadActiveWorkbookVba()
    private_ReloadActiveWorkbookVba False
End Sub

' Reloads the active workbook and schedules initialization after VBA reset.
Public Sub fn_ReloadActiveWorkbookVbaDeferred()
    private_ReloadActiveWorkbookVba True
End Sub

Private Sub private_ReloadActiveWorkbookVba( _
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
    Dim targetWorkbookName As String
    Dim importedCount As Long
    Dim previousEnableEvents As Boolean
    Dim previousScreenUpdating As Boolean
    Dim applicationStateChanged As Boolean
    Dim componentName As Variant
    Dim importFile As Variant

    On Error GoTo EH

    If Application.ActiveWorkbook Is Nothing Then
        VBA.MsgBox "There is no active workbook whose VBA modules can be reloaded.", _
            VBA.vbExclamation, "Reload VBA"
        Exit Sub
    End If

    Set targetWorkbook = Application.ActiveWorkbook
    targetWorkbookName = targetWorkbook.Name
    fn_Diagnostic_WriteLog "VBA_RELOAD_STARTED | Workbook=" & targetWorkbookName
    If targetWorkbook Is ThisWorkbook Then
        VBA.MsgBox _
            "The workbook containing the shortcut handler cannot reload itself. " & _
            "Activate the target workbook and press Ctrl+Alt+R again.", _
            VBA.vbExclamation, "Reload VBA"
        Exit Sub
    End If

    If VBA.Len(targetWorkbook.Path) = 0 Then
        VBA.MsgBox _
            "Save the active workbook first. The vba folder must be located " & _
            "next to the workbook file.", _
            VBA.vbExclamation, "Reload VBA"
        Exit Sub
    End If

    If Not private_VbaReload_TryResolveVbaFolder(targetWorkbook, vbaFolderPath) Then Exit Sub
    If Not private_VbaReload_TrySyncUiFiles(targetWorkbook, vbaFolderPath) Then Exit Sub

    Set importFiles = New Collection
    Set documentImportFiles = New Collection
    If Not private_VbaReload_TryCollectConfiguredVbaFiles( _
            vbaFolderPath, targetWorkbook, importFiles, documentImportFiles) Then Exit Sub
    If importFiles.Count = 0 And documentImportFiles.Count = 0 Then
        VBA.MsgBox _
            "No .bas, .cls, .frm, .vba, or .utf8.vba files were found in: " & vbaFolderPath, _
            VBA.vbExclamation, "Reload VBA"
        Exit Sub
    End If

    Set vbProject = targetWorkbook.VBProject
    Set componentNames = New Collection
    For Each vbComponent In vbProject.VBComponents
        Select Case CLng(vbComponent.Type)
            Case VBEXT_CT_STD_MODULE, VBEXT_CT_CLASS_MODULE, VBEXT_CT_MS_FORM
                componentNames.Add VBA.CStr(vbComponent.Name)
        End Select
    Next vbComponent

    previousScreenUpdating = Application.ScreenUpdating
    previousEnableEvents = Application.EnableEvents
    applicationStateChanged = True
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.StatusBar = "Reloading VBA modules..."

    For Each componentName In componentNames
        vbProject.VBComponents.Remove vbProject.VBComponents(VBA.CStr(componentName))
    Next componentName

    ' The profile defines the complete source set. Clear workbook and worksheet
    ' code so that event handlers from the previous profile do not stay active.
    private_VbaReload_ClearDocumentModules targetWorkbook

    For Each importFile In importFiles
        private_VbaReload_ImportVbaFile vbProject, VBA.CStr(importFile)
        importedCount = importedCount + 1
    Next importFile

    ' Document modules cannot be imported as standard modules because Excel
    ' would create a separate module and Workbook/Worksheet events would not run.
    For Each importFile In documentImportFiles
        private_VbaReload_ImportDocumentVbaFile targetWorkbook, VBA.CStr(importFile)
        importedCount = importedCount + 1
    Next importFile
    private_VbaReload_InitializeReloadedWorkbook targetWorkbook, deferInitialization

    Application.StatusBar = "Imported VBA modules: " & _
        VBA.CStr(importedCount) & "; imported at: " & _
        VBA.Format$(VBA.Now, "dd.mm.yyyy HH:nn:ss")
    fn_Diagnostic_WriteLog "VBA_RELOAD_COMPLETED | Workbook=" & targetWorkbookName & _
        " | ModuleCount=" & VBA.CStr(importedCount)

CleanExit:
    If applicationStateChanged Then
        Application.EnableEvents = previousEnableEvents
        Application.ScreenUpdating = previousScreenUpdating
    End If
    Exit Sub

EH:
    Application.StatusBar = False
    fn_Diagnostic_WriteLog "VBA_RELOAD_ERROR | Workbook=" & targetWorkbookName & _
        " | Number=" & VBA.CStr(VBA.Err.Number) & _
        " | Description=" & VBA.Err.Description
    VBA.MsgBox "Failed to reload VBA modules in workbook '" & _
        targetWorkbookName & _
        "': [" & VBA.CStr(VBA.Err.Number) & "] " & VBA.Err.Description & VBA.vbCrLf & _
        "Make sure 'Trust access to the VBA project object model' is enabled.", _
        VBA.vbCritical, "Reload VBA"
    Resume CleanExit
End Sub


' Removes standard modules and class modules from the active workbook.
Public Sub fn_ClearActiveWorkbookVba()
    Const VBEXT_CT_STD_MODULE As Long = 1
    Const VBEXT_CT_CLASS_MODULE As Long = 2
    Const VBEXT_CT_DOCUMENT As Long = 100

    Dim targetWorkbook As Workbook
    Dim vbProject As Object
    Dim vbComponent As Object
    Dim componentNames As Collection
    Dim componentName As Variant
    Dim removedCount As Long
    Dim clearedDocumentModuleCount As Long

    On Error GoTo EH
    If Application.ActiveWorkbook Is Nothing Then
        VBA.MsgBox "There is no active workbook whose VBA modules can be cleared.", _
            VBA.vbExclamation, "Clear VBA"
        Exit Sub
    End If

    Set targetWorkbook = Application.ActiveWorkbook
    If targetWorkbook Is ThisWorkbook Then
        VBA.MsgBox "PERSONAL.XLSB cannot clear its own modules.", _
            VBA.vbExclamation, "Clear VBA"
        Exit Sub
    End If

    fn_Diagnostic_WriteLog "VBA_CLEAR_STARTED | Workbook=" & targetWorkbook.Name
    Set vbProject = targetWorkbook.VBProject
    Set componentNames = New Collection
    For Each vbComponent In vbProject.VBComponents
        Select Case CLng(vbComponent.Type)
            Case VBEXT_CT_STD_MODULE, VBEXT_CT_CLASS_MODULE
                componentNames.Add VBA.CStr(vbComponent.Name)
        End Select
    Next vbComponent

    For Each componentName In componentNames
        On Error Resume Next
        Application.Run "'" & VBA.Replace$(targetWorkbook.Name, "'", "''") & _
            "'!" & VBA.CStr(componentName) & ".fn_Module_Dispose"
        On Error GoTo EH
        vbProject.VBComponents.Remove vbProject.VBComponents(VBA.CStr(componentName))
        removedCount = removedCount + 1
    Next componentName

    For Each vbComponent In vbProject.VBComponents
        If CLng(vbComponent.Type) = VBEXT_CT_DOCUMENT Then
            If vbComponent.CodeModule.CountOfLines > 0 Then
                vbComponent.CodeModule.DeleteLines 1, vbComponent.CodeModule.CountOfLines
                clearedDocumentModuleCount = clearedDocumentModuleCount + 1
            End If
        End If
    Next vbComponent

    targetWorkbook.Save
    Application.StatusBar = "Removed VBA modules and classes: " & VBA.CStr(removedCount) & _
        "; cleared document modules: " & VBA.CStr(clearedDocumentModuleCount)
    fn_Diagnostic_WriteLog "VBA_CLEAR_COMPLETED | Workbook=" & targetWorkbook.Name & _
        " | ModuleCount=" & VBA.CStr(removedCount) & _
        " | ClearedDocumentModuleCount=" & VBA.CStr(clearedDocumentModuleCount)
    VBA.MsgBox "Removed VBA modules and classes: " & VBA.CStr(removedCount) & _
        VBA.vbCrLf & "Cleared document modules: " & _
        VBA.CStr(clearedDocumentModuleCount), _
        VBA.vbInformation, "Clear VBA"
    Exit Sub

EH:
    fn_Diagnostic_WriteLog "VBA_CLEAR_ERROR | Workbook=" & targetWorkbook.Name & _
        " | Number=" & VBA.CStr(VBA.Err.Number) & _
        " | Description=" & VBA.Err.Description
    VBA.MsgBox "Failed to clear VBA modules in workbook '" & targetWorkbook.Name & _
        "': [" & VBA.CStr(VBA.Err.Number) & "] " & VBA.Err.Description & VBA.vbCrLf & _
        "Make sure 'Trust access to the VBA project object model' is enabled.", _
        VBA.vbCritical, "Clear VBA"
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
' namespace VbaReload {
' --------------------------------------
' Reinitializes runtime event handlers after a hot reload.
Private Sub private_VbaReload_InitializeReloadedWorkbook( _
    ByVal targetWorkbook As Workbook, _
    ByVal deferInitialization As Boolean _
)
    Dim bootstrapComponent As Object
    Dim lifecycleComponent As Object
    Dim macroReference As String

    On Error Resume Next
    Set bootstrapComponent = targetWorkbook.VBProject.VBComponents( _
        "ex_PersonalEventBuilder")
    On Error GoTo EH
    If Not bootstrapComponent Is Nothing Then
        macroReference = "'" & VBA.Replace$(targetWorkbook.Name, "'", "''") & _
            "'!ex_PersonalEventBuilder.fn_Initialize"
        private_VbaReload_RunOrScheduleInitialization macroReference, deferInitialization
        Exit Sub
    End If
    On Error Resume Next
    Set lifecycleComponent = targetWorkbook.VBProject.VBComponents( _
        "rt_Lifecycle")
    Set bootstrapComponent = targetWorkbook.VBProject.VBComponents( _
        "ex_DocumentGenerationBootstrap")
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
    If bootstrapComponent Is Nothing Then Exit Sub
    If bootstrapComponent.CodeModule.CountOfLines = 0 Then Exit Sub

    macroReference = "'" & VBA.Replace$(targetWorkbook.Name, "'", "''") & _
        "'!ex_DocumentGenerationBootstrap.fn_Initialize"
    private_VbaReload_RunOrScheduleInitialization macroReference, deferInitialization
    Exit Sub
EH:
    VBA.Err.Raise VBA.Err.Number, "private_VbaReload_InitializeReloadedWorkbook", _
        "Failed to initialize reloaded workbook '" & targetWorkbook.Name & _
        "': " & VBA.Err.Description
End Sub

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

Private Function private_VbaReload_TrySyncUiFiles( _
    ByVal targetWorkbook As Workbook, _
    ByVal vbaFolderPath As String _
) As Boolean
    Dim fileSystem As Object
    Dim sourceUiPath As String
    Dim targetUiPath As String
    Dim uiFile As Object

    Set fileSystem = VBA.CreateObject("Scripting.FileSystemObject")
    sourceUiPath = fileSystem.GetParentFolderName(vbaFolderPath) & "\ui"
    If Not fileSystem.FolderExists(sourceUiPath) Then
        private_VbaReload_TrySyncUiFiles = True
        Exit Function
    End If
    targetUiPath = targetWorkbook.Path & "\ui"
    If Not fileSystem.FolderExists(targetUiPath) Then fileSystem.CreateFolder targetUiPath
    For Each uiFile In fileSystem.GetFolder(sourceUiPath).Files
        If VBA.LCase$(fileSystem.GetExtensionName(uiFile.Name)) = "xaml" Then _
            fileSystem.CopyFile uiFile.Path, targetUiPath & "\" & uiFile.Name, True
    Next uiFile
    private_VbaReload_TrySyncUiFiles = True
End Function


' Source folder and profile resolution.
Private Function private_VbaReload_TryResolveVbaFolder( _
    ByVal targetWorkbook As Workbook, _
    ByRef outVbaFolderPath As String _
) As Boolean
    Dim fileSystem As Object
    Dim candidatePaths As Collection
    Dim workbookFolderPath As String
    Dim workspaceRootPath As String
    Dim candidatePath As Variant

    outVbaFolderPath = VBA.vbNullString
    Set fileSystem = VBA.CreateObject("Scripting.FileSystemObject")
    Set candidatePaths = New Collection
    workbookFolderPath = targetWorkbook.Path

    candidatePath = workbookFolderPath & Application.PathSeparator & "vba"
    If private_VbaReload_FolderExists(VBA.CStr(candidatePath)) Then _
        candidatePaths.Add VBA.CStr(candidatePath)

    ' A workbook in MacrosExcel/<project> uses PROJECTS/<project>/vba.
    If VBA.StrComp(fileSystem.GetFileName( _
            fileSystem.GetParentFolderName(workbookFolderPath)), _
            "MacrosExcel", VBA.vbTextCompare) = 0 Then
        workspaceRootPath = fileSystem.GetParentFolderName( _
            fileSystem.GetParentFolderName(workbookFolderPath))
        candidatePath = workspaceRootPath & Application.PathSeparator & _
            "PROJECTS" & Application.PathSeparator & _
            fileSystem.GetFileName(workbookFolderPath) & _
            Application.PathSeparator & "vba"
        If private_VbaReload_FolderExists(VBA.CStr(candidatePath)) Then _
            candidatePaths.Add VBA.CStr(candidatePath)
    End If

    If candidatePaths.Count = 0 Then
        VBA.MsgBox "The VBA modules folder was not found beside the workbook " & _
            "or in PROJECTS\2. DocumentsGeneration\vba.", _
            VBA.vbExclamation, "Reload VBA"
        Exit Function
    End If
    If candidatePaths.Count > 1 Then
        VBA.MsgBox "More than one VBA source folder was found. Keep only one " & _
            "source location to avoid importing an unexpected version.", _
            VBA.vbExclamation, "Reload VBA"
        Exit Function
    End If

    outVbaFolderPath = VBA.CStr(candidatePaths.Item(1))
    private_VbaReload_TryResolveVbaFolder = True
End Function


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

    profileName = fileSystem.GetBaseName(targetWorkbook.Name)
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
        VBA.Right$(lowerPath, 4) = ".cls" Or _
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
            VBA.Right$(lowerName, 4) = ".cls" Or _
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
       VBA.Right$(lowerPath, 13) = ".cls.utf8.vba" Or _
        VBA.InStr(1, sourceText, "VERSION 1.0 CLASS", VBA.vbTextCompare) > 0 Then
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