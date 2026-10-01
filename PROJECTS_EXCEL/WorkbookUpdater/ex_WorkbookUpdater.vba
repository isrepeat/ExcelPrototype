Option Explicit

Private m_target As Workbook
Private m_context As Object
Private m_plan As Collection
Private m_uiFolder As String
Private m_scheduledAt As Date
Private m_deadline As Date
Private m_backupPath As String
Private m_operationFolder As String
Private m_showErrors As Boolean
Private m_lastResult As String
Private m_lastError As String
Private m_contexts As Object
Private m_initialInstall As Boolean
Private m_clearOnly As Boolean

' --------------------------------------
' namespace API {
' --------------------------------------
Public Function fn_RequestReload( _
    ByVal target As Workbook, _
    Optional ByVal showErrors As Boolean = True, _
    Optional ByVal clearOnly As Boolean = False _
) As Boolean
    Dim errorText As String
    Dim lifecycleStarted As Boolean
    Dim ownsOperation As Boolean

    On Error GoTo EH
    If Not m_target Is Nothing Then _
        VBA.Err.Raise VBA.vbObjectError + 2300, , "An update is already pending."
    If target Is ThisWorkbook Then _
        VBA.Err.Raise VBA.vbObjectError + 2301, , "The updater cannot reload itself."
    If VBA.Len(target.Path) = 0 Or target.ReadOnly Then _
        VBA.Err.Raise VBA.vbObjectError + 2302, , "A saved writable workbook is required."
    If target.VBProject.Protection <> 0 Then _
        VBA.Err.Raise VBA.vbObjectError + 2303, , "The VBA project is protected."
    Set m_target = target
    ownsOperation = True
    m_clearOnly = clearOnly
    m_showErrors = showErrors
    m_lastResult = "Preparing"
    m_lastError = VBA.vbNullString
    m_initialInstall = private_IsEmptyProject(target)
    If m_initialInstall Then
        If private_HasBlockedMarker(target) Then _
            VBA.Err.Raise VBA.vbObjectError + 2322, , _
            "The empty project is blocked after a failed update. Recover it before installing."
        Set m_context = fn_FindContext(target)
        If Not m_context Is Nothing Then
            If VBA.CLng(m_context("Generation")) <> 0 Then _
                VBA.Err.Raise VBA.vbObjectError + 2325, , _
            "An initialized runtime lost its code. Recover the workbook before installing."
        End If
        If m_context Is Nothing Then
            Set m_context = VBA.CreateObject("Scripting.Dictionary")
            m_context.Add "StopRequested", False
            m_context.Add "ActiveCalls", 0&
            m_context.Add "Phase", "Running"
            m_context.Add "Generation", 0&
        End If
    Else
        If Not private_HasLifecycle(target) Then _
            VBA.Err.Raise VBA.vbObjectError + 2323, , _
            "The project contains code but has no ex_RuntimeLifecycle. Only empty projects can skip shutdown."
        Set m_context = Application.Run(private_Macro("ex_RuntimeLifecycle.fn_Context"))
    End If
    If m_context("Phase") <> "Running" Or m_context("StopRequested") Then _
        VBA.Err.Raise VBA.vbObjectError + 2304, , "The runtime is already stopped or faulted."
    private_RegisterContext target, m_context

    ' Read the entire plan before shutdown; subsequent source edits do not change it.
    If m_clearOnly Then
        Set m_plan = New Collection
        m_operationFolder = private_CreateOperationFolder(target)
    Else
        Set m_plan = private_ReadPlan(target, m_uiFolder)
    End If
    m_context("StopRequested") = True
    m_context("Phase") = "Requested"
    lifecycleStarted = True
    m_deadline = VBA.Now + VBA.TimeSerial(0, 0, 30)
    private_Schedule
    m_lastResult = "Pending"
    fn_RequestReload = True
    Exit Function
EH:
    errorText = VBA.Err.Description
    If lifecycleStarted Then
        m_context("StopRequested") = False
        m_context("Phase") = "Running"
    End If
    ' A repeated request must not discard an existing queued operation.
    If ownsOperation And m_scheduledAt = 0 Then private_ClearOperation
    If ownsOperation Then
        m_lastError = errorText
        m_lastResult = "Rejected"
    End If
    If showErrors Then VBA.MsgBox "Reload request failed: " & errorText, VBA.vbExclamation, "Workbook updater"
End Function

Public Function fn_RequestClear( _
    ByVal target As Workbook, _
    Optional ByVal showErrors As Boolean = True _
) As Boolean
    fn_RequestClear = fn_RequestReload(target, showErrors, True)
End Function

Public Function fn_LastResult() As String
    fn_LastResult = m_lastResult
End Function

Public Function fn_LastError() As String
    fn_LastError = m_lastError
End Function

Public Function fn_CanUnload() As Boolean
    fn_CanUnload = (m_target Is Nothing) And m_scheduledAt = 0
End Function

Public Function fn_FindContext(ByVal target As Workbook) As Object
    Dim entry As Object
    Dim registeredBook As Workbook
    Dim key As Variant

    If m_contexts Is Nothing Then Exit Function
    For Each key In m_contexts.Keys
        Set entry = m_contexts(key)
        Set registeredBook = entry("Workbook")
        If registeredBook Is target Then
            Set fn_FindContext = entry("Context")
            Exit Function
        End If
    Next key
End Function

Public Sub fn_ForgetContext(ByVal target As Workbook)
    Dim entry As Object
    Dim registeredBook As Workbook
    Dim key As Variant

    If m_contexts Is Nothing Then Exit Sub
    For Each key In m_contexts.Keys
        Set entry = m_contexts(key)
        Set registeredBook = entry("Workbook")
        If registeredBook Is target Then
            fn_CancelPending target
            m_contexts.Remove key
            Exit Sub
        End If
    Next key
End Sub

Public Sub fn_RunPending()
    Dim previousEvents As Boolean
    Dim previousScreenUpdating As Boolean
    Dim previousCalculation As Long
    Dim stateCaptured As Boolean
    Dim errorText As String
    Dim clearedTarget As Workbook

    On Error GoTo EH
    m_scheduledAt = 0
    If m_target Is Nothing Then Exit Sub
    If Not private_IsTargetOpen() Then _
        VBA.Err.Raise VBA.vbObjectError + 2305, , "The target workbook was closed."
    ' During a callback, Excel reports Run even for the target project.
    ' Confirm completed calls using the counter rather than requiring Design mode.
    If VBA.CLng(m_context("ActiveCalls")) <> 0 Then
        m_lastError = "Waiting: project mode=" & VBA.CStr(m_target.VBProject.Mode) & "; active calls=" & VBA.CStr(m_context("ActiveCalls"))
        If VBA.Now >= m_deadline Then _
        VBA.Err.Raise VBA.vbObjectError + 2306, , "Runtime did not stop within 30 seconds."
        private_Schedule
        Exit Sub
    End If
    m_lastError = VBA.vbNullString
    previousEvents = Application.EnableEvents
    previousScreenUpdating = Application.ScreenUpdating
    previousCalculation = Application.Calculation
    stateCaptured = True
    Application.EnableEvents = False
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual

    If m_context("Phase") = "Initializing" Then GoTo InitializeRuntime

    ' Recheck whether the project is empty: it may have changed while the callback was pending.
    If m_initialInstall And Not private_IsEmptyProject(m_target) Then _
        VBA.Err.Raise VBA.vbObjectError + 2324, , _
            "The project is no longer empty. Initial installation was cancelled."
    m_backupPath = m_operationFolder & "\" & m_target.Name
    m_target.SaveCopyAs m_backupPath
    m_target.Names.Add Name:="_RuntimeReloadBlocked", RefersTo:="=TRUE", Visible:=False
    If m_initialInstall Then
        ' An empty project has no existing runtime; the new lifecycle arrives with the source modules.
        m_context("Phase") = "Prepared"
    Else
        If Not VBA.CBool(Application.Run(private_Macro("ex_RuntimeLifecycle.fn_PrepareReload"))) Then _
            VBA.Err.Raise VBA.vbObjectError + 2320, , _
            "Runtime preparation failed: " & VBA.CStr(m_context("Error"))
    End If
    If m_context("Phase") <> "Prepared" Then _
        VBA.Err.Raise VBA.vbObjectError + 2307, , "Runtime preparation was not acknowledged."
    If VBA.CLng(m_context("ActiveCalls")) <> 0 Then _
        VBA.Err.Raise VBA.vbObjectError + 2308, , "A target runtime call is still active."
    m_context("Phase") = "Importing"
    private_ApplyPlan m_target, m_plan

    If m_clearOnly Then
        m_target.Names("_RuntimeReloadBlocked").Delete
        m_target.Save
        m_context("Phase") = "Cleared"
        m_context("StopRequested") = True
        Set clearedTarget = m_target
        Application.Calculation = previousCalculation
        Application.ScreenUpdating = previousScreenUpdating
        Application.EnableEvents = previousEvents
        stateCaptured = False
        m_lastResult = "Cleared"
        Application.StatusBar = "VBA modules cleared. Backup: " & m_backupPath
        private_ClearOperation
        fn_ForgetContext clearedTarget
        Exit Sub
    End If

    ' The add-in retains the context across resets of the target project variables.
    m_context("Phase") = "Initializing"
    ' Excel resets project variables after the callback that modified code returns.
    ' Create the new runtime in the next callback, after the reset has completed.
    private_Schedule
    Application.Calculation = previousCalculation
    Application.ScreenUpdating = previousScreenUpdating
    Application.EnableEvents = previousEvents
    stateCaptured = False
    Exit Sub

InitializeRuntime:
    If Not VBA.CBool(Application.Run(private_Macro("ex_RuntimeLifecycle.fn_InitializeReloaded"), m_context, m_uiFolder)) Then _
        VBA.Err.Raise VBA.vbObjectError + 2321, , _
            "Runtime initialization failed: " & VBA.CStr(m_context("Error"))
    m_target.Names("_RuntimeReloadBlocked").Delete
    m_context("Generation") = VBA.CLng(m_context("Generation")) + 1
    m_context("Phase") = "Running"
    m_context("StopRequested") = False
    Application.Calculation = previousCalculation
    Application.ScreenUpdating = previousScreenUpdating
    Application.EnableEvents = previousEvents
    stateCaptured = False
    Application.StatusBar = "VBA reload completed. Backup: " & m_backupPath
    m_lastResult = "Completed"
    private_ClearOperation
    Exit Sub
EH:
    errorText = VBA.Err.Description
    m_lastResult = "Faulted"
    m_lastError = errorText
    If Not m_context Is Nothing Then
        m_context("StopRequested") = True
        m_context("Phase") = "Faulted"
        m_context("Error") = errorText
    End If
    ' After an error, the project remains blocked; execution does not resume automatically.
    On Error Resume Next
    If Not m_target Is Nothing And Not m_context Is Nothing Then _
        Application.Run private_Macro("ex_RuntimeLifecycle.fn_AttachContext"), m_context
    If stateCaptured Then
        Application.Calculation = previousCalculation
        Application.ScreenUpdating = previousScreenUpdating
        Application.EnableEvents = previousEvents
    End If
    On Error GoTo 0
    If m_showErrors Then VBA.MsgBox "VBA reload stopped: " & errorText & VBA.vbCrLf & _
        "The runtime remains blocked. Close without saving and recover the backup if necessary." & _
        VBA.vbCrLf & "Backup: " & m_backupPath, VBA.vbCritical, "Workbook updater"
    private_ClearOperation
End Sub

Public Sub fn_CancelPending(ByVal target As Workbook)
    If m_target Is Nothing Then Exit Sub
    If Not target Is m_target Then Exit Sub
    If m_scheduledAt = 0 Then VBA.Err.Raise VBA.vbObjectError + 2309, , "An update is executing."
    If m_context("Phase") <> "Requested" Then _
        VBA.Err.Raise VBA.vbObjectError + 2326, , _
            "Import has already started. Wait for initialization before closing or cancelling."
    Application.OnTime EarliestTime:=m_scheduledAt, Procedure:=private_Callback(), Schedule:=False
    m_scheduledAt = 0
    m_context("StopRequested") = False
    m_context("Phase") = "Running"
    m_lastResult = "Cancelled"
    private_ClearOperation
End Sub
' --------------------------------------
' } // namespace API
' --------------------------------------

Private Function private_Macro(ByVal methodName As String) As String
    private_Macro = "'" & Replace$(m_target.Name, "'", "''") & "'!" & methodName
End Function

Private Function private_Callback() As String
    private_Callback = "'" & Replace$(ThisWorkbook.Name, "'", "''") & "'!ex_WorkbookUpdater.fn_RunPending"
End Function

Private Sub private_Schedule()
    Dim scheduleAt As Date

    scheduleAt = VBA.Now + VBA.TimeSerial(0, 0, 1)
    Application.OnTime EarliestTime:=scheduleAt, Procedure:=private_Callback()
    m_scheduledAt = scheduleAt
End Sub

Private Function private_IsTargetOpen() As Boolean
    Dim book As Workbook

    For Each book In Application.Workbooks
        If book Is m_target Then
            private_IsTargetOpen = True
            Exit Function
        End If
    Next book
End Function

Private Sub private_ClearOperation()
    Set m_target = Nothing
    Set m_context = Nothing
    Set m_plan = Nothing
    m_scheduledAt = 0
    m_deadline = 0
    m_backupPath = VBA.vbNullString
    m_operationFolder = VBA.vbNullString
    m_uiFolder = VBA.vbNullString
    m_initialInstall = False
    m_clearOnly = False
End Sub

Private Function private_CreateOperationFolder(ByVal target As Workbook) As String
    Dim fileSystem As Object
    Dim folder As String
    Dim backupRoot As String
    Dim suffix As Long

    Set fileSystem = VBA.CreateObject("Scripting.FileSystemObject")
    backupRoot = target.Path & "\.backup"
    If Not fileSystem.FolderExists(backupRoot) Then fileSystem.CreateFolder backupRoot
    folder = backupRoot & "\reload-" & Format$(VBA.Now, "yyyymmdd-hhnnss")
    Do While fileSystem.FolderExists(folder)
        suffix = suffix + 1
        folder = backupRoot & "\reload-" & Format$(VBA.Now, "yyyymmdd-hhnnss") & "-" & VBA.CStr(suffix)
    Loop
    fileSystem.CreateFolder folder
    private_CreateOperationFolder = folder
End Function

Private Function private_IsEmptyProject(ByVal target As Workbook) As Boolean
    Dim component As Object

    For Each component In target.VBProject.VBComponents
        If component.Type <> 100 And component.Type <> 1 And component.Type <> 2 Then Exit Function
        If component.CodeModule.CountOfLines <> 0 Then Exit Function
    Next component
    private_IsEmptyProject = True
End Function

Private Function private_HasLifecycle(ByVal target As Workbook) As Boolean
    Dim component As Object

    For Each component In target.VBProject.VBComponents
        If VBA.StrComp(component.Name, "ex_RuntimeLifecycle", VBA.vbTextCompare) = 0 And component.Type = 1 Then
            private_HasLifecycle = True
            Exit Function
        End If
    Next component
End Function

Private Function private_HasBlockedMarker(ByVal target As Workbook) As Boolean
    Dim marker As Name
    Dim markerName As String

    For Each marker In target.Names
        markerName = marker.Name
        If VBA.InStrRev(markerName, "!") > 0 Then markerName = Mid$(markerName, VBA.InStrRev(markerName, "!") + 1)
        If VBA.StrComp(markerName, "_RuntimeReloadBlocked", VBA.vbTextCompare) = 0 Then
            private_HasBlockedMarker = True
            Exit Function
        End If
    Next marker
End Function

Private Sub private_RegisterContext(ByVal target As Workbook, ByVal context As Object)
    Dim entry As Object
    Dim registeredBook As Workbook
    Dim key As Variant

    If m_contexts Is Nothing Then Set m_contexts = VBA.CreateObject("Scripting.Dictionary")
    For Each key In m_contexts.Keys
        Set entry = m_contexts(key)
        Set registeredBook = entry("Workbook")
        If registeredBook Is target Then m_contexts.Remove key
    Next key
    Set entry = VBA.CreateObject("Scripting.Dictionary")
    entry.Add "Workbook", target
    entry.Add "Context", context
    If m_contexts.Exists(target.FullName) Then m_contexts.Remove target.FullName
    m_contexts.Add target.FullName, entry
End Sub

Private Function private_ReadPlan(ByVal target As Workbook, ByRef uiFolder As String) As Collection
    Dim root As String
    Dim files As New Collection
    Dim documents As New Collection
    Dim plan As New Collection
    Dim names As Object
    Dim fileSystem As Object
    Dim sourcePath As Variant
    Dim item As Object
    Dim code As String
    Dim fileName As String
    Dim sheetName As String

    If Not private_VbaReload_TryResolveConfiguredFolder(target, "ThisWorkbook::vbaPath", root) Then _
        VBA.Err.Raise VBA.vbObjectError + 2310, , "The VBA source folder is invalid."
    If Not private_VbaReload_TryResolveConfiguredFolder(target, "ThisWorkbook::uiPath", uiFolder) Then _
        VBA.Err.Raise VBA.vbObjectError + 2311, , "The UI source folder is invalid."
    If Not private_VbaReload_TryCollectConfiguredVbaFiles(root, target, files, documents) Then _
        VBA.Err.Raise VBA.vbObjectError + 2312, , "The source profile is invalid."
    Set fileSystem = VBA.CreateObject("Scripting.FileSystemObject")
    Set names = VBA.CreateObject("Scripting.Dictionary")
    names.CompareMode = VBA.vbTextCompare
    m_operationFolder = private_CreateOperationFolder(target)
    For Each sourcePath In documents
        files.Add sourcePath
    Next sourcePath
    For Each sourcePath In files
        Set item = VBA.CreateObject("Scripting.Dictionary")
        fileName = fileSystem.GetFileName(VBA.CStr(sourcePath))
        private_ValidateSourceEncoding VBA.CStr(sourcePath), LCase$(fileSystem.GetExtensionName(VBA.CStr(sourcePath))) <> "vba"
        code = private_VbaReload_ReadUtf8TextFile(VBA.CStr(sourcePath))
        If VBA.Len(Trim$(code)) = 0 Then _
        VBA.Err.Raise VBA.vbObjectError + 2313, , "Empty source: " & VBA.CStr(sourcePath)
        item.Add "Document", private_VbaReload_IsDocumentModuleSource(LCase$(fileName))
        If item("Document") Then
            If private_VbaReload_IsThisWorkbookModuleSource(fileName) Then
                item.Add "Name", target.CodeName
            Else
                sheetName = private_VbaReload_GetDocumentModuleSourceName(fileName, "ws_")
                item.Add "Name", target.Worksheets(sheetName).CodeName
            End If
            item.Add "Type", 100&
        Else
            item.Add "Name", private_VbaReload_GetComponentName(VBA.CStr(sourcePath), code)
            item.Add "Type", private_VbaReload_GetVbaComponentType(LCase$(VBA.CStr(sourcePath)), code)
            If item("Type") = 3 And LCase$(fileSystem.GetExtensionName(VBA.CStr(sourcePath))) = "vba" Then _
                VBA.Err.Raise VBA.vbObjectError + 2314, , _
            "UserForms require a native .frm/.frx export: " & VBA.CStr(sourcePath)
        End If
        If names.Exists(item("Name")) Then _
        VBA.Err.Raise VBA.vbObjectError + 2315, , "Duplicate component: " & item("Name")
        names.Add item("Name"), True
        item.Add "Code", private_VbaReload_PrepareSourceForVbe(code)
        item.Add "Native", LCase$(fileSystem.GetExtensionName(VBA.CStr(sourcePath))) <> "vba"
        item.Add "Path", VBA.CStr(sourcePath)
        If item("Native") And Not item("Document") Then
            item("Path") = m_operationFolder & "\" & fileName
            fileSystem.CopyFile VBA.CStr(sourcePath), VBA.CStr(item("Path")), False
            If item("Type") = 3 Then
                If fileSystem.FileExists(fileSystem.BuildPath(fileSystem.GetParentFolderName(VBA.CStr(sourcePath)), fileSystem.GetBaseName(VBA.CStr(sourcePath)) & ".frx")) Then
                    fileSystem.CopyFile fileSystem.BuildPath(fileSystem.GetParentFolderName(VBA.CStr(sourcePath)), fileSystem.GetBaseName(VBA.CStr(sourcePath)) & ".frx"), m_operationFolder & "\", False
                ElseIf InStr(1, code, ".frx", VBA.vbTextCompare) > 0 Then
                    VBA.Err.Raise VBA.vbObjectError + 2316, , "Missing .frx resource: " & VBA.CStr(sourcePath)
                End If
            End If
        End If
        plan.Add item
    Next sourcePath
    If Not names.Exists("ex_RuntimeLifecycle") Or Not names.Exists(target.CodeName) Then _
        VBA.Err.Raise VBA.vbObjectError + 2317, , _
            "The profile must include the runtime lifecycle and workbook event module."
    Set private_ReadPlan = plan
End Function

Private Sub private_ApplyPlan(ByVal target As Workbook, ByVal plan As Collection)
    Dim project As Object
    Dim component As Object
    Dim item As Object
    Dim index As Long

    Set project = target.VBProject
    For index = project.VBComponents.Count To 1 Step -1
        Set component = project.VBComponents(index)
        If component.Type = 100 Then
            If component.CodeModule.CountOfLines > 0 Then component.CodeModule.DeleteLines 1, component.CodeModule.CountOfLines
        ElseIf component.Type = 1 Or component.Type = 2 Or component.Type = 3 Then
            project.VBComponents.Remove component
        Else
            VBA.Err.Raise VBA.vbObjectError + 2318, , _
            "Unsupported component type: " & VBA.CStr(component.Type)
        End If
    Next index
    Set component = Nothing
    For Each item In plan
        If item("Document") Then
            Set component = project.VBComponents(VBA.CStr(item("Name")))
        ElseIf item("Native") Then
            Set component = project.VBComponents.Import(VBA.CStr(item("Path")))
        Else
            Set component = project.VBComponents.Add(VBA.CLng(item("Type")))
            component.Name = VBA.CStr(item("Name"))
        End If
        If Not item("Native") Or item("Document") Then component.CodeModule.AddFromString VBA.CStr(item("Code"))
        If component.Name <> item("Name") Or component.Type <> item("Type") Then _
            VBA.Err.Raise VBA.vbObjectError + 2319, , "Imported component identity mismatch: " & item("Name")
    Next item
End Sub

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

Private Function private_VbaReload_IsValidRelativeModulePattern( _
    ByVal relativePattern As String _
) As Boolean
    If VBA.Len(relativePattern) = 0 Or VBA.InStr(relativePattern, "..") > 0 Or _
       VBA.InStr(relativePattern, ":") > 0 Or VBA.Left$(relativePattern, 1) = "\" Then Exit Function
    If Not private_VbaReload_IsVbaModulePath(relativePattern) Then Exit Function
    private_VbaReload_IsValidRelativeModulePattern = True
End Function

Private Function private_VbaReload_IsVbaModulePath(ByVal filePath As String) As Boolean
    Dim lowerPath As String

    lowerPath = VBA.LCase$(filePath)
    private_VbaReload_IsVbaModulePath = _
        VBA.Right$(lowerPath, 4) = ".cls" Or _
        VBA.Right$(lowerPath, 4) = ".bas" Or _
        VBA.Right$(lowerPath, 4) = ".frm" Or _
        VBA.Right$(lowerPath, 4) = ".vba"
End Function

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
    If Not (VBA.Right$(lowerName, 4) = ".cls" Or _
            VBA.Right$(lowerName, 4) = ".bas" Or _
            VBA.Right$(lowerName, 4) = ".frm" Or _
            VBA.Right$(lowerName, 4) = ".vba") Then
        VBA.MsgBox "Configured file is not a VBA module: " & relativePath, _
            VBA.vbExclamation, "Reload VBA"
        Exit Function
    End If
    outSourcePath = normalizedPath
    private_VbaReload_TryResolveConfiguredSourcePath = True
End Function

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

Private Function private_VbaReload_GetVbaComponentType( _
    ByVal lowerPath As String, _
    ByVal sourceText As String _
) As Long
    Const VBEXT_CT_STD_MODULE As Long = 1
    Const VBEXT_CT_CLASS_MODULE As Long = 2
    Const VBEXT_CT_MS_FORM As Long = 3

    If VBA.Right$(lowerPath, 4) = ".cls" Or _
       VBA.Right$(lowerPath, 8) = ".cls.vba" Or _
       VBA.Right$(lowerPath, 13) = ".cls.utf8.vba" Then
        private_VbaReload_GetVbaComponentType = VBEXT_CT_CLASS_MODULE
    ElseIf VBA.Right$(lowerPath, 4) = ".frm" Or _
           VBA.Right$(lowerPath, 8) = ".frm.vba" Or _
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

Private Sub private_ValidateSourceEncoding(ByVal filePath As String, ByVal nativeSource As Boolean)
    Dim stream As Object
    Dim bytes As Variant
    Dim index As Long
    Dim lastIndex As Long
    Dim value As Long
    Dim continuation As Long
    Dim minimum As Long
    Dim codePoint As Long
    Dim offset As Long

    Set stream = VBA.CreateObject("ADODB.Stream")
    stream.Type = 1
    stream.Open
    stream.LoadFromFile filePath
    If stream.Size = 0 Then
        stream.Close
        Exit Sub
    End If
    bytes = stream.Read
    stream.Close
    lastIndex = UBound(bytes)
    index = LBound(bytes)
    Do While index <= lastIndex
        value = CLng(bytes(index))
        If nativeSource And value > 127 Then _
            Err.Raise vbObjectError + 2340, , "Native source must be ASCII; use UTF-8 .vba for Unicode code: " & filePath
        continuation = 0
        If value <= 127 Then
            codePoint = value
        ElseIf value >= 194 And value <= 223 Then
            continuation = 1: minimum = 128: codePoint = value And 31
        ElseIf value >= 224 And value <= 239 Then
            continuation = 2: minimum = 2048: codePoint = value And 15
        ElseIf value >= 240 And value <= 244 Then
            continuation = 3: minimum = 65536: codePoint = value And 7
        Else
            GoTo InvalidEncoding
        End If
        If index + continuation > lastIndex Then GoTo InvalidEncoding
        For offset = 1 To continuation
            value = CLng(bytes(index + offset))
            If value < 128 Or value > 191 Then GoTo InvalidEncoding
            codePoint = codePoint * 64 + (value And 63)
        Next offset
        If continuation > 0 Then
            If codePoint < minimum Or codePoint > 1114111 Then GoTo InvalidEncoding
            If codePoint >= 55296 And codePoint <= 57343 Then GoTo InvalidEncoding
            If codePoint = 65533 Then GoTo InvalidEncoding
        End If
        index = index + continuation + 1
    Loop
    Exit Sub
InvalidEncoding:
    Err.Raise vbObjectError + 2341, , "Invalid UTF-8 source at byte " & CStr(index) & ": " & filePath
End Sub

Private Function private_VbaReload_ReadUtf8TextFile(ByVal filePath As String) As String
    Const AD_TYPE_TEXT As Long = 2
    Const AD_READ_ALL As Long = -1

    Dim textStream As Object

    ' Read sources as UTF-8: the system ANSI code page can corrupt Unicode.
    private_ValidateSourceEncoding filePath, False
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

Private Function private_HasUnicode(ByVal text As String) As Boolean
    Dim index As Long
    Dim value As Long
    For index = 1 To Len(text)
        value = AscW(Mid$(text, index, 1))
        If value < 0 Or value > 127 Then
            private_HasUnicode = True
            Exit Function
        End If
    Next index
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

        If currentCharacter = "'" Or _
           (LCase$(Mid$(sourceLine, charIndex, 3)) = "rem" And _
            (InStr(" " & vbTab & ":", Mid$(" " & sourceLine, charIndex, 1)) > 0) And _
            (charIndex + 3 > Len(sourceLine) Or InStr(" " & vbTab, Mid$(sourceLine, charIndex + 3, 1)) > 0)) Then
            ' An apostrophe outside a string literal starts a comment.
            resultText = resultText & VBA.Mid$(sourceLine, charIndex)
            Exit Do
        End If

        If currentCharacter <> """" Then
            If AscW(currentCharacter) < 0 Or AscW(currentCharacter) > 127 Then _
                Err.Raise vbObjectError + 2342, , "Non-ASCII VBA identifier is not portable. Use ASCII identifiers."
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
                Err.Raise vbObjectError + 2343, , "Unterminated VBA string literal."
                resultText = resultText & VBA.Mid$(sourceLine, literalStartIndex)
                Exit Do
            End If

            If InStr(1, sourceLine, "Const ", vbTextCompare) > 0 And private_HasUnicode(literalText) Then _
                Err.Raise vbObjectError + 2344, , "Unicode Const literals require runtime initialization; ChrW is not a constant expression."
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