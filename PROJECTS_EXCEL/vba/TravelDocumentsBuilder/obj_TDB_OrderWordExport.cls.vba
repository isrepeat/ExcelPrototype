VERSION 1.0 CLASS
BEGIN
  MultiUse = -1
END
Attribute VB_Name = "obj_TDB_OrderWordExport"
Option Explicit

Private m_isInitialized As Boolean
Private m_isDisposed As Boolean

Private m_vacationGenerator As obj_TDB_VacationGenerator
Private m_fileSystem As Object
Private m_namePattern As Object
Private m_legacyNamePattern As Object
Private m_config As Object

' //
' // Lifecycle
' //
Private Sub Class_Initialize()
End Sub

Private Sub Class_Terminate()
    Me.Dispose
End Sub

' //
' // API
' //
Public Function Initialize( _
    ByVal bindingContext As obj_UiBindingContext, _
    ByVal lookupService As obj_LookupService _
) As Boolean
    Dim key As Variant

    If m_isInitialized Or m_isDisposed Then Exit Function
    Set m_vacationGenerator = New obj_TDB_VacationGenerator
    If Not m_vacationGenerator.Initialize(bindingContext, lookupService) Then Exit Function
    Set m_fileSystem = VBA.CreateObject("Scripting.FileSystemObject")
    Set m_config = VBA.CreateObject("Scripting.Dictionary")
    For Each key In VBA.Split("folder,maxPath,filename.label,filename.pattern,filename.labelMaxChars,filename.legacyPattern," & _
            "text.Path,text.Stale,text.Identity,text.Loading,text.Progress,text.Success,text.Unmanaged,text.Rollback,text.Cleanup,text.Failed", ",")
        m_config.Add VBA.CStr(key), ex_TravelDocumentsBuilder.fn_Config("word." & VBA.CStr(key))
    Next key
    Set m_namePattern = VBA.CreateObject("VBScript.RegExp")
    m_namePattern.Pattern = private_Config("filename.pattern")
    m_namePattern.IgnoreCase = True
    Set m_legacyNamePattern = VBA.CreateObject("VBScript.RegExp")
    m_legacyNamePattern.Pattern = private_Config("filename.legacyPattern")
    m_legacyNamePattern.IgnoreCase = True
    m_isInitialized = True
    Initialize = True
End Function

Public Sub Dispose()
    If m_isDisposed Then Exit Sub
    m_isDisposed = True
    m_isInitialized = False
    If Not m_vacationGenerator Is Nothing Then m_vacationGenerator.Dispose
    Set m_vacationGenerator = Nothing
    Set m_fileSystem = Nothing
    Set m_namePattern = Nothing
    Set m_legacyNamePattern = Nothing
    Set m_config = Nothing
End Sub

Public Function TryExport( _
    ByVal orderNumber As String, _
    ByVal orderDate As Date, _
    ByVal templatePath As String, _
    ByVal outputRoot As String, _
    ByRef diagnostic As String _
) As Boolean
    Dim tickets As ListObject
    Dim revisions As ListObject
    Dim transaction As Object
    Dim groups As Collection
    Dim group As Collection
    Dim root As Object
    Dim wanted As Object
    Dim staged As Object
    Dim owned As New Collection
    Dim backups As Object
    Dim published As New Collection
    Dim file As Object
    Dim path As Variant
    Dim folder As String
    Dim stage As String
    Dim stageCreated As Boolean
    Dim backupFolder As String
    Dim frozenTemplate As String
    Dim templateStamp As String
    Dim folderStamp As String
    Dim name As String
    Dim identifier As String
    Dim target As ListRow
    Dim metadata As Object
    Dim progress As Long
    Dim unmanaged As Long
    Dim historyChanged As Boolean
    Dim committed As Boolean
    Dim busyEntered As Boolean
    Dim guarded As Boolean
    Dim previousEvents As Boolean
    Dim previousUpdating As Boolean
    Dim failure As String
    Dim rollback As String
    Dim text As String
    Dim activePath As String
    Dim protection As obj_TDB_HistoryProtection
    Dim protectionDiagnostic As String
    Dim errorNumber As Long
    Dim errorSource As String
    Dim errorDescription As String
    Dim totalStartedAt As Double
    Dim phaseStartedAt As Double
    Dim diagnosticBatchEntered As Boolean
    Dim renderedCount As Long
    Dim previousWindow As Window
    Dim previousSheet As Object
    Dim signatory As Object

    If Not m_isInitialized Or m_isDisposed Then Exit Function
    totalStartedAt = VBA.Timer
    phaseStartedAt = totalStartedAt
    On Error GoTo Failed
    previousEvents = Application.EnableEvents
    previousUpdating = Application.ScreenUpdating
    Set previousWindow = Application.ActiveWindow
    Set previousSheet = Application.ActiveSheet
    guarded = True
    Application.EnableEvents = False
    Application.ScreenUpdating = False
    ex_Core.fn_Diagnostic_BeginBatch
    diagnosticBatchEntered = True
    ex_Core.fn_Diagnostic_WriteLog "TDB_WORD_EXPORT_STARTED | Order=" & orderNumber & " | Date=" & VBA.Format$(orderDate, "yyyy-mm-dd")
    If VBA.Len(orderNumber) = 0 Or orderNumber Like "*[!0-9]*" Or orderDate = 0 Then ex_TDB_History.fn_Fail "OrderRequired"
    templatePath = private_Path(templatePath)
    activePath = templatePath
    outputRoot = private_Path(outputRoot)
    If Not m_fileSystem.FileExists(templatePath) Or Not m_fileSystem.FolderExists(outputRoot) Then private_Fail "Path"
    ex_TDB_History.fn_Open tickets, revisions
    Set transaction = ex_TDB_History.fn_Begin(tickets, revisions, VBA.vbNullString, VBA.vbNullString)
    Set groups = ex_TDB_History.fn_OrderGroups(tickets, revisions, orderNumber, orderDate)
    Set wanted = VBA.CreateObject("Scripting.Dictionary")
    wanted.CompareMode = VBA.vbTextCompare
    Set staged = VBA.CreateObject("Scripting.Dictionary")
    staged.CompareMode = VBA.vbTextCompare
    Set backups = VBA.CreateObject("Scripting.Dictionary")
    backups.CompareMode = VBA.vbTextCompare

    ' Only the selected order's dedicated folder belongs to this synchronization.
    name = private_Config("folder")
    If VBA.InStr(name, "{date}") = 0 Or VBA.InStr(name, "{number}") = 0 Then private_Fail "Path"
    name = VBA.Replace$(name, "{date}", VBA.Format$(orderDate, "yyyymmdd"))
    name = VBA.Replace$(name, "{number}", orderNumber)
    If VBA.Len(name) = 0 Or name <> m_fileSystem.GetFileName(name) Or name = "." Or name = ".." Then private_Fail "Path"
    folder = m_fileSystem.BuildPath(outputRoot, name)
    If Not m_fileSystem.FolderExists(folder) Then m_fileSystem.CreateFolder folder
    stage = m_fileSystem.BuildPath(folder, "._TDB-stage-" & ex_TDB_History.fn_NewId())
    frozenTemplate = m_fileSystem.BuildPath(stage, "template.docx")
    If VBA.Len(frozenTemplate) > VBA.CLng(private_Config("maxPath")) Then private_Fail "Path"
    If m_fileSystem.FolderExists(stage) Then private_Fail "Path"
    m_fileSystem.CreateFolder stage
    stageCreated = True
    backupFolder = m_fileSystem.BuildPath(stage, "backup")
    m_fileSystem.CreateFolder backupFolder
    templateStamp = private_FileStamp(templatePath)
    m_fileSystem.CopyFile templatePath, frozenTemplate, False
    folderStamp = private_FolderStamp(folder)
    For Each file In m_fileSystem.GetFolder(folder).Files
        If VBA.LCase$(m_fileSystem.GetExtensionName(file.Name)) = "docx" Then
            If m_namePattern.Test(file.Name) Or m_legacyNamePattern.Test(file.Name) Then
                If VBA.StrComp(file.Path, templatePath, VBA.vbTextCompare) = 0 Then private_Fail "Path"
                owned.Add file.Path
            Else
                unmanaged = unmanaged + 1
            End If
        End If
    Next file
    For Each group In groups
        Set root = group(1)
        If root("Status") = "Active" Then
            identifier = root("RowId")
            name = private_FileName(root)
            If VBA.Len(m_fileSystem.BuildPath(folder, name)) > VBA.CLng(private_Config("maxPath")) Then private_Fail "Path"
            wanted.Add identifier, name
            staged.Add identifier, VBA.CStr(wanted.Count) & ".docx"
            If VBA.Len(m_fileSystem.BuildPath(stage, staged(identifier))) > VBA.CLng(private_Config("maxPath")) Then private_Fail "Path"
        End If
    Next group

    ' Render a complete frozen set before replacing or removing any existing document.
    If wanted.Count > 0 Then Set signatory = ex_TDB_Signatory.fn_Resolve(orderNumber, orderDate)
    ex_Core.fn_Diagnostic_WritePerf "TDB.Word.Export.Prepare", phaseStartedAt
    phaseStartedAt = VBA.Timer
    busyEntered = ex_UiBusy.fn_Begin(private_Config("text.Loading"))
    For Each group In groups
        Set root = group(1)
        If root("Status") = "Active" Then
            progress = progress + 1
            activePath = m_fileSystem.BuildPath(stage, staged(root("RowId")))
            text = VBA.Replace$(private_Config("text.Progress"), "{step}", VBA.CStr(progress))
            text = VBA.Replace$(text, "{total}", VBA.CStr(wanted.Count))
            ex_Core.fn_StatusBar_Show text
            If Not m_vacationGenerator.TryRender(root, group, frozenTemplate, activePath, signatory, diagnostic) Then
                VBA.Err.Raise 5, , diagnostic
            End If
            renderedCount = renderedCount + 1
        End If
    Next group
    ex_Core.fn_Diagnostic_WritePerf "TDB.Word.Export.RenderAll", phaseStartedAt
    phaseStartedAt = VBA.Timer
    If private_FileStamp(templatePath) <> templateStamp Or private_FolderStamp(folder) <> folderStamp Then private_Fail "Stale"
    ex_TDB_History.fn_AssertCurrent tickets, revisions, transaction
    ex_Core.fn_Diagnostic_WritePerf "TDB.Word.Export.Verify", phaseStartedAt
    phaseStartedAt = VBA.Timer

    ' Preserve the previous set until every file and latest-document metadata is committed.
    For Each path In owned
        activePath = VBA.CStr(path)
        name = m_fileSystem.BuildPath(backupFolder, VBA.CStr(backups.Count + 1) & ".docx")
        If VBA.Len(name) > VBA.CLng(private_Config("maxPath")) Then private_Fail "Path"
        m_fileSystem.MoveFile VBA.CStr(path), name
        backups.Add VBA.CStr(path), name
    Next path
    For Each path In wanted.Keys
        name = m_fileSystem.BuildPath(folder, wanted(path))
        activePath = name
        m_fileSystem.MoveFile m_fileSystem.BuildPath(stage, staged(path)), name
        published.Add name
    Next path
    ex_Core.fn_Diagnostic_WritePerf "TDB.Word.Export.PublishFiles", phaseStartedAt
    phaseStartedAt = VBA.Timer
    ex_TDB_History.fn_AssertCurrent tickets, revisions, transaction
    Set protection = New obj_TDB_HistoryProtection
    If Not protection.Initialize(tickets, revisions) Then ex_TDB_History.fn_Fail "ReadOnly"
    protection.PrepareMetadataWrite
    historyChanged = True
    For Each group In groups
        Set root = group(1)
        activePath = tickets.Parent.Parent.FullName & " / " & root("RowId")
        Set target = ex_TDB_History.fn_Find(tickets, VBA.CStr(root("RowId")))
        If target Is Nothing Then ex_TDB_History.fn_Fail "NotFound"
        Set metadata = VBA.CreateObject("Scripting.Dictionary")
        metadata.Add "DocumentPath", VBA.vbNullString
        metadata.Add "DocumentRevision", root("Revision")
        metadata.Add "DocumentGeneratedAt", VBA.vbNullString
        If wanted.Exists(root("RowId")) Then
            metadata("DocumentPath") = m_fileSystem.BuildPath(folder, wanted(root("RowId")))
            metadata("DocumentGeneratedAt") = VBA.Format$(VBA.Now, "yyyy-mm-dd hh:nn:ss")
        End If
        ex_TDB_History.fn_Write tickets, target, metadata
    Next group
    If Not protection.TryProtect(protectionDiagnostic) Then VBA.Err.Raise 5, , protectionDiagnostic
    ex_Core.fn_Diagnostic_WritePerf "TDB.Word.Export.MetadataAndProtection", phaseStartedAt
    committed = True
    TryExport = True
    diagnostic = VBA.Replace$(private_Config("text.Success"), "{count}", VBA.CStr(wanted.Count))
    diagnostic = VBA.Replace$(diagnostic, "{oldCount}", VBA.CStr(owned.Count))
    diagnostic = VBA.Replace$(diagnostic, "{folder}", folder)
    If unmanaged > 0 Then diagnostic = diagnostic & VBA.vbLf & VBA.Replace$(private_Config("text.Unmanaged"), "{count}", VBA.CStr(unmanaged))
    GoTo Cleanup
Failed:
    errorNumber = VBA.Err.Number
    errorSource = VBA.Err.Source
    errorDescription = VBA.Err.Description
    ex_Core.fn_Diagnostic_WriteError "TDB.Word.Export | Path=" & activePath, errorNumber, errorSource, errorDescription
    failure = VBA.Replace$(private_Config("text.Failed"), "{error}", errorDescription)
    failure = VBA.Replace$(failure, "{path}", activePath)
    On Error Resume Next
    If Not committed Then
        For Each path In published
            VBA.Err.Clear
            If m_fileSystem.FileExists(VBA.CStr(path)) Then m_fileSystem.DeleteFile VBA.CStr(path), True
            If VBA.Err.Number <> 0 Then rollback = rollback & VBA.Err.Description & VBA.vbLf
        Next path
        If Not backups Is Nothing Then
            For Each path In backups.Keys
                VBA.Err.Clear
                m_fileSystem.MoveFile backups(path), VBA.CStr(path)
                If VBA.Err.Number <> 0 Then rollback = rollback & VBA.Err.Description & VBA.vbLf
            Next path
        End If
        If historyChanged Then
            VBA.Err.Clear
            If Not protection Is Nothing Then protection.PrepareMetadataWrite
            If VBA.Err.Number <> 0 Then rollback = rollback & VBA.Err.Description & VBA.vbLf
            VBA.Err.Clear
            ex_TDB_History.fn_Rollback tickets, revisions, transaction
            If VBA.Err.Number <> 0 Then rollback = rollback & VBA.Err.Description & VBA.vbLf
        End If
    End If
    On Error GoTo 0
    diagnostic = failure
    If VBA.Len(rollback) > 0 Then diagnostic = diagnostic & VBA.vbLf & private_Config("text.Rollback") & VBA.vbLf & rollback & stage
Cleanup:
    On Error Resume Next
    phaseStartedAt = VBA.Timer
    ex_TDB_WordSession.fn_EndExport
    If Not m_vacationGenerator Is Nothing Then m_vacationGenerator.Dispose
    ex_Core.fn_Diagnostic_WritePerf "TDB.Word.Export.RendererCleanup", phaseStartedAt
    phaseStartedAt = VBA.Timer
    If Not protection Is Nothing Then
        If Not protection.TryProtect(protectionDiagnostic) Then
            diagnostic = diagnostic & VBA.vbLf & protectionDiagnostic
            TryExport = False
        End If
        protection.Dispose
    End If
    ' A failed restoration retains its backup directory for explicit recovery.
    If stageCreated And VBA.Len(rollback) = 0 Then
        VBA.Err.Clear
        If VBA.StrComp(m_fileSystem.GetParentFolderName(m_fileSystem.GetAbsolutePathName(stage)), m_fileSystem.GetAbsolutePathName(folder), VBA.vbTextCompare) = 0 Then
            If VBA.Left$(m_fileSystem.GetFileName(stage), VBA.Len("._TDB-stage-T-")) = "._TDB-stage-T-" Then
                If m_fileSystem.FolderExists(stage) Then m_fileSystem.DeleteFolder stage, True
            End If
        End If
        If VBA.Err.Number <> 0 Then diagnostic = diagnostic & VBA.vbLf & private_Config("text.Cleanup") & " " & stage
    End If
    If guarded Then
        ' Restore the original Excel view before any screen refresh, as in PEB Movement export.
        If Not previousWindow Is Nothing Then
            If Not Application.ActiveWindow Is previousWindow Then previousWindow.Activate
        End If
        If Not previousSheet Is Nothing Then
            If Not Application.ActiveSheet Is previousSheet Then previousSheet.Activate
        End If
        Application.EnableEvents = previousEvents
        Application.ScreenUpdating = previousUpdating
    End If
    If busyEntered Then ex_UiBusy.fn_End
    ex_Core.fn_Diagnostic_WritePerf "TDB.Word.Export.Cleanup", phaseStartedAt
    ex_Core.fn_Diagnostic_WriteLog "TDB_WORD_EXPORT_FINISHED | Success=" & VBA.CStr(TryExport) & " | Rendered=" & VBA.CStr(renderedCount)
    ex_Core.fn_Diagnostic_WritePerf "TDB.Word.Export.Total", totalStartedAt
    If diagnosticBatchEntered Then ex_Core.fn_Diagnostic_EndBatch
    On Error GoTo 0
End Function

' //
' // Private
' //
Private Function private_Config(ByVal key As String) As String
    If Not m_config.Exists(key) Then VBA.Err.Raise 5, , "Required configuration key not found: TravelDocumentsBuilder::word." & key
    private_Config = m_config(key)
End Function

Private Sub private_Fail(ByVal key As String)
    VBA.Err.Raise VBA.vbObjectError + 2460, "TicketsWordExport", private_Config("text." & key)
End Sub

Private Function private_Path(ByVal value As String) As String
    If VBA.Len(VBA.Trim$(value)) = 0 Then private_Fail "Path"
    If VBA.Left$(value, 1) = "\" And VBA.Left$(value, 2) <> "\\" Then private_Fail "Path"
    If m_fileSystem.GetDriveName(value) = VBA.vbNullString And VBA.Left$(value, 2) <> "\\" Then value = m_fileSystem.BuildPath(ThisWorkbook.Path, value)
    private_Path = m_fileSystem.GetAbsolutePathName(value)
End Function

Private Function private_FileName(ByVal root As Object) As String
    Dim label As String
    Dim name As String
    Dim character As Variant
    Dim index As Long
    Dim matches As Object

    label = VBA.Replace$(private_Config("filename.label"), "{TicketNo}", VBA.CStr(root(ex_TDB_History.fn_Config("ticketField"))))
    label = VBA.Replace$(label, "{FIO}", VBA.CStr(root(ex_TDB_History.fn_Config("personNameField"))))
    For Each character In VBA.Array("\", "/", ":", "*", "?", """", "<", ">", "|")
        label = VBA.Replace$(label, VBA.CStr(character), "_")
    Next character
    For index = 0 To 31
        label = VBA.Replace$(label, VBA.Chr$(index), "_")
    Next index
    label = VBA.Left$(label, VBA.CLng(private_Config("filename.labelMaxChars")))
    name = "TDB_" & label & "__" & root("RowId") & ".docx"
    If Not m_namePattern.Test(name) Then private_Fail "Identity"
    Set matches = m_namePattern.Execute(name)
    If VBA.StrComp(matches(0).SubMatches(0), VBA.CStr(root("RowId")), VBA.vbTextCompare) <> 0 Then private_Fail "Identity"
    private_FileName = name
End Function

Private Function private_FileStamp(ByVal path As String) As String
    Dim file As Object

    Set file = m_fileSystem.GetFile(path)
    private_FileStamp = VBA.CStr(file.Size) & "/" & VBA.CStr(file.DateLastModified)
End Function

Private Function private_FolderStamp(ByVal path As String) As String
    Dim names As New Collection
    Dim file As Object
    Dim name As Variant
    Dim index As Long
    Dim inserted As Boolean
    Dim result As String

    For Each file In m_fileSystem.GetFolder(path).Files
        inserted = False
        For index = 1 To names.Count
            If VBA.StrComp(file.Name, VBA.CStr(names(index)), VBA.vbBinaryCompare) < 0 Then
                names.Add file.Name, Before:=index
                inserted = True
                Exit For
            End If
        Next index
        If Not inserted Then names.Add file.Name
    Next file
    For Each name In names
        result = result & VBA.Len(VBA.CStr(name)) & ":" & name & "/" & private_FileStamp(m_fileSystem.BuildPath(path, VBA.CStr(name))) & "|"
    Next name
    private_FolderStamp = result
End Function