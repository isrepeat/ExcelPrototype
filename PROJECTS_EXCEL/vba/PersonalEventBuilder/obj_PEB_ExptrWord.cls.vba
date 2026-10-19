VERSION 1.0 CLASS
BEGIN
  MultiUse = -1
END
Attribute VB_Name = "obj_PEB_ExptrWord"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private m_isInitialized As Boolean
Private m_isDisposed As Boolean

Private m_templatePath As String
Private m_documentMode As String

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
Public Function Initialize() As Boolean
    If m_isInitialized Or m_isDisposed Then
        Exit Function
    End If
    m_documentMode = ex_PEB_Export.fn_Config("export.Word.documentMode")
    If m_documentMode <> "blank" And m_documentMode <> "template" Then
        Err.Raise 5, , "Word documentMode must be blank or template."
    End If
    If m_documentMode = "template" Then
        m_templatePath = ex_PEB_Export.fn_Path(ex_PEB_Export.fn_Config("export.Word.documentTemplate"))
    End If
    m_isInitialized = True
    Initialize = True
End Function

Public Sub Dispose()
    If m_isDisposed Then
        Exit Sub
    End If
    m_isDisposed = True
    m_isInitialized = False
End Sub

Public Function TryExport( _
    ByVal values As Object, _
    ByRef diagnostic As String _
) As Boolean
    TryExport = private_Export(values, diagnostic, False, "INSERT", VBA.vbNullString)
End Function

Public Function TryGenerateHistory( _
    ByVal events As Collection, _
    ByRef output As String, _
    ByRef diagnostic As String _
) As Boolean
    Dim service As obj_PEB_ExportService
    Dim templates As obj_PEB_WordTemplates
    Dim anchors As obj_PEB_WordAnchors
    Dim sections As Object
    Dim sectionTemplates As Object
    Dim rendered As Object
    Dim fileSystem As Object
    Dim wordApp As Object
    Dim document As Object
    Dim item As Object
    Dim first As Object
    Dim context As Object
    Dim definition As Object
    Dim group As Collection
    Dim eventId As Variant
    Dim text As String
    Dim body As String
    Dim temporary As String
    Dim originalStamp As String
    Dim stamp As String
    Dim published As Boolean
    Dim existing As Boolean

    On Error GoTo EH
    If Not m_isInitialized Or m_isDisposed Then Err.Raise 5, , ex_PEB_Export.fn_Config("text.HistoryError036")
    If ex_PEB_Export.fn_Config("export.Word.enabled") <> "true" Then Err.Raise 5, , ex_PEB_Export.fn_Config("text.ExportWordUnavailable")
    If m_documentMode <> "template" Then Err.Raise 5, , ex_PEB_Export.fn_Config("text.HistoryError037")
    Set service = New obj_PEB_ExportService
    If Not service.Initialize() Then Err.Raise 5, , ex_PEB_Export.fn_Config("text.HistoryError038")
    Set templates = New obj_PEB_WordTemplates
    If Not templates.Initialize(diagnostic) Then GoTo CleanupHistory
    Set sections = VBA.CreateObject("Scripting.Dictionary")
    Set sectionTemplates = VBA.CreateObject("Scripting.Dictionary")
    Set rendered = VBA.CreateObject("Scripting.Dictionary")
    For Each item In events
        If Not service.TryPrepareHistory(item, diagnostic) Then
            diagnostic = "History " & VBA.CStr(item("RowId")) & ": " & diagnostic
            GoTo CleanupHistory
        End If
        If item("WordPolicy") <> "Omit" Then
            If first Is Nothing Then
                Set first = item
                output = item("WordOutput")
            ElseIf item("WordOutput") <> output Or item("Signatory") <> first("Signatory") Then
                Err.Raise 5, , ex_PEB_Export.fn_Config("text.HistoryError039")
            End If
            eventId = item("EventId")
            If Not sections.Exists(eventId) Then
                Set group = New Collection
                sections.Add eventId, group
                sectionTemplates.Add eventId, item("TemplateName")
            End If
            sections(eventId).Add item
        End If
    Next item
    If first Is Nothing Then Err.Raise 5, , ex_PEB_Export.fn_Config("text.HistoryNoWordEvents")
    ' Prepare every section before opening Word; ordinary events accumulate as well.
    For Each eventId In sections.Keys
        Set definition = templates.Definition(sectionTemplates(eventId))
        If definition.selectNodes("p:group").Length > 0 Then
            body = ex_PEB_WordGroups.fn_Render(definition, sections(eventId), templates.SharedTemplates)
        Else
            body = VBA.vbNullString
            For Each item In sections(eventId)
                If Not templates.TryRender(sectionTemplates(eventId), item, text, diagnostic) Then GoTo CleanupHistory
                If VBA.Len(body) > 0 Then body = body & VBA.vbCrLf & VBA.vbCrLf
                body = body & text
            Next item
        End If
        rendered.Add eventId, body
    Next eventId
    Set fileSystem = VBA.CreateObject("Scripting.FileSystemObject")
    If Not fileSystem.FileExists(m_templatePath) Then Err.Raise 5, , VBA.Replace$(ex_PEB_Export.fn_Config("text.HistoryError040"), "{value}", VBA.CStr(m_templatePath))
    If VBA.StrComp(output, m_templatePath, VBA.vbTextCompare) = 0 Then Err.Raise 5, , ex_PEB_Export.fn_Config("text.HistoryError041")
    existing = fileSystem.FileExists(output)
    If existing Then originalStamp = VBA.CStr(fileSystem.GetFile(output).DateLastModified) & ":" & VBA.CStr(fileSystem.GetFile(output).Size)
    temporary = fileSystem.BuildPath(fileSystem.GetParentFolderName(output), fileSystem.GetTempName & ".docx")
    Set wordApp = ex_PEB_WordState.fn_AcquireWord()
    wordApp.Visible = False
    wordApp.DisplayAlerts = 0
    wordApp.AutomationSecurity = 3
    Set document = wordApp.Documents.Add(m_templatePath, False, 0, False)
    private_ReplaceDocumentFields document, first
    ' Each document starts from the empty template; populate only rendered sections.
    For Each eventId In rendered.Keys
        If VBA.Len(rendered(eventId)) > 0 Then
            Set context = VBA.CreateObject("Scripting.Dictionary")
            context.Add "EventId", eventId
            context.Add "WordText", rendered(eventId)
            Set anchors = New obj_PEB_WordAnchors
            If Not anchors.Initialize(document, diagnostic) Then GoTo CleanupHistory
            If Not anchors.TryInsert(document, context, diagnostic) Then GoTo CleanupHistory
            If Not anchors.TryVerify(document, context, diagnostic) Then GoTo CleanupHistory
            anchors.HideMarkers document, context
            anchors.Dispose
            Set anchors = Nothing
        End If
    Next eventId
    document.SaveAs2 temporary, 16
    document.Close 0
    Set document = Nothing
    If existing Then
        If Not fileSystem.FileExists(output) Then Err.Raise 5, , ex_PEB_Export.fn_Config("text.HistoryError042")
        stamp = VBA.CStr(fileSystem.GetFile(output).DateLastModified) & ":" & VBA.CStr(fileSystem.GetFile(output).Size)
        If stamp <> originalStamp Then Err.Raise 5, , ex_PEB_Export.fn_Config("text.HistoryError043")
        ex_PEB_WordState.fn_Publish temporary, output
    Else
        If fileSystem.FileExists(output) Then Err.Raise 5, , ex_PEB_Export.fn_Config("text.HistoryError044")
        fileSystem.MoveFile temporary, output
    End If
    published = True
    TryGenerateHistory = True
CleanupHistory:
    On Error Resume Next
    If Not document Is Nothing Then document.Close 0
    If Not anchors Is Nothing Then anchors.Dispose
    If Not templates Is Nothing Then templates.Dispose
    If Not service Is Nothing Then service.Dispose
    If Not published And VBA.Len(temporary) > 0 Then
        If fileSystem.FileExists(temporary) Then fileSystem.DeleteFile temporary
    End If
    On Error GoTo 0
    Exit Function
EH:
    diagnostic = VBA.Err.Description
    Resume CleanupHistory
End Function

Public Function TryUpdate( _
    ByVal values As Object, _
    ByVal eventUid As String, _
    ByRef diagnostic As String _
) As Boolean
    TryUpdate = private_Export(values, diagnostic, False, "UPDATE", eventUid)
End Function

Public Function TryDelete( _
    ByVal values As Object, _
    ByVal eventUid As String, _
    ByRef diagnostic As String _
) As Boolean
    TryDelete = private_Export(values, diagnostic, False, "DELETE", eventUid)
End Function

Public Function TryExportExistingTest( _
    ByVal values As Object, _
    ByRef diagnostic As String _
) As Boolean
    TryExportExistingTest = private_Export(values, diagnostic, True, "INSERT", VBA.vbNullString)
End Function

Public Function TryReadEvents( _
    ByVal values As Object, _
    ByRef events As Collection, _
    ByRef diagnostic As String _
) As Boolean
    Dim wordApp As Object
    Dim document As Object
    Dim output As String

    On Error GoTo EH
    Set events = Nothing
    If Not m_isInitialized Or m_isDisposed Then
        Err.Raise 5, , "Word exporter is not initialized."
    End If
    output = ex_PEB_Export.fn_Text(ex_PEB_Export.fn_Required(values, "WordOutput"))
    Set wordApp = ex_PEB_WordState.fn_AcquireWord()
    wordApp.Visible = False
    wordApp.DisplayAlerts = 0
    wordApp.AutomationSecurity = 3
    Set document = wordApp.Documents.Open(output, False, True, False)
    Set events = ex_PEB_WordDocument.fn_Events(document, values)
    TryReadEvents = True
Cleanup:
    On Error Resume Next
    If Not document Is Nothing Then
        document.Close 0
    End If
    Set wordApp = Nothing
    On Error GoTo 0
    Exit Function
EH:
    diagnostic = VBA.Err.Description
    Resume Cleanup
End Function

' //
' // Private
' //
Private Function private_Export( _
    ByVal values As Object, _
    ByRef diagnostic As String, _
    ByVal existingTest As Boolean, _
    ByVal operation As String, _
    ByVal eventUid As String _
) As Boolean
    Dim wordApp As Object
    Dim document As Object
    Dim fileSystem As Object
    Dim output As String
    Dim temporary As String
    Dim published As Boolean
    Dim cleanupError As String
    Dim enabled As String
    Dim grouped As Boolean
    Dim existing As Boolean
    Dim identifier As String
    Dim originalStamp As String
    Dim resultCode As String

    On Error GoTo EH
    diagnostic = VBA.vbNullString
    values("WordExportAction") = VBA.vbNullString
    If Not m_isInitialized Or m_isDisposed Then
        Err.Raise 5, , "Word exporter is not initialized."
    End If
    If operation <> "DELETE" Then
        If ex_PEB_Export.fn_Text(ex_PEB_Export.fn_Required(values, "WordPolicy")) = "Omit" Then
            If operation <> "INSERT" Then
                Err.Raise 5, , "Word update requires an exported event route."
            End If
            values("WordExportAction") = "SKIP"
            private_Export = True
            Exit Function
        End If
    End If
    enabled = ex_PEB_Export.fn_Config("export.Word.enabled")
    If existingTest Then
        If ex_PEB_Export.fn_Config("environment.name") <> "TEST" Then
            Err.Raise 5, , "Existing-document test export requires TEST."
        End If
        enabled = "true"
    End If
    If enabled = "false" Then
        Err.Raise 5, , ex_PEB_Export.fn_Config("text.ExportWordUnavailable")
    End If
    If enabled <> "true" Then
        Err.Raise 5, , "Word enabled must be true or false."
    End If
    output = ex_PEB_Export.fn_Text(ex_PEB_Export.fn_Required(values, "WordOutput"))
    Set fileSystem = VBA.CreateObject("Scripting.FileSystemObject")
    existing = fileSystem.FileExists(output)
    If operation <> "INSERT" And Not existing Then
        Err.Raise 5, , ex_PEB_Export.fn_Config("text.WordStoreNotFound")
    End If
    If existingTest Then
        grouped = VBA.CBool(ex_PEB_Export.fn_Required(values, "WordGrouped"))
        If Not existing Or Not grouped Then
            Err.Raise 5, , "Test export requires an existing grouped DOCX."
        End If
        If VBA.StrComp(fileSystem.GetParentFolderName(fileSystem.GetAbsolutePathName(output)), _
                fileSystem.GetSpecialFolder(2).Path, VBA.vbTextCompare) <> 0 Or _
                VBA.Left$(fileSystem.GetFileName(output), 13) <> "PEB_Grouping_" Then
            Err.Raise 5, , "Test export is restricted to owned temporary grouping documents."
        End If
    End If
    If m_documentMode = "template" And Not existing Then
        If Not fileSystem.FileExists(m_templatePath) Then
            Err.Raise 5, , "Word document template not found: " & m_templatePath
        End If
    End If
    If VBA.LCase$(fileSystem.GetExtensionName(output)) <> "docx" Then
        Err.Raise 5, , "Word output must have a .docx extension."
    End If
    If Not fileSystem.FolderExists(fileSystem.GetParentFolderName(output)) Then
        Err.Raise 5, , "Word output folder not found: " & output
    End If
    temporary = fileSystem.BuildPath(fileSystem.GetParentFolderName(output), fileSystem.GetTempName & ".docx")
    If existing Then
        originalStamp = VBA.CStr(fileSystem.GetFile(output).DateLastModified) & ":" & VBA.CStr(fileSystem.GetFile(output).Size)
        fileSystem.CopyFile output, temporary, False
    End If

    ' Reuse the workbook-owned Word process while keeping documents short-lived.
    Set wordApp = ex_PEB_WordState.fn_AcquireWord()
    wordApp.Visible = False
    wordApp.ScreenUpdating = False
    wordApp.DisplayAlerts = 0
    wordApp.AutomationSecurity = 3
    If existing Then
        Set document = wordApp.Documents.Open(temporary, False, False, False)
    ElseIf m_documentMode = "template" Then
        Set document = wordApp.Documents.Add(m_templatePath, False, 0, False)
    Else
        Set document = wordApp.Documents.Add
    End If
    If Not existing And m_documentMode = "template" Then
        private_ReplaceDocumentFields document, values
    End If
    If Not existing And m_documentMode = "blank" Then
        identifier = ex_PEB_Export.fn_Config("export.Word.textAnchors." & VBA.CStr(values("EventId")))
        document.Content.Text = "{\export:" & identifier & "_Begin}" & VBA.vbCr & "{\export:" & identifier & "_End}" & VBA.vbCr
    End If
    ex_PEB_WordDocument.fn_Change document, values, operation, eventUid, resultCode
    If existing Then
        document.Save
    Else
        document.SaveAs2 temporary, 16
    End If
    document.Close 0
    Set document = Nothing
    Set wordApp = Nothing

    ' Publish validated grouped orders by replacing their previous version.
    If existing Then
        If originalStamp <> VBA.CStr(fileSystem.GetFile(output).DateLastModified) & ":" & VBA.CStr(fileSystem.GetFile(output).Size) Then
            Err.Raise 5, , "Word output changed during export; retry from the current document."
        End If
        ex_PEB_WordState.fn_Publish temporary, output
    Else
        fileSystem.MoveFile temporary, output
    End If
    published = True
    values("WordExportAction") = resultCode
    values("WordEventUid") = eventUid
    private_Export = True
Cleanup:
    On Error Resume Next
    If Not document Is Nothing Then
        document.Close 0
    End If
    Set wordApp = Nothing
    If VBA.Err.Number <> 0 Then
        cleanupError = VBA.Err.Description
    End If
    VBA.Err.Clear
    If Not published And VBA.Len(temporary) > 0 Then
        If fileSystem.FileExists(temporary) Then
            fileSystem.DeleteFile temporary
        End If
    End If
    If VBA.Err.Number <> 0 Then
        cleanupError = cleanupError & " " & VBA.Err.Description
    End If
    On Error GoTo 0
    If VBA.Len(cleanupError) > 0 Then
        diagnostic = diagnostic & " Cleanup: " & cleanupError
    End If
    Exit Function
EH:
    If operation = "INSERT" And VBA.Err.Description = ex_PEB_Export.fn_Config("text.WordStoreDuplicate") Then
        values("WordExportAction") = "DUPLICATE"
        diagnostic = VBA.Err.Description
        private_Export = True
        Resume Cleanup
    End If
    diagnostic = "Word export: " & VBA.Err.Description
    Resume Cleanup
End Function

Private Sub private_ReplaceDocumentFields( _
    ByVal document As Object, _
    ByVal values As Object _
)
    Dim replacements As Object
    Dim story As Object
    Dim current As Object
    Dim target As Object
    Dim token As Variant
    Dim position As Long

    If Not values.Exists("WordDocumentReplacements") Then
        Err.Raise 5, , "Missing Word document replacements."
    End If
    Set replacements = values("WordDocumentReplacements")
    For Each story In document.StoryRanges
        Set current = story
        Do While Not current Is Nothing
            For Each token In replacements.Keys
                position = current.Start
                Do
                    Set target = current.Duplicate
                    target.SetRange position, current.End
                    With target.Find
                        .ClearFormatting
                        .Text = VBA.CStr(token)
                        .Forward = True
                        .Wrap = 0
                        .Format = False
                        .MatchCase = True
                        .MatchWholeWord = False
                        .MatchWildcards = False
                        .MatchSoundsLike = False
                        .MatchAllWordForms = False
                        If Not .Execute Then Exit Do
                    End With
                    target.Text = VBA.CStr(replacements(token))
                    position = target.End
                Loop
            Next token
            Set current = current.NextStoryRange
        Loop
    Next story
End Sub