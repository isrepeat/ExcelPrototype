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
    TryExport = private_Export(values, diagnostic, False)
End Function

Public Function TryExportExistingTest( _
    ByVal values As Object, _
    ByRef diagnostic As String _
) As Boolean
    TryExportExistingTest = private_Export(values, diagnostic, True)
End Function

' //
' // Private
' //
Private Function private_Export( _
    ByVal values As Object, _
    ByRef diagnostic As String, _
    ByVal existingTest As Boolean _
) As Boolean
    Dim wordApp As Object
    Dim document As Object
    Dim target As Object
    Dim fileSystem As Object
    Dim output As String
    Dim rendered As String
    Dim temporary As String
    Dim published As Boolean
    Dim cleanupError As String
    Dim enabled As String
    Dim wordAnchors As obj_PEB_WordAnchors
    Dim grouped As Boolean
    Dim existing As Boolean
    Dim identifier As String
    Dim originalStamp As String

    On Error GoTo EH
    diagnostic = VBA.vbNullString
    If Not m_isInitialized Or m_isDisposed Then
        Err.Raise 5, , "Word exporter is not initialized."
    End If
    If ex_PEB_Export.fn_Text(ex_PEB_Export.fn_Required(values, "WordPolicy")) = "Omit" Then
        private_Export = True
        Exit Function
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
    rendered = VBA.CStr(ex_PEB_Export.fn_Required(values, "WordText"))
    grouped = VBA.CBool(ex_PEB_Export.fn_Required(values, "WordGrouped"))
    Set fileSystem = VBA.CreateObject("Scripting.FileSystemObject")
    existing = fileSystem.FileExists(output)
    If existingTest Then
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
    If existing And Not grouped Then
        Err.Raise 5, , "Word output already exists: " & output
    End If
    If Not fileSystem.FolderExists(fileSystem.GetParentFolderName(output)) Then
        Err.Raise 5, , "Word output folder not found: " & output
    End If
    temporary = fileSystem.BuildPath(fileSystem.GetParentFolderName(output), fileSystem.GetTempName & ".docx")
    If existing Then
        originalStamp = VBA.CStr(fileSystem.GetFile(output).DateLastModified) & ":" & VBA.CStr(fileSystem.GetFile(output).Size)
        fileSystem.CopyFile output, temporary, False
    End If

    ' Отдельный экземпляр Word владеет создаваемым документом.
    Set wordApp = VBA.CreateObject("Word.Application")
    wordApp.Visible = False
    wordApp.ScreenUpdating = True
    wordApp.DisplayAlerts = 0
    wordApp.AutomationSecurity = 3
    If existing Then
        Set document = wordApp.Documents.Open(temporary, False, False, False)
    ElseIf m_documentMode = "template" Then
        Set document = wordApp.Documents.Add(m_templatePath, False, 0, False)
    Else
        Set document = wordApp.Documents.Add
    End If
    If grouped Then
        If Not existing And m_documentMode = "blank" Then
            identifier = ex_PEB_Export.fn_Config("export.Word.textAnchors." & VBA.CStr(values("EventId")))
            document.Content.Text = "{\export:" & identifier & "_Begin}" & VBA.vbCr & "{\export:" & identifier & "_End}" & VBA.vbCr
        End If
        ex_PEB_WordState.fn_Apply document, values
    ElseIf m_documentMode = "template" Then
        Set wordAnchors = New obj_PEB_WordAnchors
        If Not wordAnchors.Initialize(document, diagnostic) Then GoTo Cleanup
        If Not wordAnchors.TryInsert(document, values, diagnostic) Then GoTo Cleanup
    Else
        Set target = document.Content
        target.Text = VBA.Replace$(rendered, VBA.vbCrLf, VBA.vbCr)
    End If
    If existing Then
        document.Save
    Else
        document.SaveAs2 temporary, 16
    End If
    document.Close 0
    Set document = Nothing
    wordApp.Quit 0
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
    private_Export = True
Cleanup:
    On Error Resume Next
    If Not wordAnchors Is Nothing Then wordAnchors.Dispose
    If Not document Is Nothing Then
        document.Close 0
    End If
    If Not wordApp Is Nothing Then
        wordApp.Quit 0
    End If
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
    diagnostic = "Word export: " & VBA.Err.Description
    Resume Cleanup
End Function