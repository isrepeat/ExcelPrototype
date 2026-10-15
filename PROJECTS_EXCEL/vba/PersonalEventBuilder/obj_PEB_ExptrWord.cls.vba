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
Private m_bookmark As String
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
        m_bookmark = ex_PEB_Export.fn_Config("export.Word.bookmark")
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

    On Error GoTo EH
    diagnostic = VBA.vbNullString
    If Not m_isInitialized Or m_isDisposed Then
        Err.Raise 5, , "Word exporter is not initialized."
    End If
    enabled = ex_PEB_Export.fn_Config("export.Word.enabled")
    If enabled = "false" Then
        Err.Raise 5, , ex_PEB_Export.fn_Config("text.ExportWordUnavailable")
    End If
    If enabled <> "true" Then
        Err.Raise 5, , "Word enabled must be true or false."
    End If
    output = ex_PEB_Export.fn_Text(ex_PEB_Export.fn_Required(values, "WordOutput"))
    rendered = VBA.CStr(ex_PEB_Export.fn_Required(values, "WordText"))
    Set fileSystem = VBA.CreateObject("Scripting.FileSystemObject")
    If m_documentMode = "template" Then
        If Not fileSystem.FileExists(m_templatePath) Then
            Err.Raise 5, , "Word document template not found: " & m_templatePath
        End If
    End If
    If VBA.LCase$(fileSystem.GetExtensionName(output)) <> "docx" Then
        Err.Raise 5, , "Word output must have a .docx extension."
    End If
    If fileSystem.FileExists(output) Then
        Err.Raise 5, , "Word output already exists: " & output
    End If
    If Not fileSystem.FolderExists(fileSystem.GetParentFolderName(output)) Then
        Err.Raise 5, , "Word output folder not found: " & output
    End If
    temporary = fileSystem.BuildPath(fileSystem.GetParentFolderName(output), fileSystem.GetTempName & ".docx")

    ' Отдельный экземпляр Word владеет создаваемым документом.
    Set wordApp = VBA.CreateObject("Word.Application")
    wordApp.Visible = False
    wordApp.DisplayAlerts = 0
    wordApp.AutomationSecurity = 3
    If m_documentMode = "template" Then
        Set document = wordApp.Documents.Add(m_templatePath, False, 0, False)
        If Not document.Bookmarks.Exists(m_bookmark) Then
            Err.Raise 5, , "Word template bookmark not found: " & m_bookmark
        End If
        Set target = document.Bookmarks(m_bookmark).Range
    Else
        Set document = wordApp.Documents.Add
        Set target = document.Content
    End If
    target.Text = VBA.Replace$(rendered, VBA.vbCrLf, VBA.vbCr)
    If m_documentMode = "template" Then
        document.Bookmarks.Add m_bookmark, target
    End If
    document.SaveAs2 temporary, 16
    document.Close 0
    Set document = Nothing
    wordApp.Quit 0
    Set wordApp = Nothing

    ' Публикуем готовый документ без перезаписи существующего файла.
    fileSystem.MoveFile temporary, output
    published = True
    TryExport = True
Cleanup:
    On Error Resume Next
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