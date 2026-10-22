VERSION 1.0 CLASS
BEGIN
  MultiUse = -1
END
Attribute VB_Name = "obj_PEB_WordTemplates"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private m_isInitialized As Boolean
Private m_isDisposed As Boolean

Private m_templates As Object
Private m_sharedTemplates As Object
Private m_resources As Object
Private m_resourcesXml As String
Private m_documentFilepath As String

' //
' // Lifecycle
' //
Private Sub Class_Initialize()
End Sub

Private Sub Class_Terminate()
    Me.Dispose
End Sub

' //
' // Properties
' //
Public Property Get SharedTemplates() As Object
    If Not m_isInitialized Or m_isDisposed Then
        Err.Raise 5, , "Word templates are not initialized."
    End If
    Set SharedTemplates = m_sharedTemplates
End Property

Public Property Get DocumentFilepath() As String
    If Not m_isInitialized Or m_isDisposed Then
        Err.Raise 5, , "Word templates are not initialized."
    End If
    DocumentFilepath = m_documentFilepath
End Property

Public Property Get Resources() As Object
    If Not m_isInitialized Or m_isDisposed Then
        Err.Raise 5, , "Word templates are not initialized."
    End If
    Set Resources = m_resources
End Property

Public Property Get ResourcesXml() As String
    If Not m_isInitialized Or m_isDisposed Then
        Err.Raise 5, , "Word templates are not initialized."
    End If
    ResourcesXml = m_resourcesXml
End Property

' //
' // API
' //
Public Function Initialize(ByRef diagnostic As String) As Boolean
    Dim document As Object
    Dim node As Object
    Dim textNode As Object
    Dim nodes As Object
    Dim name As String
    Dim path As String
    Dim fileSystem As Object
    Dim reference As Object
    Dim expression As Object
    Dim knownKeys As String

    diagnostic = VBA.vbNullString
    If m_isInitialized Or m_isDisposed Then
        diagnostic = "Word templates are already initialized or disposed."
        Exit Function
    End If
    On Error GoTo EH
    path = ex_PEB_Export.fn_Path(ex_PEB_Export.fn_Config("export.Word.templates.path"))
    Set fileSystem = VBA.CreateObject("Scripting.FileSystemObject")
    If Not fileSystem.FileExists(path) Then
        Err.Raise 5, , "Word DSL XML not found: " & path
    End If

    ' Загружаем XML без внешних сущностей и разрешения DTD.
    Set document = VBA.CreateObject("MSXML2.DOMDocument.6.0")
    document.async = False
    document.resolveExternals = False
    document.validateOnParse = False
    document.setProperty "ProhibitDTD", True
    If Not document.Load(path) Then
        Err.Raise 5, , "Word DSL XML: " & document.parseError.reason
    End If
    document.setProperty "SelectionNamespaces", "xmlns:p='urn:excelprototype:profiles'"
    Set node = document.selectSingleNode("/p:wordResultTemplates")
    If node Is Nothing Then
        Err.Raise 5, , "Expected wordResultTemplates in urn:excelprototype:profiles."
    End If
    If node.getAttributeNode("version") Is Nothing Then
        Err.Raise 5, , "Word DSL XML version is required."
    End If
    If node.getAttribute("version") <> "2" Then
        Err.Raise 5, , "Unsupported Word DSL XML version."
    End If
    knownKeys = "|" & ex_PEB_Export.fn_Config("export.Word.context.xmlKeys") & "|"
    Set m_resources = ex_PEB_DslResources.fn_Load(node, knownKeys & "ModeId|")
    If node.selectNodes("p:resources").Length = 1 Then
        m_resourcesXml = node.selectSingleNode("p:resources").XML
    End If
    For Each textNode In document.selectNodes("//p:derived | //p:aggregate")
        knownKeys = knownKeys & VBA.CStr(textNode.getAttribute("name")) & "|"
    Next textNode
    Set expression = VBA.CreateObject("VBScript.RegExp")
    expression.Global = True
    expression.Pattern = "\[([A-Z][A-Za-z0-9_]*)\]"
    For Each textNode In document.selectNodes("//p:text")
        ex_PEB_DslResources.fn_ValidateText textNode.Text, m_resources
        For Each reference In expression.Execute(textNode.Text)
            name = VBA.CStr(reference.SubMatches(0))
            If VBA.InStr(1, knownKeys, "|" & name & "|", VBA.vbBinaryCompare) = 0 Then
                Err.Raise 5, , "Word XML placeholder has no declared context provider: " & name
            End If
        Next reference
    Next textNode
    Set m_templates = VBA.CreateObject("Scripting.Dictionary")
    m_templates.CompareMode = VBA.vbBinaryCompare
    Set m_sharedTemplates = VBA.CreateObject("Scripting.Dictionary")
    m_sharedTemplates.CompareMode = VBA.vbBinaryCompare
    For Each node In document.selectNodes("/p:wordResultTemplates/p:sharedTemplates/p:sharedTemplate")
        name = VBA.Trim$(VBA.CStr(node.getAttribute("id")))
        If VBA.Len(name) = 0 Or m_sharedTemplates.Exists(name) Then
            Err.Raise 5, , "Empty or duplicate shared template ID: " & name
        End If
        Set nodes = node.selectNodes("p:text")
        If nodes.Length <> 1 Then
            Err.Raise 5, , "Shared template requires one text element: " & name
        End If
        m_sharedTemplates.Add name, VBA.CStr(nodes.Item(0).Text)
    Next node
    For Each node In document.selectNodes("/p:wordResultTemplates/p:template")
        name = VBA.Trim$(VBA.CStr(node.getAttribute("name")))
        If VBA.Len(name) = 0 Or m_templates.Exists(name) Then
            Err.Raise 5, , "Empty or duplicate template name: " & name
        End If
        ex_PEB_WordGroups.fn_Validate node
        m_templates.Add name, node
    Next node
    If m_templates.Count = 0 Then
        Err.Raise 5, , "Word DSL XML contains no templates."
    End If
    Set textNode = document.selectSingleNode("/p:wordResultTemplates/p:documentFilepath")
    If Not textNode Is Nothing Then
        m_documentFilepath = VBA.CStr(textNode.Text)
    End If
    m_isInitialized = True
    Initialize = True
    Exit Function
EH:
    diagnostic = VBA.Err.Description
    Me.Dispose
End Function

Public Sub Dispose()
    If m_isDisposed Then
        Exit Sub
    End If
    m_isDisposed = True
    m_isInitialized = False
    Set m_templates = Nothing
    Set m_sharedTemplates = Nothing
    Set m_resources = Nothing
    m_resourcesXml = VBA.vbNullString
    m_documentFilepath = VBA.vbNullString
End Sub

Public Function GroupTemplate(ByVal name As String) As Object
    If Not m_isInitialized Or m_isDisposed Then
        Err.Raise 5, , "Word templates are not initialized."
    End If
    If Not m_templates.Exists(name) Then
        Err.Raise 5, , "Unknown Word template: " & name
    End If
    If m_templates(name).selectNodes("p:group").Length > 1 Then
        Err.Raise 5, , "Expected at most one root Word group."
    End If
    If m_templates(name).selectNodes("p:group").Length = 1 Then
        Set GroupTemplate = m_templates(name)
    End If
End Function

Public Function Definition(ByVal name As String) As Object
    If Not m_isInitialized Or m_isDisposed Then
        Err.Raise 5, , "Word templates are not initialized."
    End If
    If Not m_templates.Exists(name) Then
        Err.Raise 5, , "Unknown Word template: " & name
    End If
    Set Definition = m_templates(name)
End Function

Public Function TryRender( _
    ByVal name As String, _
    ByVal values As Object, _
    ByRef rendered As String, _
    ByRef diagnostic As String _
) As Boolean
    Dim events As Collection

    On Error GoTo EH
    rendered = VBA.vbNullString
    diagnostic = VBA.vbNullString
    If Not m_isInitialized Or m_isDisposed Then
        Err.Raise 5, , "Word templates are not initialized."
    End If
    If Not m_templates.Exists(name) Then
        Err.Raise 5, , "Word DSL template not found: " & name
    End If
    Set events = New Collection
    events.Add values
    rendered = ex_PEB_WordGroups.fn_Render(m_templates(name), events, m_sharedTemplates, m_resources)
    TryRender = True
    Exit Function
EH:
    diagnostic = VBA.Err.Description
End Function