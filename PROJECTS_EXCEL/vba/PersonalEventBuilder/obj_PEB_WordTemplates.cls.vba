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
    If node.getAttribute("version") <> "1" Then
        Err.Raise 5, , "Unsupported Word DSL XML version."
    End If
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
        Set nodes = node.selectNodes("p:text")
        If nodes.Length <> 1 Then
            Err.Raise 5, , "Template requires one text element: " & name
        End If
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
    m_documentFilepath = VBA.vbNullString
End Sub

Public Function TryGetText( _
    ByVal name As String, _
    ByRef templateText As String, _
    ByRef diagnostic As String _
) As Boolean
    Dim node As Object
    Dim header As Object
    Dim groupNames As Variant
    Dim groupName As Variant
    Dim headers As Object
    Dim by As String
    Dim result As String

    On Error GoTo EH
    templateText = VBA.vbNullString
    diagnostic = VBA.vbNullString
    If Not m_isInitialized Or m_isDisposed Then
        Err.Raise 5, , "Word templates are not initialized."
    End If
    If Not m_templates.Exists(name) Then
        Err.Raise 5, , "Word DSL template not found: " & name
    End If
    Set node = m_templates(name)

    ' Для одного события выводим заголовки групп перед текстом пункта.
    Set headers = VBA.CreateObject("Scripting.Dictionary")
    headers.CompareMode = VBA.vbBinaryCompare
    For Each header In node.selectNodes("p:groupHeader")
        by = VBA.Trim$(VBA.CStr(header.getAttribute("by")))
        If VBA.Len(by) = 0 Or headers.Exists(by) Then
            Err.Raise 5, , "Empty or duplicate groupHeader: " & name
        End If
        headers.Add by, VBA.CStr(header.Text)
    Next header
    If Not node.getAttributeNode("groupBy") Is Nothing Then
        by = VBA.Trim$(VBA.CStr(node.getAttribute("groupBy")))
        If VBA.Len(by) = 0 Then
            Err.Raise 5, , "groupHeader requires groupBy: " & name
        End If
        groupNames = VBA.Split(by, ";")
        For Each groupName In groupNames
            by = VBA.Trim$(VBA.CStr(groupName))
            If Not headers.Exists(by) Then
                Err.Raise 5, , "Missing groupHeader for " & name & ": " & by
            End If
            result = result & headers(by)
            headers.Remove by
        Next groupName
    End If
    If headers.Count > 0 Then
        Err.Raise 5, , "groupHeader is absent from groupBy: " & name
    End If
    result = result & node.selectSingleNode("p:text").Text
    templateText = result
    TryGetText = True
    Exit Function
EH:
    diagnostic = VBA.Err.Description
End Function