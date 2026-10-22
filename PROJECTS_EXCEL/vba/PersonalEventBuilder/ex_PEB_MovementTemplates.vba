Attribute VB_Name = "ex_PEB_MovementTemplates"
Option Explicit

Private m_document As Object
Private m_templates As Object
Private m_shared As Object
Private m_columnTexts As Object
Private m_knownFields As Object
Private m_resources As Object

' --------------------------------------
' namespace Lifecycle {
' --------------------------------------
Public Sub fn_Module_Dispose()
    Set m_document = Nothing
    Set m_templates = Nothing
    Set m_shared = Nothing
    Set m_columnTexts = Nothing
    Set m_knownFields = Nothing
    Set m_resources = Nothing
End Sub
' --------------------------------------
' } // namespace Lifecycle
' --------------------------------------

' --------------------------------------
' namespace API {
' --------------------------------------
Public Function fn_Template(ByVal name As String) As Object
    private_EnsureLoaded
    If Not m_templates.Exists(name) Then
        Err.Raise 5, , "Result text template not found: " & name
    End If
    Set fn_Template = m_templates(name)
End Function

Public Function fn_SharedTemplates() As Object
    private_EnsureLoaded
    Set fn_SharedTemplates = m_shared
End Function

Public Function fn_Texts(ByVal name As String) As Object
    private_EnsureLoaded
    If Not m_columnTexts.Exists(name) Then
        Err.Raise 5, , "Result column texts not found: " & name
    End If
    Set fn_Texts = m_columnTexts(name)
End Function

Public Function fn_Resources() As Object
    private_EnsureLoaded
    Set fn_Resources = m_resources
End Function
' --------------------------------------
' } // namespace API
' --------------------------------------

' --------------------------------------
' namespace Private {
' --------------------------------------
Private Sub private_EnsureLoaded()
    Dim root As Object
    Dim node As Object
    Dim child As Object
    Dim columnNode As Object
    Dim columns As Object
    Dim columnName As String
    Dim body As String
    Dim identifier As String
    Dim key As Variant
    Dim section As Variant
    Dim index As Long
    Dim propertyName As String
    Dim message As String

    If Not m_document Is Nothing Then
        Exit Sub
    End If
    On Error GoTo EH
    Set m_document = private_LoadXml(ex_PEB_Export.fn_Path(ex_PEB_Export.fn_Config("export.Movement.templates.path")))
    Set root = m_document.selectSingleNode("/p:resultTemplates")
    If root Is Nothing Then
        Err.Raise 5, , "Expected resultTemplates in urn:excelprototype:profiles."
    End If
    private_CheckAttributes root, "version"
    If private_Attribute(root, "version") <> "1" Then
        Err.Raise 5, , "Unsupported result template version."
    End If
    If root.selectNodes("p:sharedTemplates").Length <> 1 Then
        Err.Raise 5, , "Expected one sharedTemplates section."
    End If
    For Each child In root.childNodes
        If child.nodeType = 1 Then
            If child.namespaceURI <> "urn:excelprototype:profiles" Or _
                    InStr(1, "|resources|sharedTemplates|template|", "|" & child.baseName & "|", vbBinaryCompare) = 0 Then
                Err.Raise 5, , "Unknown result template element: " & child.nodeName
            End If
        End If
    Next child
    Set m_templates = private_Dictionary()
    Set m_shared = private_Dictionary()
    Set m_columnTexts = private_Dictionary()
    Set m_knownFields = private_Dictionary()
    For Each key In VBA.Split(ex_PEB_Export.fn_Config("export.fields") & "|" & ex_PEB_Export.fn_Config("export.shared.inputs.scalars"), "|")
        m_knownFields(key) = True
    Next key
    For Each key In VBA.Array("EventName", "SourceEventName", "SourceDestination", "TicketDocument", "FoodSupportDate", "SourceStatus", "TargetStatus", "Operation")
        m_knownFields(key) = True
    Next key
    For Each section In VBA.Array("lookups", "aliases", "constants")
        propertyName = "name"
        If section = "lookups" Then
            propertyName = "output"
        End If
        For index = 1 To ex_PEB_Export.fn_Count("export.shared." & section & ".count")
            m_knownFields(ex_PEB_Export.fn_Config("export.shared." & section & "." & index & "." & propertyName)) = True
        Next index
    Next section
    For Each key In VBA.Split(ex_PEB_Export.fn_Config("export.shared.inputs.collections"), "|")
        m_knownFields(key) = True
        For Each section In VBA.Split(ex_PEB_Export.fn_Config("export.shared.inputs." & key & ".fields"), "|")
            m_knownFields(section) = True
        Next section
    Next key
    Set m_resources = ex_PEB_DslResources.fn_Load(root, "|" & VBA.Join(m_knownFields.Keys, "|") & "|")
    For Each node In root.selectNodes("p:sharedTemplates/p:sharedTemplate")
        private_CheckAttributes node, "id"
        private_RequireChildren node, "text"
        identifier = private_Attribute(node, "id")
        If m_shared.Exists(identifier) Then
            Err.Raise 5, , "Duplicate shared text template: " & identifier
        End If
        m_shared.Add identifier, node.selectSingleNode("p:text").Text
    Next node
    For Each key In m_shared.Keys
        private_ValidateText m_shared(key), "|" & key & "|", 0
    Next key
    For Each node In root.selectNodes("p:template")
        private_CheckAttributes node, "name"
        identifier = private_Attribute(node, "name")
        private_Add m_templates, identifier, node
        Set columns = private_Dictionary()
        columns.CompareMode = VBA.vbTextCompare
        For Each child In node.childNodes
            If child.nodeType = 1 Then
                If child.namespaceURI <> "urn:excelprototype:profiles" Or child.baseName <> "column" Then
                    Err.Raise 5, , "Expected column in result template: " & identifier
                End If
            End If
        Next child
        For Each columnNode In node.selectNodes("p:column")
            private_CheckAttributes columnNode, "configKey"
            private_RequireChildren columnNode, "text"
            columnName = ex_PEB_Export.fn_Config(private_Attribute(columnNode, "configKey"))
            If columns.Exists(columnName) Then
                Err.Raise 5, , "Duplicate result column: " & identifier & "/" & columnName
            End If
            body = columnNode.selectSingleNode("p:text").Text
            If VBA.Len(VBA.Trim$(body)) = 0 Then
                Err.Raise 5, , "Empty result column text: " & identifier & "/" & columnName
            End If
            private_ValidateText body, "|", 0
            columns.Add columnName, body
        Next columnNode
        If columns.Count = 0 Then
            Err.Raise 5, , "Result template requires at least one column: " & identifier
        End If
        m_columnTexts.Add identifier, columns
    Next node
    If m_templates.Count = 0 Then
        Err.Raise 5, , "Result XML contains no text templates."
    End If
    Exit Sub
EH:
    message = VBA.Err.Description
    fn_Module_Dispose
    Err.Raise 5, , "Result text templates: " & message
End Sub

Private Sub private_ValidateText( _
    ByVal text As String, _
    ByVal trail As String, _
    ByVal depth As Long _
)
    Dim expression As Object
    Dim reference As Object
    Dim name As String

    ex_PEB_DslResources.fn_ValidateText text, m_resources

    If depth > ex_PEB_Export.fn_Count("export.dsl.maxDepth") Then
        Err.Raise 5, , "Movement include-depth limit exceeded."
    End If
    Set expression = VBA.CreateObject("VBScript.RegExp")
    expression.Global = True
    expression.Pattern = "\{#include\s+([^}]+)\}"
    For Each reference In expression.Execute(text)
        name = VBA.Trim$(VBA.CStr(reference.SubMatches(0)))
        If Not m_shared.Exists(name) Then
            Err.Raise 5, , "Unknown Movement include: " & name
        End If
        If VBA.InStr(1, trail, "|" & name & "|", VBA.vbBinaryCompare) > 0 Then
            Err.Raise 5, , "Cyclic Movement include: " & name
        End If
        private_ValidateText m_shared(name), trail & name & "|", depth + 1
    Next reference
    expression.Pattern = "\[([A-Z][A-Za-z0-9_]*)\]"
    For Each reference In expression.Execute(text)
        name = VBA.CStr(reference.SubMatches(0))
        If Not m_knownFields.Exists(name) Then
            Err.Raise 5, , "Unknown Movement placeholder: " & name
        End If
    Next reference
End Sub

Private Function private_LoadXml(ByVal path As String) As Object
    Dim document As Object

    Set document = VBA.CreateObject("MSXML2.DOMDocument.6.0")
    document.async = False
    document.resolveExternals = False
    document.validateOnParse = False
    document.setProperty "ProhibitDTD", True
    If Not document.Load(path) Then
        Err.Raise 5, , path & ":" & document.parseError.Line & ": " & document.parseError.reason
    End If
    document.setProperty "SelectionNamespaces", "xmlns:p='urn:excelprototype:profiles'"
    Set private_LoadXml = document
End Function

Private Function private_Dictionary() As Object
    Set private_Dictionary = VBA.CreateObject("Scripting.Dictionary")
    private_Dictionary.CompareMode = VBA.vbBinaryCompare
End Function

Private Sub private_Add( _
    ByVal dictionary As Object, _
    ByVal identifier As String, _
    ByVal value As Variant _
)
    Dim expression As Object

    Set expression = VBA.CreateObject("VBScript.RegExp")
    expression.Pattern = "^[A-Za-z][A-Za-z0-9_]*$"
    If Not expression.Test(identifier) Or dictionary.Exists(identifier) Then
        Err.Raise 5, , "Invalid or duplicate XML identifier: " & identifier
    End If
    dictionary.Add identifier, value
End Sub

Private Function private_Attribute( _
    ByVal node As Object, _
    ByVal name As String _
) As String
    private_Attribute = private_OptionalAttribute(node, name)
    If VBA.Len(private_Attribute) = 0 Then
        Err.Raise 5, , "Missing XML attribute: " & node.nodeName & "/@" & name
    End If
End Function

Private Function private_OptionalAttribute( _
    ByVal node As Object, _
    ByVal name As String _
) As String
    Dim xmlAttribute As Object

    Set xmlAttribute = node.getAttributeNode(name)
    If Not xmlAttribute Is Nothing Then
        private_OptionalAttribute = VBA.Trim$(VBA.CStr(xmlAttribute.Text))
    End If
End Function

Private Sub private_CheckAttributes( _
    ByVal node As Object, _
    ByVal allowed As String _
)
    Dim xmlAttribute As Object

    For Each xmlAttribute In node.Attributes
        If xmlAttribute.nodeName <> "xmlns" Then
            If VBA.InStr(1, "|" & allowed & "|", "|" & xmlAttribute.nodeName & "|", VBA.vbBinaryCompare) = 0 Then
                Err.Raise 5, , "Unknown XML attribute: " & node.nodeName & "/@" & xmlAttribute.nodeName
            End If
        End If
    Next xmlAttribute
End Sub

Private Sub private_RequireChildren( _
    ByVal node As Object, _
    ByVal names As String _
)
    Dim name As Variant
    Dim child As Object

    For Each name In VBA.Split(names, "|")
        If node.selectNodes("p:" & name).Length <> 1 Then
            Err.Raise 5, , "Expected one XML element: " & node.nodeName & "/" & name
        End If
    Next name
    For Each child In node.childNodes
        If child.nodeType = 1 Then
            If VBA.InStr(1, "|" & names & "|", "|" & child.baseName & "|", VBA.vbBinaryCompare) = 0 Then
                Err.Raise 5, , "Unknown XML element: " & child.nodeName
            End If
        End If
    Next child
End Sub

' --------------------------------------
' } // namespace Private
' --------------------------------------