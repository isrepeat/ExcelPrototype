Attribute VB_Name = "ex_TDB_BasisText"
Option Explicit

' --------------------------------------
' namespace Lifecycle {
' --------------------------------------
Public Sub fn_Module_Dispose()
End Sub
' --------------------------------------
' } // namespace Lifecycle
' --------------------------------------

' --------------------------------------
' namespace API {
' --------------------------------------
Public Function fn_Prepare( _
    ByVal fields As Object, _
    ByVal profiles As Collection _
) As Object
    Dim repository As Object
    Dim allowed As Object
    Dim field As Variant
    Dim profile As Variant
    Dim fileSystem As Object
    Dim folder As String
    Dim asset As Object

    Set repository = VBA.CreateObject("Scripting.Dictionary")
    repository.CompareMode = VBA.vbTextCompare
    Set allowed = VBA.CreateObject("Scripting.Dictionary")
    For Each field In fields.Keys
        allowed.Add VBA.CStr(field), fields(field)
    Next field
    Set fileSystem = VBA.CreateObject("Scripting.FileSystemObject")
    folder = ex_TravelDocumentsBuilder.fn_Config("basisChain.textTemplates.folder")
    If fileSystem.GetDriveName(folder) = VBA.vbNullString And VBA.Left$(folder, 2) <> "\\" Then
        folder = fileSystem.BuildPath(ThisWorkbook.Path, folder)
    End If
    folder = fileSystem.GetAbsolutePathName(folder)
    For Each profile In profiles
        private_Load repository, allowed, fileSystem.GetAbsolutePathName(folder), _
            ex_TravelDocumentsBuilder.fn_Config("basisChain.textTemplates.profiles." & profile), "|"
    Next profile
    Set allowed = VBA.CreateObject("Scripting.Dictionary")
    For Each field In VBA.Split("order.number|order.date|ownUnit.genitive", "|")
        allowed.Add VBA.CStr(field), True
    Next field
    private_Load repository, allowed, fileSystem.GetAbsolutePathName(folder), _
        ex_TravelDocumentsBuilder.fn_Config("basisChain.textTemplates.ownOrder"), "|"
    For Each field In repository.Keys
        If VBA.Left$(VBA.CStr(field), 1) <> "@" Then
            If Not fileSystem.FileExists(fileSystem.BuildPath(folder, field)) Then
                private_Fail fileSystem.BuildPath(folder, field)
            End If
            Set asset = fileSystem.GetFile(fileSystem.BuildPath(folder, field))
            If repository("@stamp:" & field) <> (VBA.CStr(asset.Size) & "|" & VBA.CStr(asset.DateLastModified)) Then
                private_Fail VBA.CStr(field)
            End If
        End If
    Next field
    Set fn_Prepare = repository
End Function

Public Function fn_Render( _
    ByVal repository As Object, _
    ByVal file As String, _
    ByVal values As Object _
) As String
    Dim text As String

    If Not repository.Exists(file) Then
        private_Fail file
    End If
    text = private_Children(repository(file), repository, values)
    text = VBA.Replace$(VBA.Replace$(VBA.Replace$(text, VBA.vbCr, " "), VBA.vbLf, " "), VBA.vbTab, " ")
    Do While VBA.InStr(text, "  ") > 0
        text = VBA.Replace$(text, "  ", " ")
    Loop
    text = VBA.Trim$(text)
    If VBA.Len(text) = 0 Then
        private_Fail file
    End If
    fn_Render = text
End Function
' --------------------------------------
' } // namespace API
' --------------------------------------

Private Sub private_Load( _
    ByVal repository As Object, _
    ByVal fields As Object, _
    ByVal folder As String, _
    ByVal file As String, _
    ByVal trail As String _
)
    Dim fileSystem As Object
    Dim document As Object
    Dim path As String
    Dim stamp As String

    Set fileSystem = VBA.CreateObject("Scripting.FileSystemObject")
    If UBound(VBA.Split(trail, "|")) > 64 Then
        private_Fail trail
    End If
    If fileSystem.GetFileName(file) <> file Or VBA.LCase$(fileSystem.GetExtensionName(file)) <> "xml" Or VBA.Left$(file, 1) = "@" Then
        private_Fail file
    End If
    If VBA.InStr(1, trail, "|" & file & "|", VBA.vbTextCompare) > 0 Then
        private_Fail trail & file
    End If
    If repository.Exists(file) Then
        private_Validate repository(file), repository, fields, folder, trail & file & "|"
        Exit Sub
    End If
    path = fileSystem.BuildPath(folder, file)
    If Not fileSystem.FileExists(path) Then
        private_Fail path
    End If
    stamp = VBA.CStr(fileSystem.GetFile(path).Size) & "|" & VBA.CStr(fileSystem.GetFile(path).DateLastModified)
    Set document = VBA.CreateObject("MSXML2.DOMDocument.6.0")
    document.async = False
    document.resolveExternals = False
    document.validateOnParse = False
    document.setProperty "ProhibitDTD", True
    If Not document.Load(path) Then
        private_Fail path & ": " & document.parseError.reason
    End If
    If stamp <> (VBA.CStr(fileSystem.GetFile(path).Size) & "|" & VBA.CStr(fileSystem.GetFile(path).DateLastModified)) Then
        private_Fail path
    End If
    If document.documentElement.nodeName <> "textTemplate" Then
        private_Fail path
    End If
    If private_Attribute(document.documentElement, "version") <> "1" Then
        private_Fail path
    End If
    repository.Add file, document.documentElement
    repository.Add "@stamp:" & file, stamp
    private_Validate document.documentElement, repository, fields, folder, trail & file & "|"
End Sub

Private Sub private_Validate( _
    ByVal node As Object, _
    ByVal repository As Object, _
    ByVal fields As Object, _
    ByVal folder As String, _
    ByVal trail As String _
)
    Dim child As Object
    Dim xmlAttribute As Object
    Dim name As String
    Dim allowed As String
    Dim required As String
    Dim cases As Object
    Dim defaults As Long
    Dim identifier As String
    Dim metadata As Object
    Dim options As Object
    Dim elements As Long

    name = node.nodeName
    Select Case name
        Case "textTemplate"
            allowed = "|version|"
        Case "text", "otherwise"
            allowed = "|"
        Case "value", "if", "ifFilled", "choose"
            allowed = "|field|"
            required = "field"
        Case "group"
            allowed = "|prefix|suffix|"
        Case "include"
            allowed = "|file|"
            required = "file"
        Case "case"
            allowed = "|value|"
            required = "value"
        Case Else
            private_Fail trail & name
    End Select
    If (name = "case" Or name = "otherwise") And node.parentNode.nodeName <> "choose" Then
        private_Fail trail & name
    End If
    If name = "textTemplate" And node.parentNode.nodeType <> 9 Then
        private_Fail trail & name
    End If
    For Each xmlAttribute In node.Attributes
        If VBA.InStr(allowed, "|" & xmlAttribute.nodeName & "|") = 0 Then
            private_Fail trail & xmlAttribute.nodeName
        End If
    Next xmlAttribute
    If VBA.Len(required) > 0 Then
        If VBA.Len(private_Attribute(node, required)) = 0 Then
            private_Fail trail & required
        End If
    End If
    If required = "field" Then
        If Not fields.Exists(private_Attribute(node, "field")) Then
            private_Fail trail & private_Attribute(node, "field")
        End If
        If IsObject(fields(private_Attribute(node, "field"))) Then
            Set metadata = fields(private_Attribute(node, "field"))
            If metadata.Exists("AllowedValues") Then
                Set options = metadata("AllowedValues")
            End If
        End If
        If name = "if" Then
            If options Is Nothing Then
                private_Fail trail & private_Attribute(node, "field")
            End If
            If options.Count <> 2 Or Not options.Exists("true") Or Not options.Exists("false") Then
                private_Fail trail & private_Attribute(node, "field")
            End If
        End If
    End If
    Set cases = VBA.CreateObject("Scripting.Dictionary")
    For Each child In node.childNodes
        If child.nodeType = 1 Then
            elements = elements + 1
            If name = "text" Or name = "value" Or name = "include" Then
                private_Fail trail & name
            End If
            If name = "choose" Then
                If child.nodeName = "case" Then
                    identifier = private_Attribute(child, "value")
                    If Not options Is Nothing Then
                        If Not options.Exists(identifier) Then
                            private_Fail trail & identifier
                        End If
                    End If
                    If cases.Exists(identifier) Then
                        private_Fail trail & identifier
                    End If
                    cases.Add identifier, True
                ElseIf child.nodeName = "otherwise" Then
                    defaults = defaults + 1
                    If defaults > 1 Then
                        private_Fail trail & name
                    End If
                Else
                    private_Fail trail & child.nodeName
                End If
            End If
            private_Validate child, repository, fields, folder, trail
        ElseIf child.nodeType <> 8 Then
            If name <> "text" And VBA.Len(VBA.Trim$(VBA.Replace$(VBA.Replace$(VBA.Replace$(child.Text, VBA.vbLf, " "), VBA.vbCr, " "), VBA.vbTab, " "))) > 0 Then
                private_Fail trail & name
            End If
        End If
    Next child
    If name = "choose" And cases.Count = 0 Then
        private_Fail trail & name
    End If
    If name = "textTemplate" And elements = 0 Then
        private_Fail trail & name
    End If
    If name = "include" Then
        private_Load repository, fields, folder, private_Attribute(node, "file"), trail
    End If
End Sub

Private Function private_Children( _
    ByVal node As Object, _
    ByVal repository As Object, _
    ByVal values As Object _
) As String
    Dim child As Object
    Dim text As String
    Dim output As String

    For Each child In node.childNodes
        If child.nodeType = 1 Then
            text = private_Render(child, repository, values)
            If VBA.Len(text) > 0 Then
                If VBA.Len(output) > 0 Then
                    output = output & " "
                End If
                output = output & text
            End If
        End If
    Next child
    private_Children = output
End Function

Private Function private_Render( _
    ByVal node As Object, _
    ByVal repository As Object, _
    ByVal values As Object _
) As String
    Dim child As Object
    Dim fallback As Object
    Dim text As String
    Dim value As String

    Select Case node.nodeName
        Case "text"
            private_Render = node.Text
        Case "value"
            private_Render = private_Value(node, values)
        Case "include"
            private_Render = private_Children(repository(private_Attribute(node, "file")), repository, values)
        Case "group"
            text = private_Children(node, repository, values)
            If VBA.Len(text) > 0 Then
                private_Render = private_Attribute(node, "prefix") & text & private_Attribute(node, "suffix")
            End If
        Case "ifFilled"
            If VBA.Len(VBA.Trim$(private_Value(node, values))) > 0 Then
                private_Render = private_Children(node, repository, values)
            End If
        Case "if"
            value = private_Value(node, values)
            If value <> "true" And value <> "false" Then
                private_Fail private_Attribute(node, "field")
            End If
            If value = "true" Then
                private_Render = private_Children(node, repository, values)
            End If
        Case "choose"
            value = private_Value(node, values)
            For Each child In node.childNodes
                If child.nodeType = 1 Then
                    If child.nodeName = "case" Then
                        If private_Attribute(child, "value") = value Then
                            private_Render = private_Children(child, repository, values)
                            Exit Function
                        End If
                    Else
                        Set fallback = child
                    End If
                End If
            Next child
            If Not fallback Is Nothing Then
                private_Render = private_Children(fallback, repository, values)
            End If
        Case Else
            private_Fail node.nodeName
    End Select
End Function

Private Function private_Value( _
    ByVal node As Object, _
    ByVal values As Object _
) As String
    Dim field As String

    field = private_Attribute(node, "field")
    If Not values.Exists(field) Then
        private_Fail field
    End If
    private_Value = VBA.CStr(values(field))
End Function

Private Function private_Attribute( _
    ByVal node As Object, _
    ByVal name As String _
) As String
    Dim xmlAttribute As Object

    Set xmlAttribute = node.Attributes.getNamedItem(name)
    If Not xmlAttribute Is Nothing Then
        private_Attribute = xmlAttribute.Text
    End If
End Function

Private Sub private_Fail(ByVal value As String)
    VBA.Err.Raise 5, "Basis.TextTemplate", VBA.Replace$(ex_TravelDocumentsBuilder.fn_Config("basisChain.textTemplates.text.Invalid"), "{value}", value)
End Sub