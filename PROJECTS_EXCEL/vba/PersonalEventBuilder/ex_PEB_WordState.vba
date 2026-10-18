Attribute VB_Name = "ex_PEB_WordState"
Option Explicit

Private Const STATE_NS As String = "urn:excelprototype:word-events:v1"
Private Const STATE_VERSION As String = "2"

#If VBA7 Then
Private Declare PtrSafe Function private_ReplaceFile Lib "kernel32" Alias "ReplaceFileW" (ByVal target As LongPtr, ByVal replacement As LongPtr, ByVal backup As LongPtr, ByVal flags As Long, ByVal exclude As LongPtr, ByVal reserved As LongPtr) As Long
#Else
Private Declare Function private_ReplaceFile Lib "kernel32" Alias "ReplaceFileW" (ByVal target As Long, ByVal replacement As Long, ByVal backup As Long, ByVal flags As Long, ByVal exclude As Long, ByVal reserved As Long) As Long
#End If

' --------------------------------------
' namespace API {
' --------------------------------------
Public Sub fn_Apply( _
    ByVal document As Object, _
    ByVal values As Object, _
    Optional ByRef resultCode As String = "" _
)
    Dim templates As obj_PEB_WordTemplates
    Dim template As Object
    Dim parts As Object
    Dim part As Object
    Dim state As Object
    Dim root As Object
    Dim section As Object
    Dim node As Object
    Dim replacement As Object
    Dim events As Collection
    Dim item As Object
    Dim scalar As Variant
    Dim key As String
    Dim name As String
    Dim orderKey As String
    Dim plan As String
    Dim baseline As String
    Dim rendered As String
    Dim duplicate As Boolean
    Dim context As Object
    Dim identities As Object
    Dim anchors As obj_PEB_WordAnchors
    Dim target As Object
    Dim diagnostic As String

    Set templates = New obj_PEB_WordTemplates
    If Not templates.Initialize(diagnostic) Then
        Err.Raise 5, , diagnostic
    End If
    name = VBA.CStr(ex_PEB_Export.fn_Required(values, "TemplateName"))
    Set template = templates.GroupTemplate(name)
    If template Is Nothing Then
        Err.Raise 5, , "Word state requires a grouped template."
    End If
    orderKey = ex_PEB_WordGroups.fn_Key(values, template.getAttribute("orderBy"))
    key = ex_PEB_WordGroups.fn_Key(values, template.getAttribute("identityBy"))
    ' Keep the exact plan: editing a live XML template requires an explicit migration.
    plan = private_Normalize(template.xml)
    Set parts = document.CustomXMLParts.SelectByNamespace(STATE_NS)
    If parts.Count > 1 Then
        Err.Raise 5, , "Duplicate Word event stores."
    End If
    Set state = VBA.CreateObject("MSXML2.DOMDocument.6.0")
    state.preserveWhiteSpace = True
    state.setProperty "ProhibitDTD", True
    state.resolveExternals = False
    If parts.Count = 1 Then
        Set part = parts(1)
        If Not state.LoadXML(part.XML) Then
            Err.Raise 5, , "Invalid embedded Word event XML."
        End If
    Else
        If Not state.LoadXML("<state xmlns=""" & STATE_NS & """ version=""" & STATE_VERSION & """/>") Then
            Err.Raise 5
        End If
    End If
    state.setProperty "SelectionNamespaces", "xmlns:s='" & STATE_NS & "'"
    Set root = state.selectSingleNode("/s:state")
    If root Is Nothing Then
        Err.Raise 5, , "Invalid Word event store root."
    End If
    If root.getAttributeNode("version") Is Nothing Then
        Err.Raise 5, , "Missing Word event store version."
    End If
    If root.getAttribute("version") <> STATE_VERSION Then
        Err.Raise 5, , "Unsupported Word event store version; migration is required."
    End If
    If root.getAttributeNode("order") Is Nothing Then
        If Not part Is Nothing Then
            Err.Raise 5, , "Missing stored Word order key."
        End If
        root.setAttribute "order", orderKey
    ElseIf root.getAttribute("order") <> orderKey Then
        Err.Raise 5, , "Word document belongs to a different order."
    End If
    For Each node In root.childNodes
        If node.nodeType = 1 Then
            If node.namespaceURI <> STATE_NS Or node.baseName <> "section" Then
                Err.Raise 5, , "Unknown Word event store element."
            End If
            If node.getAttributeNode("name") Is Nothing Then
                Err.Raise 5, , "Missing stored Word section name."
            End If
        End If
    Next node
    For Each node In root.selectNodes("s:section")
        If node.getAttribute("name") = name Then
            If Not section Is Nothing Then
                Err.Raise 5, , "Duplicate stored Word section."
            End If
            Set section = node
        End If
    Next node
    Set anchors = New obj_PEB_WordAnchors
    If Not anchors.Initialize(document, diagnostic) Then
        Err.Raise 5, , diagnostic
    End If
    If Not anchors.TryGetInsertedRange(document, values, target, diagnostic) Then
        Err.Raise 5, , diagnostic
    End If
    If section Is Nothing Then
        baseline = VBA.Trim$(VBA.Replace$(private_Normalize(target.Text), VBA.vbLf, " "))
        If VBA.Len(baseline) <> 0 Then
            Err.Raise 5, , "Word anchor contains text without embedded event data; migration is required."
        End If
        Set section = state.createNode(1, "section", STATE_NS)
        section.setAttribute "name", name
        private_AddText state, section, "plan", plan
        private_AddText state, section, "shared", private_SharedPlan(templates.SharedTemplates)
        root.appendChild section
    Else
        If private_Text(section, "s:plan") <> plan Then
            Err.Raise 5, , "Word grouping template changed; migration is required."
        End If
        If private_Text(section, "s:shared") <> private_SharedPlan(templates.SharedTemplates) Then
            Err.Raise 5, , "Shared Word templates changed; migration is required."
        End If
        If private_Normalize(target.Text) <> private_Text(section, "s:rendered") Then
            Err.Raise 5, , "Managed Word section was edited manually; export cancelled."
        End If
    End If
    Set events = New Collection
    Set identities = VBA.CreateObject("Scripting.Dictionary")
    For Each node In section.selectNodes("s:event")
        Set item = private_ReadEvent(node)
        baseline = private_Snapshot(item)
        If item("EventId") <> values("EventId") Or item("TemplateName") <> name Then
            Err.Raise 5, , "Stored Word event route differs from its section."
        End If
        If ex_PEB_WordGroups.fn_Key(item, template.getAttribute("orderBy")) <> orderKey Then
            Err.Raise 5, , "Stored Word event belongs to a different order."
        End If
        scalar = ex_PEB_WordGroups.fn_Key(item, template.getAttribute("identityBy"))
        If identities.Exists(scalar) Then
            Err.Raise 5, , "Duplicate stored Word event identity."
        End If
        identities.Add scalar, True
        If scalar = key Then
            If duplicate Then
                Err.Raise 5, , "Duplicate stored Word event identity."
            End If
            duplicate = True
            If private_Snapshot(values) <> private_Snapshot(item) Then
                Err.Raise 5, , "Word event already exists with different data."
            End If
        End If
        events.Add item
    Next node
    If Not duplicate Then
        Set replacement = private_NewEvent(state, values)
        section.appendChild replacement
        events.Add private_ReadEvent(replacement)
    End If
    rendered = ex_PEB_WordGroups.fn_Render(template, events, templates.SharedTemplates)
    Set context = VBA.CreateObject("Scripting.Dictionary")
    context.Add "EventId", values("EventId")
    context.Add "WordText", rendered
    If Not anchors.TryInsert(document, context, diagnostic) Then
        Err.Raise 5, , diagnostic
    End If
    If Not anchors.TryVerify(document, context, diagnostic) Then
        Err.Raise 5, , diagnostic
    End If
    ' Keep durable text anchors out of the printed order without hiding inserted text.
    anchors.HideMarkers document, context
    Set node = section.selectSingleNode("s:rendered")
    If Not node Is Nothing Then
        section.removeChild node
    End If
    private_AddText state, section, "rendered", private_Normalize(rendered)
    ' A new part is verified before the old part is deleted; callers edit a temporary DOCX.
    Set replacement = document.CustomXMLParts.Add(state.XML)
    If replacement Is Nothing Then
        Err.Raise 5, , "Cannot save embedded Word events."
    End If
    If Not state.LoadXML(replacement.XML) Then
        Err.Raise 5, , "Cannot verify embedded Word events."
    End If
    If state.selectNodes("/s:state/s:section").Length = 0 Then
        Err.Raise 5, , "Embedded Word events were not saved."
    End If
    If Not part Is Nothing Then
        part.Delete
    End If
    templates.Dispose
    anchors.Dispose
    If duplicate Then
        resultCode = "SKIP"
    Else
        resultCode = "INSERT"
    End If
End Sub

Public Sub fn_Publish( _
    ByVal temporary As String, _
    ByVal output As String _
)
    Dim files As Object
    Dim backup As String
    Dim errorNumber As Long

    Set files = VBA.CreateObject("Scripting.FileSystemObject")
    backup = files.BuildPath(files.GetParentFolderName(output), files.GetTempName & ".bak")
    If private_ReplaceFile(VBA.StrPtr(output), VBA.StrPtr(temporary), VBA.StrPtr(backup), 0, 0, 0) = 0 Then
        errorNumber = VBA.Err.LastDllError
        If Not files.FileExists(output) And files.FileExists(backup) Then
            files.MoveFile backup, output
        End If
        Err.Raise 5, , "Cannot atomically replace Word output; Windows error " & VBA.CStr(errorNumber) & "; recovery backup: " & backup
    End If
    private_DeleteBackup files, backup
End Sub
' --------------------------------------
' } // namespace API
' --------------------------------------

Private Function private_NewEvent( _
    ByVal state As Object, _
    ByVal values As Object _
) As Object
    Dim eventNode As Object
    Dim node As Object
    Dim key As Variant
    Dim scalar As Variant

    Set eventNode = state.createNode(1, "event", STATE_NS)
    For Each key In VBA.Split(ex_PEB_Export.fn_Config("export.Word.grouping.stateFields"), "|")
        scalar = ex_PEB_Export.fn_Required(values, VBA.CStr(key))
        If VBA.IsObject(scalar) Or VBA.IsArray(scalar) Or VBA.IsNull(scalar) Or VBA.IsError(scalar) Then
            Err.Raise 5, , "Embedded Word state requires scalar fields: " & key
        End If
        Set node = state.createNode(1, "field", STATE_NS)
        node.setAttribute "name", key
        ' Persist canonical rendered scalars, avoiding locale-dependent date/number parsing on reopen.
        node.Text = private_Normalize(ex_PEB_Export.fn_Text(scalar))
        eventNode.appendChild node
    Next key
    Set private_NewEvent = eventNode
End Function

Private Sub private_DeleteBackup( _
    ByVal files As Object, _
    ByVal path As String _
)
    On Error GoTo EH
    If files.FileExists(path) Then
        files.DeleteFile path
    End If
    Exit Sub
EH:
    ex_Core.fn_Diagnostic_WriteError "Word.Publish.BackupCleanup", VBA.Err.Number, VBA.Err.Source, VBA.Err.Description
End Sub

Private Function private_ReadEvent(ByVal eventNode As Object) As Object
    Dim result As Object
    Dim node As Object
    Dim name As String
    Dim allowed As String

    Set result = VBA.CreateObject("Scripting.Dictionary")
    allowed = "|" & ex_PEB_Export.fn_Config("export.Word.grouping.stateFields") & "|"
    For Each node In eventNode.selectNodes("s:field")
        name = VBA.CStr(node.getAttribute("name"))
        If VBA.Len(name) = 0 Or result.Exists(name) Then
            Err.Raise 5, , "Invalid or duplicate embedded Word field."
        End If
        If VBA.InStr(1, allowed, "|" & name & "|", VBA.vbBinaryCompare) = 0 Then
            Err.Raise 5, , "Unknown embedded Word field: " & name
        End If
        result.Add name, node.Text
    Next node
    Set private_ReadEvent = result
End Function

Private Function private_Snapshot(ByVal values As Object) As String
    Dim field As Variant
    Dim text As String
    Dim result As String

    For Each field In VBA.Split(ex_PEB_Export.fn_Config("export.Word.grouping.stateFields"), "|")
        text = private_Normalize(ex_PEB_Export.fn_Text(ex_PEB_Export.fn_Required(values, VBA.CStr(field))))
        result = result & VBA.Len(text) & ":" & text
    Next field
    private_Snapshot = result
End Function

Private Function private_SharedPlan(ByVal sharedTemplates As Object) As String
    Dim key As Variant
    Dim text As String
    Dim result As String

    For Each key In sharedTemplates.Keys
        text = private_Normalize(VBA.CStr(sharedTemplates(key)))
        result = result & VBA.Len(key) & ":" & key & VBA.Len(text) & ":" & text
    Next key
    private_SharedPlan = result
End Function

Private Sub private_AddText( _
    ByVal state As Object, _
    ByVal parent As Object, _
    ByVal name As String, _
    ByVal text As String _
)
    Dim node As Object

    Set node = state.createNode(1, name, STATE_NS)
    node.Text = text
    parent.appendChild node
End Sub

Private Function private_Text( _
    ByVal parent As Object, _
    ByVal path As String _
) As String
    Dim nodes As Object

    Set nodes = parent.selectNodes(path)
    If nodes.Length <> 1 Then
        Err.Raise 5, , "Missing or duplicate Word state element: " & path
    End If
    private_Text = nodes(0).Text
End Function

Private Function private_Normalize(ByVal text As String) As String
    private_Normalize = VBA.Replace$(VBA.Replace$(text, VBA.vbCrLf, VBA.vbLf), VBA.vbCr, VBA.vbLf)
End Function