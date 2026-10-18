VERSION 1.0 CLASS
BEGIN
  MultiUse = -1
END
Attribute VB_Name = "obj_PEB_WordEventStore"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private m_isInitialized As Boolean
Private m_isDisposed As Boolean

Private Const STORE_NS As String = "urn:excelprototype:word-events:v1"
Private m_document As Object
Private m_root As Object
Private m_scopes As Object
Private m_identityFields As String
Private m_temporaryIpn As String

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
Public Property Get Xml() As String
    private_RequireReady
    Xml = m_document.XML
End Property

Public Property Get Sections() As Object
    private_RequireReady
    Set Sections = m_root.selectNodes("s:sections/s:section")
End Property

' //
' // API
' //
Public Function Initialize( _
    ByVal xml As String, _
    ByVal orderKey As String, _
    ByRef diagnostic As String _
) As Boolean
    Dim scope As Variant
    Dim field As Variant

    If m_isInitialized Or m_isDisposed Then
        Exit Function
    End If
    On Error GoTo EH
    Set m_scopes = VBA.CreateObject("Scripting.Dictionary")
    For Each scope In VBA.Array("person", "position", "basis")
        For Each field In VBA.Split(ex_PEB_Export.fn_Config("export.Word.store." & scope & "Fields"), "|")
            If VBA.Len(field) = 0 Or m_scopes.Exists(field) Then
                Err.Raise 5, , "Empty or overlapping Word value map field: " & field
            End If
            m_scopes.Add VBA.CStr(field), VBA.CStr(scope)
        Next field
    Next scope
    m_identityFields = ex_PEB_Export.fn_Config("export.Word.store.identityFields")
    m_temporaryIpn = ex_PEB_Export.fn_Config("lookup.ReturnTickets.temporaryIpn")
    Set m_document = VBA.CreateObject("MSXML2.DOMDocument.6.0")
    m_document.async = False
    m_document.preserveWhiteSpace = True
    m_document.resolveExternals = False
    m_document.setProperty "ProhibitDTD", True
    If VBA.Len(xml) = 0 Then
        xml = "<state xmlns=""" & STORE_NS & """ version=""3"" nextUid=""1""><maps/><sections/></state>"
    End If
    If Not m_document.LoadXML(xml) Then
        Err.Raise 5, , "Invalid Word event store XML."
    End If
    m_document.setProperty "SelectionNamespaces", "xmlns:s='" & STORE_NS & "'"
    Set m_root = m_document.selectSingleNode("/s:state")
    If m_root Is Nothing Then
        Err.Raise 5, , "Invalid Word event store root."
    End If
    If m_root.getAttribute("version") <> "3" Then
        private_Fail "Migration"
    End If
    If m_root.selectNodes("s:maps").Length <> 1 Or m_root.selectNodes("s:sections").Length <> 1 Then
        Err.Raise 5, , "Invalid Word event store containers."
    End If
    If m_root.getAttributeNode("order") Is Nothing Then
        m_root.setAttribute "order", orderKey
    ElseIf m_root.getAttribute("order") <> orderKey Then
        private_Fail "OrderMismatch"
    End If
    private_Validate
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
    m_isInitialized = False
    m_isDisposed = True
    Set m_root = Nothing
    Set m_document = Nothing
    Set m_scopes = Nothing
End Sub

Public Function BindSection( _
    ByVal anchor As String, _
    ByVal templateName As String, _
    ByVal eventId As String, _
    ByVal plan As String, _
    ByVal sharedPlan As String _
) As Object
    Dim section As Object

    private_RequireReady
    ' MSXML text nodes normalize CRLF and CR to LF when read back.
    plan = VBA.Replace$(VBA.Replace$(plan, VBA.vbCrLf, VBA.vbLf), VBA.vbCr, VBA.vbLf)
    sharedPlan = VBA.Replace$(VBA.Replace$(sharedPlan, VBA.vbCrLf, VBA.vbLf), VBA.vbCr, VBA.vbLf)
    For Each section In Me.Sections
        If section.getAttribute("anchor") = anchor Then
            If section.getAttribute("template") <> templateName Or _
                    Me.Text(section, "plan") <> plan Or Me.Text(section, "shared") <> sharedPlan Then
                private_Fail "PlanChanged"
            End If
            Set BindSection = section
            Exit Function
        End If
    Next section
    Set section = private_Node("section")
    section.setAttribute "anchor", anchor
    section.setAttribute "template", templateName
    section.setAttribute "eventId", eventId
    private_AddText section, "plan", plan
    private_AddText section, "shared", sharedPlan
    m_root.selectSingleNode("s:sections").appendChild section
    Set BindSection = section
End Function

Public Sub Change( _
    ByVal operation As String, _
    ByVal section As Object, _
    ByVal values As Object, _
    ByVal contextKeys As Collection, _
    ByRef uid As String _
)
    Dim existing As Object
    Dim eventNode As Object
    Dim candidate As Object
    Dim person As Object
    Dim position As Object
    Dim basis As Object
    Dim fields As Object
    Dim key As Variant
    Dim identity As String
    Dim backup As String
    Dim errorNumber As Long
    Dim errorText As String
    Dim previousUid As String

    private_RequireReady
    backup = Me.Xml
    previousUid = uid
    On Error GoTo EH
    If operation <> "INSERT" And operation <> "UPDATE" And operation <> "DELETE" Then
        Err.Raise 5, , "Unknown Word event operation: " & operation
    End If
    Set existing = private_FindEvent(uid)
    If operation = "INSERT" Then
        If VBA.Len(uid) > 0 Then
            private_Fail "Duplicate"
        End If
    ElseIf existing Is Nothing Or VBA.Len(uid) = 0 Then
        private_Fail "NotFound"
    End If
    If operation = "DELETE" Then
        existing.ParentNode.removeChild existing
        private_CollectMaps
        Exit Sub
    End If

    Set person = private_Person(values)
    Set position = private_Map("position", private_Value(values, "PositionCode"))
    Set basis = private_Basis(values, contextKeys)
    Set eventNode = private_Node("event")
    If operation = "INSERT" Then
        uid = private_Uid("e")
    End If
    eventNode.setAttribute "uid", uid
    eventNode.setAttribute "personRef", person.getAttribute("id")
    eventNode.setAttribute "positionRef", position.getAttribute("id")
    eventNode.setAttribute "basisRef", basis.getAttribute("id")
    Set fields = VBA.CreateObject("Scripting.Dictionary")
    For Each key In contextKeys
        If Not values.Exists(key) Then
            Err.Raise 5, , "Missing stored Word context: " & key
        End If
        Select Case private_Scope(VBA.CStr(key))
            Case "person"
                private_SetMapField person, VBA.CStr(key), values(key), operation
            Case "position"
                private_SetMapField position, VBA.CStr(key), values(key), operation
            Case "basis"
                ' Basis maps are immutable and are selected by their complete values.
            Case Else
                private_WriteField eventNode, VBA.CStr(key), values(key)
        End Select
    Next key
    Set fields = private_Context(eventNode)
    fields("PersonKey") = person.getAttribute("id")
    identity = ex_PEB_WordGroups.fn_Key(fields, m_identityFields)
    eventNode.setAttribute "identity", identity
    For Each candidate In m_root.selectNodes("s:sections/s:section/s:event")
        If candidate.getAttribute("identity") = identity And candidate.getAttribute("uid") <> uid Then
            private_Fail "Duplicate"
        End If
    Next candidate
    If Not existing Is Nothing Then
        If existing.ParentNode Is section Then
            section.replaceChild eventNode, existing
        Else
            existing.ParentNode.removeChild existing
            section.appendChild eventNode
        End If
    Else
        section.appendChild eventNode
    End If
    private_CollectMaps
    private_Validate
    Exit Sub
EH:
    errorNumber = VBA.Err.Number
    errorText = VBA.Err.Description
    uid = previousUid
    m_document.LoadXML backup
    Set m_root = m_document.selectSingleNode("/s:state")
    Err.Raise errorNumber, , errorText
End Sub

Public Function Events(ByVal section As Object) As Collection
    Dim result As New Collection
    Dim node As Object
    Dim context As Object

    private_RequireReady
    For Each node In section.selectNodes("s:event")
        Set context = private_Context(node)
        context("WordEventUid") = node.getAttribute("uid")
        context("PersonKey") = node.getAttribute("personRef")
        result.Add context
    Next node
    Set Events = result
End Function

Public Function Text( _
    ByVal parent As Object, _
    ByVal name As String _
) As String
    Dim node As Object

    private_RequireReady
    Set node = parent.selectSingleNode("s:" & name)
    If Not node Is Nothing Then
        Text = node.Text
    End If
End Function

Public Sub SetRendered( _
    ByVal section As Object, _
    ByVal rendered As String _
)
    Dim node As Object

    private_RequireReady
    Set node = section.selectSingleNode("s:rendered")
    If Not node Is Nothing Then
        section.removeChild node
    End If
    private_AddText section, "rendered", rendered
End Sub

' //
' // Private
' //
Private Sub private_RequireReady()
    If Not m_isInitialized Or m_isDisposed Then
        Err.Raise 5, , "Word event store is not initialized."
    End If
End Sub

Private Function private_Person(ByVal values As Object) As Object
    Dim node As Object
    Dim found As Object
    Dim ipn As String
    Dim fio As String
    Dim position As String
    Dim marker As String

    ipn = private_Value(values, "PersonId")
    fio = private_Value(values, "PersonName")
    position = private_Value(values, "PositionCode")
    marker = m_temporaryIpn
    If VBA.StrComp(ipn, marker, VBA.vbTextCompare) = 0 Then
        ipn = VBA.vbNullString
    End If
    If VBA.Len(ipn) = 0 And (VBA.Len(fio) = 0 Or VBA.Len(position) = 0) Then
        private_Fail "PersonRequired"
    End If
    For Each node In m_root.selectNodes("s:maps/s:map[@scope='person']")
        If VBA.Len(ipn) > 0 And node.getAttribute("ipn") = ipn Then
            Set private_Person = node
            Exit Function
        End If
        If VBA.LCase$(ex_TableQuery.fn_NormalizeText(node.getAttribute("fio"))) = VBA.LCase$(ex_TableQuery.fn_NormalizeText(fio)) And _
                node.getAttribute("position") = position And VBA.Len(node.getAttribute("ipn")) = 0 Then
            If Not found Is Nothing Then
                private_Fail "PersonAmbiguous"
            End If
            Set found = node
        End If
    Next node
    If found Is Nothing Then
        Set found = private_Map("person", private_Uid("p"))
        found.setAttribute "fio", fio
        found.setAttribute "position", position
    End If
    If VBA.Len(ipn) > 0 Then
        found.setAttribute "ipn", ipn
    End If
    Set private_Person = found
End Function

Private Function private_Map( _
    ByVal scope As String, _
    ByVal key As String _
) As Object
    Dim node As Object

    For Each node In m_root.selectNodes("s:maps/s:map")
        If node.getAttribute("scope") = scope And node.getAttribute("key") = key Then
            Set private_Map = node
            Exit Function
        End If
    Next node
    Set node = private_Node("map")
    node.setAttribute "id", private_Uid("v")
    node.setAttribute "scope", scope
    node.setAttribute "key", key
    m_root.selectSingleNode("s:maps").appendChild node
    Set private_Map = node
End Function

Private Function private_Basis( _
    ByVal values As Object, _
    ByVal keys As Collection _
) As Object
    Dim probe As Object
    Dim node As Object
    Dim key As Variant
    Dim signature As String
    Dim fieldNode As Object

    Set probe = private_Node("basis")
    For Each key In keys
        If private_Scope(VBA.CStr(key)) = "basis" Then
            private_WriteField probe, VBA.CStr(key), values(key)
        End If
    Next key
    signature = private_FieldsXml(probe)
    For Each node In m_root.selectNodes("s:maps/s:map[@scope='basis']")
        If private_FieldsXml(node) = signature Then
            Set private_Basis = node
            Exit Function
        End If
    Next node
    Set node = private_Map("basis", private_Uid("b"))
    For Each fieldNode In probe.childNodes
        node.appendChild fieldNode.cloneNode(True)
    Next fieldNode
    Set private_Basis = node
End Function

Private Function private_FieldsXml(ByVal parent As Object) As String
    Dim node As Object
    Dim result As String

    For Each node In parent.selectNodes("s:field")
        result = result & node.XML
    Next node
    private_FieldsXml = result
End Function

Private Sub private_SetMapField( _
    ByVal parent As Object, _
    ByVal name As String, _
    ByVal value As Variant, _
    ByVal operation As String _
)
    Dim old As Object
    Dim probe As Object

    Set old = parent.selectSingleNode("s:field[@name='" & name & "']")
    Set probe = private_Node("probe")
    private_WriteField probe, name, value
    If Not old Is Nothing Then
        If old.XML <> probe.firstChild.XML And operation <> "UPDATE" Then
            private_Fail "MapConflict"
        End If
        parent.removeChild old
    End If
    parent.appendChild probe.firstChild
End Sub

Private Function private_Context(ByVal eventNode As Object) As Object
    Dim result As Object
    Dim scope As Variant
    Dim node As Object
    Dim map As Object
    Dim ref As String

    Set result = VBA.CreateObject("Scripting.Dictionary")
    result.CompareMode = VBA.vbBinaryCompare
    For Each scope In VBA.Array("person", "position", "basis")
        ref = eventNode.getAttribute(scope & "Ref")
        Set map = Nothing
        For Each node In m_root.selectNodes("s:maps/s:map")
            If node.getAttribute("id") = ref Then
                Set map = node
                Exit For
            End If
        Next node
        If map Is Nothing Then
            Err.Raise 5, , "Missing Word value map: " & ref
        End If
        private_ReadFields map, result
    Next scope
    private_ReadFields eventNode, result
    Set private_Context = result
End Function

Private Sub private_WriteField( _
    ByVal parent As Object, _
    ByVal name As String, _
    ByVal value As Variant _
)
    Dim node As Object
    Dim item As Variant
    Dim child As Object
    Dim key As Variant

    Set node = private_Node("field")
    node.setAttribute "name", name
    If VBA.IsObject(value) Then
        If VBA.TypeName(value) <> "Collection" Then
            Err.Raise 5, , "Unsupported Word context object: " & name
        End If
        node.setAttribute "kind", "collection"
        For Each item In value
            If VBA.TypeName(item) <> "Dictionary" Then
                Err.Raise 5, , "Word context collection items must be dictionaries."
            End If
            Set child = private_Node("item")
            For Each key In item.Keys
                private_WriteField child, VBA.CStr(key), item(key)
            Next key
            node.appendChild child
        Next item
    ElseIf VBA.VarType(value) = VBA.vbBoolean Then
        node.setAttribute "kind", "boolean"
        node.Text = VBA.CStr(value)
    ElseIf VBA.VarType(value) = VBA.vbDate Then
        node.setAttribute "kind", "date"
        node.Text = VBA.Format$(value, "yyyy-mm-dd")
    ElseIf VBA.IsError(value) Or VBA.IsNull(value) Or VBA.IsArray(value) Then
        Err.Raise 5, , "Invalid stored Word context: " & name
    Else
        node.setAttribute "kind", "text"
        node.Text = VBA.CStr(value)
    End If
    parent.appendChild node
End Sub

Private Sub private_ReadFields( _
    ByVal parent As Object, _
    ByVal output As Object _
)
    Dim node As Object
    Dim child As Object
    Dim rows As Collection
    Dim row As Object
    Dim name As String
    Dim parts As Variant

    For Each node In parent.selectNodes("s:field")
        name = node.getAttribute("name")
        If output.Exists(name) Then
            Err.Raise 5, , "Duplicate stored Word field: " & name
        End If
        Select Case node.getAttribute("kind")
            Case "collection"
                Set rows = New Collection
                For Each child In node.selectNodes("s:item")
                    Set row = VBA.CreateObject("Scripting.Dictionary")
                    private_ReadFields child, row
                    rows.Add row
                Next child
                Set output(name) = rows
            Case "boolean"
                output.Add name, (node.Text = "True")
            Case "date"
                parts = VBA.Split(node.Text, "-")
                output.Add name, VBA.DateSerial(VBA.CInt(parts(0)), VBA.CInt(parts(1)), VBA.CInt(parts(2)))
            Case "text"
                output.Add name, node.Text
            Case Else
                Err.Raise 5, , "Unknown stored Word field kind: " & name
        End Select
    Next node
End Sub

Private Function private_Scope(ByVal name As String) As String
    If m_scopes.Exists(name) Then
        private_Scope = m_scopes(name)
    End If
End Function

Private Function private_FindEvent(ByVal uid As String) As Object
    Dim node As Object

    For Each node In m_root.selectNodes("s:sections/s:section/s:event")
        If node.getAttribute("uid") = uid Then
            Set private_FindEvent = node
            Exit Function
        End If
    Next node
End Function

Private Sub private_CollectMaps()
    Dim used As Object
    Dim node As Object
    Dim scope As Variant

    Set used = VBA.CreateObject("Scripting.Dictionary")
    For Each node In m_root.selectNodes("s:sections/s:section/s:event")
        For Each scope In VBA.Array("person", "position", "basis")
            used(node.getAttribute(scope & "Ref")) = True
        Next scope
    Next node
    For Each node In m_root.selectNodes("s:maps/s:map")
        If Not used.Exists(node.getAttribute("id")) Then
            node.ParentNode.removeChild node
        End If
    Next node
End Sub

Private Sub private_Validate()
    Dim ids As Object
    Dim identities As Object
    Dim node As Object
    Dim context As Object
    Dim anchors As Object
    Dim probe As Object

    Set ids = VBA.CreateObject("Scripting.Dictionary")
    Set identities = VBA.CreateObject("Scripting.Dictionary")
    Set anchors = VBA.CreateObject("Scripting.Dictionary")
    If Not VBA.IsNumeric(m_root.getAttribute("nextUid")) Then
        Err.Raise 5, , "Missing Word UID sequence."
    End If
    For Each node In m_root.selectNodes("s:sections/s:section")
        If VBA.Len(node.getAttribute("anchor")) = 0 Or VBA.Len(node.getAttribute("template")) = 0 Or _
                VBA.Len(node.getAttribute("eventId")) = 0 Then
            Err.Raise 5, , "Missing Word section identity."
        End If
        If anchors.Exists(node.getAttribute("anchor")) Then
            Err.Raise 5, , "Duplicate Word section anchor."
        End If
        anchors.Add node.getAttribute("anchor"), True
        If node.selectNodes("s:plan").Length <> 1 Or node.selectNodes("s:shared").Length <> 1 Or _
                node.selectNodes("s:rendered").Length > 1 Then
            Err.Raise 5, , "Invalid Word section metadata."
        End If
    Next node
    For Each node In m_root.selectNodes("s:maps/s:map | s:sections/s:section/s:event")
        If node.baseName = "map" Then
            If VBA.Len(node.getAttribute("id")) = 0 Then
                Err.Raise 5, , "Missing Word value map UID."
            End If
            If ids.Exists(node.getAttribute("id")) Then
                Err.Raise 5, , "Duplicate Word map UID."
            End If
            ids.Add node.getAttribute("id"), True
            Set probe = VBA.CreateObject("Scripting.Dictionary")
            private_ReadFields node, probe
        Else
            If VBA.Len(node.getAttribute("uid")) = 0 Or VBA.Len(node.getAttribute("identity")) = 0 Then
                Err.Raise 5, , "Missing Word event identity."
            End If
            If ids.Exists(node.getAttribute("uid")) Or identities.Exists(node.getAttribute("identity")) Then
                private_Fail "Duplicate"
            End If
            ids.Add node.getAttribute("uid"), True
            identities.Add node.getAttribute("identity"), True
            Set context = private_Context(node)
            context("PersonKey") = node.getAttribute("personRef")
            If ex_PEB_WordGroups.fn_Key(context, m_identityFields) <> node.getAttribute("identity") Then
                Err.Raise 5, , "Stored Word event identity differs from its fields."
            End If
        End If
    Next node
End Sub

Private Function private_Uid(ByVal prefix As String) As String
    Dim number As Long

    number = VBA.CLng(m_root.getAttribute("nextUid"))
    If number < 1 Then
        Err.Raise 5, , "Invalid Word UID sequence."
    End If
    m_root.setAttribute "nextUid", VBA.CStr(number + 1)
    private_Uid = prefix & VBA.CStr(number)
End Function

Private Function private_Node(ByVal name As String) As Object
    Set private_Node = m_document.createNode(1, name, STORE_NS)
End Function

Private Sub private_AddText( _
    ByVal parent As Object, _
    ByVal name As String, _
    ByVal text As String _
)
    Dim node As Object

    Set node = private_Node(name)
    node.Text = text
    parent.appendChild node
End Sub

Private Function private_Value( _
    ByVal values As Object, _
    ByVal name As String _
) As String
    private_Value = ex_PEB_Export.fn_Text(ex_PEB_Export.fn_Required(values, name))
End Function

Private Sub private_Fail(ByVal code As String)
    Err.Raise 5, , ex_PEB_Export.fn_Config("text.WordStore" & code)
End Sub