Attribute VB_Name = "ex_PEB_EventRoutes"
Option Explicit

Private m_events As Object
Private m_statuses As Object
Private m_bindingField As String

' --------------------------------------
' namespace Lifecycle {
' --------------------------------------
Public Sub fn_Module_Dispose()
    Set m_events = Nothing
    Set m_statuses = Nothing
    ex_PEB_MovementTemplates.fn_Module_Dispose
    ex_PEB_MovementRegistry.fn_Module_Dispose
    m_bindingField = VBA.vbNullString
End Sub
' --------------------------------------
' } // namespace Lifecycle
' --------------------------------------

' --------------------------------------
' namespace API {
' --------------------------------------
Public Function fn_Ids() As Variant
    private_EnsureLoaded
    fn_Ids = m_events.Keys
End Function

Public Function fn_VisibleIds() As Variant
    Dim eventId As Variant
    Dim result As Collection
    Dim values() As String
    Dim index As Long

    private_EnsureLoaded
    Set result = New Collection
    For Each eventId In m_events.Keys
        If m_events(eventId)("visible") = "true" Then
            result.Add eventId
        End If
    Next eventId
    If result.Count = 0 Then
        Err.Raise 5, , "Event routes XML has no visible UI events."
    End If
    ReDim values(0 To result.Count - 1)
    For index = 1 To result.Count
        values(index - 1) = result(index)
    Next index
    fn_VisibleIds = values
End Function

Public Function fn_Value( _
    ByVal eventId As String, _
    ByVal key As String _
) As String
    private_EnsureLoaded
    If Not m_events.Exists(eventId) Then
        Err.Raise 5, , "Unknown XML business event: " & eventId
    End If
    If Not m_events(eventId).Exists(key) Then
        Err.Raise 5, , "Unknown event route property: " & eventId & "/" & key
    End If
    fn_Value = m_events(eventId)(key)
End Function

Public Function fn_HasRoute( _
    ByVal eventId As String, _
    ByVal exporter As String _
) As Boolean
    private_EnsureLoaded
    If Not m_events.Exists(eventId) Then
        Err.Raise 5, , "Unknown XML business event: " & eventId
    End If
    If exporter <> "Word" And exporter <> "Movement" Then
        Err.Raise 5, , "Unknown routed exporter: " & exporter
    End If
    fn_HasRoute = m_events(eventId).Exists(VBA.LCase$(exporter) & ".template")
End Function

Public Function fn_Apply( _
    ByVal values As Object, _
    ByVal exporter As String _
) As String
    Dim eventId As String
    Dim statusId As String
    Dim destination As String

    private_EnsureLoaded
    eventId = ex_PEB_Export.fn_Text(ex_PEB_Export.fn_Required(values, m_bindingField))
    values("EventId") = eventId
    values("ModeDirection") = fn_Value(eventId, "direction")
    values("ModeForm") = fn_Value(eventId, "form")
    values("Operation") = values("ModeDirection")
    If values("Operation") = "StatusChange" Then values("Operation") = "Transition"
    If fn_HasRoute(eventId, "Movement") Then
        values("SourceStatus") = fn_Value(eventId, "movement.sourceStatus")
        values("TargetStatus") = fn_Value(eventId, "movement.targetStatus")
        values("Operation") = fn_Value(eventId, "movement.operation")
        values("MovementTemplateId") = fn_Value(eventId, "movement.template")
        values("MovementPlanId") = fn_Value(eventId, "movement.profile")
        values("MovementOperation") = fn_Value(eventId, "movement.operation")
        statusId = fn_Value(eventId, "movement.targetStatus")
        If VBA.Len(statusId) = 0 Then
            statusId = fn_Value(eventId, "movement.sourceStatus")
        End If
        values("EventName") = m_statuses(statusId)
        statusId = fn_Value(eventId, "movement.sourceStatus")
        values("SourceEventName") = VBA.vbNullString
        If VBA.Len(statusId) > 0 Then
            values("SourceEventName") = m_statuses(statusId)
        End If
        values("SourceDestination") = fn_Value(eventId, "movement.sourceDestination")
    Else
        values("SourceStatus") = VBA.vbNullString
        values("TargetStatus") = VBA.vbNullString
        If values("Operation") <> "Departure" Then
            values("SourceStatus") = ex_PEB_Export.fn_Config("eventContext." & eventId & ".sourceStatus")
        End If
        If values("Operation") <> "Arrival" Then
            values("TargetStatus") = ex_PEB_Export.fn_Config("eventContext." & eventId & ".targetStatus")
        End If
        values("EventName") = fn_Value(eventId, "label")
        values("SourceEventName") = VBA.vbNullString
        values("SourceDestination") = VBA.vbNullString
    End If
    values("WordPolicy") = fn_Value(eventId, "word.policy")
    values("RoutedWordTemplate") = VBA.vbNullString
    If fn_HasRoute(eventId, "Word") Then
        values("RoutedWordTemplate") = fn_Value(eventId, "word.template")
    End If
    If exporter = "Word" Then
        fn_Apply = eventId
        If fn_HasRoute(eventId, "Word") Then
            fn_Apply = fn_Value(eventId, "word.template")
        End If
    ElseIf exporter = "Movement" Then
        If Not fn_HasRoute(eventId, "Movement") Then
            Err.Raise 5, , "Movement route is not configured: " & eventId
        End If
        destination = fn_Value(eventId, "movement.destinationValue")
        If VBA.Len(destination) = 0 Then
            destination = ex_PEB_Export.fn_Text(ex_PEB_Export.fn_Required(values, fn_Value(eventId, "movement.destinationField")))
        End If
        values("Destination") = destination
        fn_Apply = eventId
    Else
        Err.Raise 5, , "Unknown routed exporter: " & exporter
    End If
End Function
' --------------------------------------
' } // namespace API
' --------------------------------------

' --------------------------------------
' namespace Private {
' --------------------------------------
Private Sub private_EnsureLoaded()
    Dim document As Object
    Dim templates As Object
    Dim root As Object
    Dim node As Object
    Dim child As Object
    Dim route As Object
    Dim knownTemplates As Object
    Dim movementPlan As Object
    Dim textTemplate As Object
    Dim identifier As String
    Dim key As Variant
    Dim profile As String
    Dim count As Long
    Dim message As String

    If Not m_events Is Nothing Then
        Exit Sub
    End If
    On Error GoTo EH
    Set document = private_LoadXml(ex_PEB_Export.fn_Path(ex_PEB_Export.fn_Config("events.path")))
    Set root = document.selectSingleNode("/p:eventRoutes")
    If root Is Nothing Then
        Err.Raise 5, , "Expected eventRoutes in urn:excelprototype:profiles."
    End If
    private_CheckAttributes root, "version|bindingField"
    If private_Attribute(root, "version") <> "1" Then
        Err.Raise 5, , "Unsupported event routes XML version."
    End If
    m_bindingField = private_Attribute(root, "bindingField")
    private_ValidateFields m_bindingField
    private_RequireChildren root, "events"
    Set m_statuses = private_Dictionary()
    If document.selectNodes("/p:eventRoutes/p:events/p:event/p:movement").Length > 0 Then
        Set m_statuses = ex_PEB_MovementRegistry.fn_Statuses()
    End If
    Set knownTemplates = private_Dictionary()
    If document.selectNodes("/p:eventRoutes/p:events/p:event/p:word").Length > 0 Then
        Set templates = private_LoadXml(ex_PEB_Export.fn_Path(ex_PEB_Export.fn_Config("export.Word.templates.path")))
        For Each node In templates.documentElement.childNodes
            If node.nodeType = 1 Then
                If node.namespaceURI <> "urn:excelprototype:profiles" Or VBA.InStr(1, "|documentFilepath|sharedTemplates|template|", "|" & node.baseName & "|", VBA.vbBinaryCompare) = 0 Then
                    Err.Raise 5, , "Unknown Word template section: " & node.nodeName
                End If
            End If
        Next node
        For Each node In templates.selectNodes("/p:wordResultTemplates/p:template")
            private_Add knownTemplates, private_Attribute(node, "name"), True
        Next node
    End If
    Set m_events = private_Dictionary()
    For Each node In root.selectNodes("p:events/p:event")
        private_CheckAttributes node, "id|caption|label|direction|form|visible"
        identifier = private_Attribute(node, "id")
        Set route = private_Dictionary()
        For Each key In VBA.Array("caption", "label", "direction", "form", "visible")
            route.Add key, private_Attribute(node, VBA.CStr(key))
        Next key
        If route("direction") <> "Arrival" And route("direction") <> "Departure" And route("direction") <> "StatusChange" Then
            Err.Raise 5, , "Invalid event direction: " & identifier
        End If
        If route("form") <> "Leave" And route("form") <> "Treatment" Then
            Err.Raise 5, , "Invalid event form: " & identifier
        End If
        If route("visible") <> "true" And route("visible") <> "false" Then
            Err.Raise 5, , "Invalid event visibility: " & identifier
        End If
        private_ValidateRoutes node
        route.Add "word.policy", "Omit"
        Set child = node.selectSingleNode("p:movement")
        If Not child Is Nothing Then
            private_CheckAttributes child, "template"
            route.Add "movement.template", private_Attribute(child, "template")
            Set textTemplate = ex_PEB_MovementTemplates.fn_Template(route("movement.template"))
            Set movementPlan = ex_PEB_MovementRegistry.fn_Event(identifier)
            For Each key In VBA.Array("operation", "profile", "destinationField", "required", "sourceStatus", "targetStatus", "destinationValue", "sourceDestination")
                route.Add "movement." & key, movementPlan(key)
            Next key
        End If
        Set child = node.selectSingleNode("p:word")
        If Not child Is Nothing Then
            private_CheckAttributes child, "template"
            route.Add "word.template", private_Attribute(child, "template")
            If Not knownTemplates.Exists(route("word.template")) Then
                Err.Raise 5, , "Unknown Word template: " & identifier & "/" & route("word.template")
            End If
            route("word.policy") = "Render"
            profile = "export.Word." & route("word.template")
            private_ValidateFields ex_PEB_Export.fn_Config(profile & ".required")
            private_ValidateFields ex_PEB_Export.fn_Config(profile & ".previewRequired")
            count = ex_PEB_Export.fn_Count(profile & ".lookups.count")
        End If
        private_Add m_events, identifier, route
    Next node
    If m_events.Count = 0 Then
        Err.Raise 5, , "Event routes XML contains no events."
    End If
    Exit Sub
EH:
    message = VBA.Err.Description
    fn_Module_Dispose
    Err.Raise 5, , "Event routes XML: " & message
End Sub

Private Sub private_ValidateRoutes(ByVal node As Object)
    Dim child As Object
    Dim count As Long

    For Each child In node.childNodes
        If child.nodeType = 1 Then
            If child.namespaceURI <> "urn:excelprototype:profiles" Or (child.baseName <> "movement" And child.baseName <> "word") Then
                Err.Raise 5, , "Unknown exporter route: " & child.nodeName
            End If
            count = count + 1
            If child.selectNodes("*").Length > 0 Then
                Err.Raise 5, , "Exporter route must not contain child elements: " & child.nodeName
            End If
        End If
    Next child
    If count = 0 Then
        Err.Raise 5, , "Event must declare at least one exporter route: " & private_Attribute(node, "id")
    End If
    If node.selectNodes("p:movement").Length > 1 Or node.selectNodes("p:word").Length > 1 Then
        Err.Raise 5, , "Duplicate exporter route: " & private_Attribute(node, "id")
    End If
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

Private Sub private_ValidateFields(ByVal fields As String)
    Dim field As Variant
    Dim known As String

    known = "|" & ex_PEB_Export.fn_Config("export.fields") & "|"
    For Each field In VBA.Split(fields, "|")
        If VBA.Len(field) = 0 Or VBA.InStr(1, known, "|" & field & "|", VBA.vbBinaryCompare) = 0 Then
            Err.Raise 5, , "Unknown event input field: " & field
        End If
    Next field
End Sub
' --------------------------------------
' } // namespace Private
' --------------------------------------