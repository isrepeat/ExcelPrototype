Attribute VB_Name = "ex_PEB_WordDocument"
Option Explicit

Private Const STORE_NS As String = "urn:excelprototype:word-events:v1"

' --------------------------------------
' namespace API {
' --------------------------------------
Public Function fn_Events( _
    ByVal document As Object, _
    ByVal values As Object _
) As Collection
    Dim store As obj_PEB_WordEventStore
    Dim parts As Object
    Dim section As Object
    Dim item As Object
    Dim result As New Collection
    Dim orderKey As String
    Dim diagnostic As String
    Dim errorNumber As Long
    Dim errorText As String

    On Error GoTo EH
    Set parts = document.CustomXMLParts.SelectByNamespace(STORE_NS)
    If parts.Count <> 1 Then
        Err.Raise 5, , ex_PEB_Export.fn_Config("text.WordStoreMigration")
    End If
    orderKey = ex_PEB_WordGroups.fn_Key(values, "AcceptedOrderNumber; OrderDateShort")
    Set store = New obj_PEB_WordEventStore
    If Not store.Initialize(parts(1).XML, orderKey, diagnostic) Then
        Err.Raise 5, , diagnostic
    End If
    For Each section In store.Sections
        For Each item In store.Events(section)
            item.Add "WordAnchor", section.getAttribute("anchor")
            result.Add item
        Next item
    Next section
    Set fn_Events = result
Cleanup:
    If Not store Is Nothing Then
        store.Dispose
    End If
    If errorNumber <> 0 Then
        On Error GoTo 0
        Err.Raise errorNumber, , errorText
    End If
    Exit Function
EH:
    errorNumber = VBA.Err.Number
    errorText = VBA.Err.Description
    Resume Cleanup
End Function

Public Sub fn_Change( _
    ByVal document As Object, _
    ByVal values As Object, _
    ByVal operation As String, _
    ByRef uid As String, _
    ByRef resultCode As String _
)
    Dim store As obj_PEB_WordEventStore
    Dim verification As obj_PEB_WordEventStore
    Dim templates As obj_PEB_WordTemplates
    Dim anchors As obj_PEB_WordAnchors
    Dim parts As Object
    Dim part As Object
    Dim replacement As Object
    Dim section As Object
    Dim targetSection As Object
    Dim template As Object
    Dim target As Object
    Dim context As Object
    Dim events As Collection
    Dim keys As Collection
    Dim item As Object
    Dim rendered As String
    Dim text As String
    Dim xml As String
    Dim diagnostic As String
    Dim orderKey As String
    Dim sharedPlan As String
    Dim key As Variant
    Dim errorNumber As Long
    Dim errorText As String

    On Error GoTo EH
    resultCode = VBA.vbNullString
    Set store = New obj_PEB_WordEventStore
    Set templates = New obj_PEB_WordTemplates
    If Not templates.Initialize(diagnostic) Then
        Err.Raise 5, , diagnostic
    End If
    Set parts = document.CustomXMLParts.SelectByNamespace(STORE_NS)
    If parts.Count > 1 Then
        Err.Raise 5, , "Duplicate Word event stores."
    End If
    If parts.Count = 1 Then
        Set part = parts(1)
        xml = part.XML
    End If
    orderKey = ex_PEB_WordGroups.fn_Key(values, "AcceptedOrderNumber; OrderDateShort")
    If Not store.Initialize(xml, orderKey, diagnostic) Then
        Err.Raise 5, , diagnostic
    End If
    For Each key In templates.SharedTemplates.Keys
        text = templates.SharedTemplates(key)
        sharedPlan = sharedPlan & VBA.Len(key) & ":" & key & VBA.Len(text) & ":" & text
    Next key
    If operation <> "DELETE" Then
        Set template = templates.Definition(VBA.CStr(values("TemplateName")))
        Set targetSection = store.BindSection( _
            ex_PEB_Export.fn_Config("export.Word.textAnchors." & VBA.CStr(values("EventId"))), _
            VBA.CStr(values("TemplateName")), _
            VBA.CStr(values("EventId")), _
            template.XML, _
            sharedPlan _
        )
        Set keys = private_ContextKeys(template, templates.SharedTemplates)
    End If

    ' Reject duplicate inserts in memory before scanning Word anchor ranges.
    If operation = "INSERT" Then
        store.Change operation, targetSection, values, keys, uid
    End If

    ' Verify every managed section before publishing changes to Word.
    Set anchors = New obj_PEB_WordAnchors
    If Not anchors.Initialize(document, diagnostic) Then
        Err.Raise 5, , diagnostic
    End If
    For Each section In store.Sections
        Set template = templates.Definition(section.getAttribute("template"))
        If private_Normalize(template.XML) <> store.Text(section, "plan") Or _
                private_Normalize(sharedPlan) <> store.Text(section, "shared") Then
            Err.Raise 5, , ex_PEB_Export.fn_Config("text.WordStorePlanChanged")
        End If
        Set context = VBA.CreateObject("Scripting.Dictionary")
        context.Add "EventId", section.getAttribute("eventId")
        If ex_PEB_Export.fn_Config("export.Word.textAnchors." & section.getAttribute("eventId")) <> section.getAttribute("anchor") Then
            Err.Raise 5, , "Stored Word anchor mapping changed."
        End If
        If Not anchors.TryGetInsertedRange(document, context, target, diagnostic) Then
            Err.Raise 5, , diagnostic
        End If
        If section.selectNodes("s:rendered").Length = 0 Then
            If VBA.Len(VBA.Trim$(VBA.Replace$(private_Normalize(target.Text), VBA.vbLf, " "))) > 0 Then
                Err.Raise 5, , ex_PEB_Export.fn_Config("text.WordStoreMigration")
            End If
        ElseIf private_Normalize(target.Text) <> store.Text(section, "rendered") Then
            Err.Raise 5, , ex_PEB_Export.fn_Config("text.WordStoreManualEdit")
        End If
    Next section
    If operation <> "INSERT" Then
        store.Change operation, targetSection, values, keys, uid
    End If

    ' Rebuild both grouped and ordinary sections from their structured events.
    For Each section In store.Sections
        Set events = store.Events(section)
        Set template = templates.Definition(section.getAttribute("template"))
        rendered = VBA.vbNullString
        If events.Count > 0 Then
            If template.selectNodes("p:group").Length > 0 Then
                rendered = ex_PEB_WordGroups.fn_Render(template, events, templates.SharedTemplates)
            Else
                For Each item In events
                    If Not templates.TryRender(section.getAttribute("template"), item, text, diagnostic) Then
                        Err.Raise 5, , diagnostic
                    End If
                    If VBA.Len(rendered) > 0 Then
                        rendered = rendered & VBA.vbCrLf & VBA.vbCrLf
                    End If
                    rendered = rendered & text
                Next item
            End If
        End If
        anchors.Dispose
        Set anchors = New obj_PEB_WordAnchors
        If Not anchors.Initialize(document, diagnostic) Then
            Err.Raise 5, , diagnostic
        End If
        Set context = VBA.CreateObject("Scripting.Dictionary")
        context.Add "EventId", section.getAttribute("eventId")
        context.Add "WordText", rendered
        If Not anchors.TryInsert(document, context, diagnostic) Then
            Err.Raise 5, , diagnostic
        End If
        If Not anchors.TryVerify(document, context, diagnostic) Then
            Err.Raise 5, , diagnostic
        End If
        anchors.HideMarkers document, context
        store.SetRendered section, private_Normalize(rendered)
    Next section
    Set replacement = document.CustomXMLParts.Add(store.Xml)
    If replacement Is Nothing Then
        Err.Raise 5, , "Cannot save Word event store."
    End If
    Set verification = New obj_PEB_WordEventStore
    If Not verification.Initialize(replacement.XML, orderKey, diagnostic) Then
        Err.Raise 5, , diagnostic
    End If
    If verification.Xml <> store.Xml Then
        Err.Raise 5, , "Saved Word event store differs from the prepared state."
    End If
    If Not part Is Nothing Then
        part.Delete
    End If
    resultCode = operation
Cleanup:
    If Not anchors Is Nothing Then
        anchors.Dispose
    End If
    If Not templates Is Nothing Then
        templates.Dispose
    End If
    If Not store Is Nothing Then
        store.Dispose
    End If
    If Not verification Is Nothing Then
        verification.Dispose
    End If
    If errorNumber <> 0 Then
        On Error GoTo 0
        Err.Raise errorNumber, , errorText
    End If
    Exit Sub
EH:
    errorNumber = VBA.Err.Number
    errorText = VBA.Err.Description
    Resume Cleanup
End Sub
' --------------------------------------
' } // namespace API
' --------------------------------------

Private Function private_ContextKeys( _
    ByVal template As Object, _
    ByVal sharedTemplates As Object _
) As Collection
    Dim known As Object
    Dim included As Object
    Dim expression As Object
    Dim fieldExpression As Object
    Dim match As Object
    Dim result As New Collection
    Dim queue As New Collection
    Dim text As String
    Dim key As Variant
    Dim name As String

    Set known = VBA.CreateObject("Scripting.Dictionary")
    Set included = VBA.CreateObject("Scripting.Dictionary")
    Set expression = VBA.CreateObject("VBScript.RegExp")
    expression.Global = True
    expression.Pattern = "#include\s+([A-Za-z0-9_.]+)"
    queue.Add template.XML
    Do While queue.Count > 0
        name = queue(1)
        queue.Remove 1
        text = text & VBA.vbLf & name
        For Each match In expression.Execute(name)
            name = match.SubMatches(0)
            If Not included.Exists(name) Then
                If Not sharedTemplates.Exists(name) Then
                    Err.Raise 5, , "Missing shared Word block: " & name
                End If
                included.Add name, True
                queue.Add sharedTemplates(name)
            End If
        Next match
    Loop
    For Each key In VBA.Split("EventId|TemplateName|PersonId|PersonName|PositionCode|AcceptedOrderNumber|OrderDateShort|DateFromShort", "|")
        known.Add VBA.CStr(key), True
    Next key
    Set fieldExpression = VBA.CreateObject("VBScript.RegExp")
    ' Qualified fields belong to collection items, not to the root event context.
    fieldExpression.Global = True
    fieldExpression.Pattern = "[A-Za-z_][A-Za-z0-9_]*\.\[[A-Za-z_][A-Za-z0-9_]*\]"
    text = fieldExpression.Replace(text, VBA.vbNullString)
    For Each key In VBA.Split(ex_PEB_Export.fn_Config("export.Word.context.xmlKeys") & "|" & ex_PEB_Export.fn_Config("export.Word.context.requiredKeys"), "|")
        fieldExpression.Pattern = "(^|[^A-Za-z0-9_])" & VBA.CStr(key) & "([^A-Za-z0-9_]|$)"
        If fieldExpression.Test(text) Then
            known(VBA.CStr(key)) = True
        End If
    Next key
    For Each key In known.Keys
        result.Add VBA.CStr(key)
    Next key
    Set private_ContextKeys = result
End Function

Private Function private_Normalize(ByVal text As String) As String
    private_Normalize = VBA.Replace$(VBA.Replace$(text, VBA.vbCrLf, VBA.vbLf), VBA.vbCr, VBA.vbLf)
End Function