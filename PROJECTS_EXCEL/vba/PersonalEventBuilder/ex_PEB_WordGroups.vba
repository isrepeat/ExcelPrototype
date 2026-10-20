Attribute VB_Name = "ex_PEB_WordGroups"
Option Explicit

' --------------------------------------
' namespace API {
' --------------------------------------
Public Sub fn_Validate(ByVal template As Object)
    Dim node As Object
    Dim group As Object
    Dim fields As String
    Dim bodyCount As Long
    Dim bodySeen As Boolean

    private_CheckAttributes template, "name|hash|identityBy|orderBy"
    fields = private_Attribute(template, "name")
    bodyCount = template.selectNodes("p:text | p:group").Length
    If bodyCount <> 1 Then
        Err.Raise 5, , "Word template requires exactly one text or group body."
    End If
    For Each node In template.childNodes
        If node.nodeType = 1 Then
            If node.namespaceURI <> "urn:excelprototype:profiles" Then
                Err.Raise 5, , "Unknown Word template namespace."
            End If
            Select Case node.baseName
                Case "derived"
                    If bodySeen Then
                        Err.Raise 5, , "Derived Word fields must precede the template body."
                    End If
                    private_CheckAttributes node, "name"
                    fields = private_Attribute(node, "name")
                    private_ValidateTextBlock node
                Case "text", "group"
                    bodySeen = True
                Case Else
                    Err.Raise 5, , "Unknown Word template element: " & node.nodeName
            End Select
        End If
    Next node
    Set group = template.selectSingleNode("p:group")
    If group Is Nothing Then
        If Not template.getAttributeNode("identityBy") Is Nothing Or Not template.getAttributeNode("orderBy") Is Nothing Then
            Err.Raise 5, , "Word state keys require a group body."
        End If
        private_ValidateText template.selectSingleNode("p:text")
    Else
        fields = private_Attribute(template, "identityBy")
        fields = private_Attribute(template, "orderBy")
        private_ValidateGroup group
    End If
End Sub

Public Function fn_Render( _
    ByVal template As Object, _
    ByVal events As Collection, _
    ByVal sharedTemplates As Object _
) As String
    Dim renderer As obj_PEB_DslRenderer
    Dim item As Object
    Dim derived As Object
    Dim items As Collection
    Dim copy As Object
    Dim key As Variant
    Dim group As Object
    Dim result As String
    Dim errorText As String

    On Error GoTo EH
    Set renderer = New obj_PEB_DslRenderer
    If Not renderer.Initialize(sharedTemplates) Then
        Err.Raise 5, , "Cannot initialize grouped Word renderer."
    End If
    Set items = New Collection
    For Each item In events
        Set copy = VBA.CreateObject("Scripting.Dictionary")
        For Each key In item.Keys
            If Not VBA.IsObject(item(key)) Then
                copy.Add key, item(key)
            Else
                Set copy(key) = item(key)
            End If
        Next key
        For Each derived In template.selectNodes("p:derived")
            key = private_Attribute(derived, "name")
            If copy.Exists(key) Then
                Err.Raise 5, , "Derived Word field collides with an input: " & key
            End If
            copy.Add key, private_RenderText(derived, copy, renderer)
        Next derived
        items.Add copy
    Next item
    Set group = template.selectSingleNode("p:group")
    If group Is Nothing Then
        If items.Count <> 1 Then
            Err.Raise 5, , "A text-only Word template requires exactly one event."
        End If
        result = private_RenderText(template, items(1), renderer)
    Else
        result = private_RenderGroup(group, items, renderer)
    End If
    renderer.Dispose
    fn_Render = result
    Exit Function
EH:
    errorText = VBA.Err.Description
    If Not renderer Is Nothing Then
        renderer.Dispose
    End If
    Err.Raise 5, , "Word grouping: " & errorText
End Function

Public Function fn_Key( _
    ByVal values As Object, _
    ByVal fields As String _
) As String
    Dim field As Variant

    Dim value As Variant
    Dim text As String
    Dim result As String

    For Each field In VBA.Split(fields, ";")
        field = VBA.Trim$(VBA.CStr(field))
        If VBA.Len(field) = 0 Then
            Err.Raise 5, , "Empty Word grouping key field."
        End If
        value = ex_PEB_Export.fn_Required(values, VBA.CStr(field))
        If VBA.IsObject(value) Or VBA.IsArray(value) Or VBA.IsNull(value) Or VBA.IsError(value) Then
            Err.Raise 5, , "Word grouping key must be scalar: " & field
        End If
        text = ex_PEB_Export.fn_Text(value)
        If VBA.Len(text) = 0 Then
            Err.Raise 5, , "Empty Word grouping key: " & field
        End If
        ' Length-prefixed segments prevent collisions when values contain separators.
        result = result & VBA.CStr(VBA.Len(text)) & ":" & text
    Next field
    fn_Key = result
End Function
' --------------------------------------
' } // namespace API
' --------------------------------------

Private Function private_RenderGroup( _
    ByVal group As Object, _
    ByVal items As Collection, _
    ByVal renderer As obj_PEB_DslRenderer _
) As String
    Dim buckets As Object
    Dim item As Object
    Dim members As Collection
    Dim key As Variant
    Dim context As Object
    Dim aggregate As Object
    Dim branch As Object
    Dim child As Object
    Dim result As String
    Dim field As Variant
    Dim itemContext As Object
    Dim itemIndex As Long

    Set buckets = VBA.CreateObject("Scripting.Dictionary")
    For Each item In items
        key = fn_Key(item, private_Attribute(group, "by"))
        If Not buckets.Exists(key) Then
            Set members = New Collection
            buckets.Add key, members
        End If
        buckets(key).Add item
    Next item
    For Each key In buckets.Keys
        Set members = buckets(key)
        Set context = VBA.CreateObject("Scripting.Dictionary")
        For Each field In members(1).Keys
            context.Add field, members(1)(field)
        Next field
        For Each aggregate In group.selectNodes("p:aggregate")
            If private_Attribute(aggregate, "operation") <> "numberRanges" Then
                Err.Raise 5, , "Unsupported Word aggregate operation."
            End If
            field = private_Attribute(aggregate, "name")
            If context.Exists(field) Then
                Err.Raise 5, , "Word aggregate collides with input: " & field
            End If
            context.Add field, private_NumberRanges(members, aggregate)
        Next aggregate
        If members.Count = 1 Then
            context("__GroupFirst") = True
            Set branch = group.selectSingleNode("p:single")
        Else
            Set branch = group.selectSingleNode("p:multiple")
        End If
        If branch Is Nothing Then
            Err.Raise 5, , "Word group requires single and multiple branches."
        End If
        For Each child In branch.childNodes
            If child.nodeType = 1 Then
                Select Case child.baseName
                    Case "header", "footer", "text"
                        If child.baseName = "text" Then
                            private_Append result, private_RenderText(branch, context, renderer)
                        Else
                            private_RequireCommon child, members
                            private_Append result, private_RenderText(child, context, renderer)
                        End If
                    Case "item"
                        itemIndex = 0
                        For Each item In members
                            itemIndex = itemIndex + 1
                            Set itemContext = VBA.CreateObject("Scripting.Dictionary")
                            For Each field In item.Keys
                                itemContext.Add field, item(field)
                            Next field
                            itemContext("__GroupFirst") = (itemIndex = 1)
                            private_Append result, private_RenderText(child, itemContext, renderer)
                        Next item
                    Case "group"
                        private_Append result, private_RenderGroup(child, members, renderer)
                    Case Else
                        Err.Raise 5, , "Unsupported Word group element: " & child.nodeName
                End Select
            End If
        Next child
    Next key
    private_RenderGroup = result
End Function

Private Sub private_ValidateGroup(ByVal group As Object)
    Dim branch As Object
    Dim node As Object
    Dim value As String
    Dim names As Object

    private_CheckAttributes group, "by"
    value = private_Attribute(group, "by")
    If group.selectNodes("p:single").Length <> 1 Or group.selectNodes("p:multiple").Length <> 1 Then
        Err.Raise 5, , "Word group requires exactly one single and multiple branch."
    End If
    private_CheckAttributes group.selectSingleNode("p:single"), VBA.vbNullString
    private_ValidateTextBlock group.selectSingleNode("p:single")
    Set names = VBA.CreateObject("Scripting.Dictionary")
    For Each node In group.selectNodes("p:aggregate")
        private_CheckAttributes node, "name|field|operation|pattern|separator|rangeSeparator"
        value = private_Attribute(node, "name")
        If names.Exists(value) Then
            Err.Raise 5, , "Duplicate Word aggregate: " & value
        End If
        names.Add value, True
        value = private_Attribute(node, "field")
        value = private_Attribute(node, "pattern")
        value = private_Attribute(node, "separator")
        value = private_Attribute(node, "rangeSeparator")
        If private_Attribute(node, "operation") <> "numberRanges" Then
            Err.Raise 5, , "Unknown Word aggregate operation."
        End If
    Next node
    For Each node In group.childNodes
        If node.nodeType = 1 Then
            If node.namespaceURI <> "urn:excelprototype:profiles" Then
                Err.Raise 5, , "Unknown Word group namespace."
            End If
            If node.baseName <> "aggregate" And node.baseName <> "single" And node.baseName <> "multiple" Then
                Err.Raise 5, , "Unknown Word group element: " & node.nodeName
            End If
        End If
    Next node
    Set branch = group.selectSingleNode("p:multiple")
    private_CheckAttributes branch, VBA.vbNullString
    If branch.selectNodes("p:item | p:group").Length <> 1 Then
        Err.Raise 5, , "Word multiple branch requires one item or nested group."
    End If
    For Each node In branch.childNodes
        If node.nodeType = 1 Then
            If node.namespaceURI <> "urn:excelprototype:profiles" Then
                Err.Raise 5, , "Unknown Word multiple namespace."
            End If
            Select Case node.baseName
                Case "header", "footer"
                    private_CheckAttributes node, "common"
                    value = private_Attribute(node, "common")
                    If branch.selectNodes("p:" & node.baseName).Length <> 1 Then
                        Err.Raise 5, , "Duplicate shared Word group block."
                    End If
                    private_ValidateTextBlock node
                Case "item"
                    private_CheckAttributes node, VBA.vbNullString
                    private_ValidateTextBlock node
                Case "group"
                    private_ValidateGroup node
                Case Else
                    Err.Raise 5, , "Unknown Word multiple element: " & node.nodeName
            End Select
        End If
    Next node
End Sub

Private Sub private_CheckAttributes( _
    ByVal node As Object, _
    ByVal allowed As String _
)
    Dim xmlAttribute As Object

    For Each xmlAttribute In node.Attributes
        If xmlAttribute.namespaceURI <> "http://www.w3.org/2000/xmlns/" Then
            If VBA.InStr(1, "|" & allowed & "|", "|" & xmlAttribute.nodeName & "|", VBA.vbBinaryCompare) = 0 Then
                Err.Raise 5, , "Unknown Word attribute: " & node.nodeName & "/@" & xmlAttribute.nodeName
            End If
        End If
    Next xmlAttribute
End Sub

Private Sub private_ValidateTextBlock(ByVal node As Object)
    Dim child As Object

    If node.selectNodes("p:text").Length <> 1 Then
        Err.Raise 5, , "Word block requires exactly one text element."
    End If
    For Each child In node.childNodes
        If child.nodeType = 1 Then
            If child.namespaceURI <> "urn:excelprototype:profiles" Or child.baseName <> "text" Then
                Err.Raise 5, , "Unexpected element in Word text block: " & child.nodeName
            End If
            private_ValidateText child
        End If
    Next child
End Sub

Private Sub private_ValidateText(ByVal node As Object)
    Dim child As Object

    private_CheckAttributes node, VBA.vbNullString
    For Each child In node.childNodes
        If child.nodeType = 1 Then
            Err.Raise 5, , "Word text must contain DSL text, not XML elements."
        End If
    Next child
End Sub

Private Sub private_RequireCommon( _
    ByVal node As Object, _
    ByVal members As Collection _
)
    Dim field As Variant
    Dim item As Object
    Dim expected As String

    ' Shared text declares the scalar inputs that must agree for all group members.
    For Each field In VBA.Split(private_Attribute(node, "common"), ";")
        field = VBA.Trim$(VBA.CStr(field))
        expected = ex_PEB_Export.fn_Text(ex_PEB_Export.fn_Required(members(1), VBA.CStr(field)))
        For Each item In members
            If ex_PEB_Export.fn_Text(ex_PEB_Export.fn_Required(item, VBA.CStr(field))) <> expected Then
                Err.Raise 5, , "Conflicting shared Word group field: " & field
            End If
        Next item
    Next field
End Sub

Private Function private_NumberRanges( _
    ByVal members As Collection, _
    ByVal aggregate As Object _
) As String
    Dim rx As Object
    Dim match As Object
    Dim matches As Object
    Dim item As Object
    Dim series As Object
    Dim seen As Object
    Dim values As Collection
    Dim key As Variant
    Dim text As String
    Dim prefix As String
    Dim suffix As String
    Dim digits As String
    Dim numbers() As Long
    Dim labels As Object
    Dim index As Long
    Dim other As Long
    Dim swap As Long
    Dim last As Long
    Dim result As String
    Dim entry As Variant

    Set rx = VBA.CreateObject("VBScript.RegExp")
    rx.Pattern = private_Attribute(aggregate, "pattern")
    Set series = VBA.CreateObject("Scripting.Dictionary")
    Set seen = VBA.CreateObject("Scripting.Dictionary")
    For Each item In members
        text = ex_PEB_Export.fn_Text(ex_PEB_Export.fn_Required(item, private_Attribute(aggregate, "field")))
        If Not rx.Test(text) Then
            Err.Raise 5, , "Document number does not match the declared range pattern: " & text
        End If
        If seen.Exists(text) Then
            Err.Raise 5, , "Duplicate grouped document number: " & text
        End If
        seen.Add text, True
        Set matches = rx.Execute(text)
        Set match = matches(0)
        If match.FirstIndex <> 0 Or match.Length <> VBA.Len(text) Then
            Err.Raise 5, , "Document range pattern must match the entire number."
        End If
        If match.SubMatches.Count <> 3 Then
            Err.Raise 5, , "Range pattern requires prefix, integer and suffix captures."
        End If
        prefix = match.SubMatches(0)
        digits = match.SubMatches(1)
        suffix = match.SubMatches(2)
        If Not VBA.IsNumeric(digits) Then
            Err.Raise 5, , "Invalid document range integer."
        End If
        key = VBA.Len(prefix) & ":" & prefix & VBA.Len(suffix) & ":" & suffix
        If VBA.Len(digits) > 1 And VBA.Left$(digits, 1) = "0" Then
            key = key & ":width=" & VBA.Len(digits)
        End If
        If Not series.Exists(key) Then
            Set values = New Collection
            series.Add key, values
        End If
        series(key).Add VBA.Array(VBA.CLng(digits), text)
    Next item
    For Each key In series.Keys
        Set values = series(key)
        Set labels = VBA.CreateObject("Scripting.Dictionary")
        ReDim numbers(1 To values.Count)
        For index = 1 To values.Count
            entry = values(index)
            numbers(index) = entry(0)
            If labels.Exists(numbers(index)) Then
                Err.Raise 5, , "Ambiguous document number formatting."
            End If
            labels.Add numbers(index), entry(1)
        Next index
        For index = 2 To values.Count
            other = index
            Do While other > 1
                If numbers(other - 1) <= numbers(other) Then
                    Exit Do
                End If
                swap = numbers(other - 1)
                numbers(other - 1) = numbers(other)
                numbers(other) = swap
                other = other - 1
            Loop
        Next index
        index = 1
        Do While index <= values.Count
            last = index
            Do While last < values.Count
                If VBA.CDbl(numbers(last + 1)) <> VBA.CDbl(numbers(last)) + 1# Then
                    Exit Do
                End If
                last = last + 1
            Loop
            If VBA.Len(result) > 0 Then
                result = result & private_Attribute(aggregate, "separator")
            End If
            result = result & labels(numbers(index))
            If last > index Then
                result = result & private_Attribute(aggregate, "rangeSeparator") & labels(numbers(last))
            End If
            index = last + 1
        Loop
    Next key
    private_NumberRanges = result
End Function

Private Function private_RenderText( _
    ByVal node As Object, _
    ByVal values As Object, _
    ByVal renderer As obj_PEB_DslRenderer _
) As String
    Dim text As Object
    Dim result As String
    Dim diagnostic As String

    Set text = node.selectNodes("p:text")
    If text.Length <> 1 Then
        Err.Raise 5, , "Grouped Word block requires one text element."
    End If
    If Not renderer.TryRender(text(0).Text, values, result, diagnostic) Then
        Err.Raise 5, , diagnostic
    End If
    private_RenderText = result
End Function

Private Function private_Attribute( _
    ByVal node As Object, _
    ByVal name As String _
) As String
    Dim value As Object

    Set value = node.getAttributeNode(name)
    If value Is Nothing Then
        Err.Raise 5, , "Missing Word grouping attribute: " & name
    End If
    ' MSXML .Text trims attribute whitespace, including significant separator spaces.
    private_Attribute = VBA.CStr(value.nodeValue)
    If VBA.Len(private_Attribute) = 0 Then
        Err.Raise 5, , "Empty Word grouping attribute: " & name
    End If
End Function

Private Sub private_Append( _
    ByRef result As String, _
    ByVal text As String _
)
    If VBA.Len(text) = 0 Then
        Exit Sub
    End If
    If VBA.Len(result) > 0 Then
        result = result & VBA.vbCrLf & VBA.vbCrLf
    End If
    result = result & text
End Sub