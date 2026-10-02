Attribute VB_Name = "ex_UiFormMarkup"
Option Explicit

' namespace API {
Public Function fn_Prepare(ByVal document As Object, ByRef outError As String) As Boolean
    Dim form As Object
    Dim field As Object
    Dim node As Object
    Dim editor As Object
    Dim label As Object
    Dim attributeNode As Object
    Dim kind As String
    Dim fieldName As String
    Dim fieldValue As String
    Dim position As String
    Dim formName As String

    outError = VBA.vbNullString
    For Each form In document.SelectNodes("//*[local-name()='form']")
        formName = private_Attribute(form, "name")
        If Len(formName) = 0 Or Len(private_Attribute(form, "source")) = 0 Then
            outError = "A form requires name and source."
            Exit Function
        End If
        kind = LCase$(private_Attribute(form, "orientation"))
        If kind <> "vertical" And kind <> "horizontal" Then
            outError = "Form '" & formName & "' requires orientation=vertical or horizontal."
            Exit Function
        End If
        For Each node In form.ChildNodes
            If node.NodeType = 1 Then
                kind = LCase$(CStr(node.baseName))
                If kind <> "field" And kind <> "control" And private_Attribute(node, "generatedField") <> "true" Then
                    outError = "Form '" & formName & "' only accepts field and control children."
                    Exit Function
                End If
            End If
        Next node
    Next form
    For Each field In document.SelectNodes("//*[local-name()='field']")
        Set form = field.parentNode
        If LCase$(CStr(form.baseName)) <> "form" Then
            outError = "A field must be a direct child of form."
            Exit Function
        End If
        fieldName = private_Attribute(field, "name")
        If Len(fieldName) = 0 Or Len(fieldName & "_label") > 31 Or Len(fieldName & "_input") > 31 Then
            outError = "A field requires a name of at most 25 characters: " & fieldName
            Exit Function
        End If
        If Len(private_Attribute(field, "label")) = 0 Then
            outError = "A field requires a label: " & fieldName
            Exit Function
        End If
        kind = LCase$(private_Attribute(field, "type"))
        If kind <> "text" And kind <> "select" And kind <> "checkbox" Then
            outError = "Unsupported field type: " & kind
            Exit Function
        End If
        position = private_Inherited(field, form, "labelPosition", "left")
        If position <> "left" And position <> "top" Then
            outError = "Unsupported labelPosition: " & position
            Exit Function
        End If
        Set node = document.createNode(1, "stackPanel", form.namespaceURI)
        node.setAttribute "generatedField", "true"
        node.setAttribute "orientation", IIf(position = "left", "horizontal", "vertical")
        Set label = document.createNode(1, "control", form.namespaceURI)
        label.setAttribute "type", "Label"
        label.setAttribute "name", fieldName & "_label"
        label.setAttribute "text", private_Attribute(field, "label")
        label.setAttribute "columnSpan", private_Inherited(field, form, "labelColumnSpan", "2")
        label.setAttribute "style", private_Inherited(field, form, "labelStyle", VBA.vbNullString)
        node.appendChild label
        Set editor = document.createNode(1, "control", form.namespaceURI)
        For Each attributeNode In field.Attributes
            editor.setAttribute attributeNode.nodeName, attributeNode.Text
        Next attributeNode
        editor.setAttribute "type", IIf(kind = "select", "Select", "Input")
        editor.setAttribute "inputType", kind
        editor.setAttribute "name", fieldName & "_input"
        editor.setAttribute "formName", private_Attribute(form, "name")
        editor.setAttribute "columnSpan", private_Inherited(field, form, "fieldWidth", "4")
        editor.setAttribute "rowSpan", private_Inherited(field, form, "height", "1")
        If Len(private_Attribute(field, "style")) = 0 Then
            editor.setAttribute "style", private_Inherited(field, form, "fieldStyle", VBA.vbNullString)
        End If
        editor.setAttribute "onChange", private_Inherited(field, form, "onChange", VBA.vbNullString)
        editor.setAttribute "readOnly", private_Inherited(field, form, "readOnly", "false")
        fieldValue = private_Attribute(field, "value")
        If Len(fieldValue) = 0 Then fieldValue = "{Binding Path=" & fieldName & "}"
        editor.setAttribute "value", fieldValue
        node.appendChild editor
        form.replaceChild node, field
    Next field
    fn_Prepare = True
End Function
' } // namespace API

Private Function private_Attribute(ByVal node As Object, ByVal name As String) As String
    Dim value As Variant
    value = node.getAttribute(name)
    If Not IsNull(value) And Not IsEmpty(value) Then private_Attribute = CStr(value)
End Function

Private Function private_Inherited(ByVal field As Object, ByVal form As Object, _
    ByVal name As String, ByVal defaultValue As String) As String
    Dim value As String
    value = private_Attribute(field, name)
    If Len(value) = 0 Then value = private_Attribute(form, name)
    If Len(value) = 0 Then value = defaultValue
    private_Inherited = value
End Function