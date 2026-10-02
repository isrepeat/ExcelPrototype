VERSION 1.0 CLASS
BEGIN
  MultiUse = -1
END
Attribute VB_Name = "obj_UiMarkupValidator"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private m_isInitialized As Boolean
Private m_isDisposed As Boolean

' //
' // Lifecycle
' //
Private Sub Class_Initialize()
End Sub

Private Sub Class_Terminate()
    Me.Dispose
End Sub

' //
' // API
' //
Public Function Initialize() As Boolean
    If m_isDisposed Or m_isInitialized Then
        Exit Function
    End If
    m_isInitialized = True
    Initialize = True
End Function
Public Sub Dispose()
    If m_isDisposed Then
        Exit Sub
    End If
    m_isDisposed = True
    m_isInitialized = False
End Sub

Public Function Validate(ByVal root As Object, ByVal errors As Collection) As Boolean
    Dim schema As obj_UiMarkupSchema
    Dim diagnostic As String
    Dim initialCount As Long

    initialCount = errors.Count
    If root Is Nothing Then
        private_AddError errors, "/", "", "Root element is required."
    ElseIf VBA.LCase$(VBA.CStr(root.baseName)) <> "page" Or VBA.CStr(root.namespaceURI) <> "urn:excelprototype:profiles" Then
        private_AddError errors, "/", VBA.CStr(root.baseName), "Root must be page."
    ElseIf Not ex_UiElementFactory.fn_TryGetSchema(VBA.CStr(root.baseName), schema, diagnostic, VBA.CStr(root.namespaceURI)) Then
        private_AddError errors, "/page", "", diagnostic
    Else
        private_ValidateNode root, schema, "/page", errors
    End If
    Validate = (errors.Count = initialCount)
End Function

' //
' // Private
' //
Private Sub private_AddError(ByVal errors As Collection, ByVal path As String, ByVal member As String, ByVal message As String)
    Dim diagnostic As New obj_UiMarkupDiagnostic

    diagnostic.Initialize path, member, message
    errors.Add diagnostic
End Sub

Private Sub private_ValidateNode(ByVal node As Object, ByVal schema As obj_UiMarkupSchema, ByVal path As String, ByVal errors As Collection)
    Dim alternative As Variant
    Dim candidate As Variant
    Dim found As Boolean
    Dim attributeNode As Object
    Dim child As Object
    Dim childSchema As obj_UiMarkupSchema
    Dim counts As Object
    Dim key As Variant
    Dim rule As Variant
    Dim value As String
    Dim namespaceUri As String
    Dim localName As String
    Dim tag As String
    Dim childPath As String
    Dim diagnostic As String

    For Each attributeNode In node.Attributes
        key = VBA.CStr(attributeNode.nodeName)
        If key <> "xmlns" And VBA.Left$(key, 6) <> "xmlns:" Then
            If Not schema.Attributes.Exists(key) Then
                private_AddError errors, path, VBA.CStr(key), "Unsupported attribute."
            Else
                rule = schema.Attributes(key)
                value = VBA.CStr(attributeNode.Text)
                If Not private_AttributeValid(value, rule) Then private_AddError errors, path, VBA.CStr(key), "Invalid attribute value: " & value
            End If
        End If
    Next attributeNode
    For Each key In schema.Attributes.Keys
        rule = schema.Attributes(key)
        If rule(1) Then
            If VBA.Len(VBA.Trim$(ex_UiElementFactory.fn_Attribute(node, VBA.CStr(key)))) = 0 Then _
                private_AddError errors, path, VBA.CStr(key), "Required attribute is missing or empty."
        End If
    Next key
    For Each alternative In schema.RequiredAny
        found = False
        For Each candidate In VBA.Split(VBA.CStr(alternative), "|")
            If VBA.Len(VBA.Trim$(ex_UiElementFactory.fn_Attribute(node, VBA.CStr(candidate)))) > 0 Then found = True
        Next candidate
        If Not found Then private_AddError errors, path, VBA.CStr(alternative), "At least one attribute is required."
    Next alternative
    Set counts = VBA.CreateObject("Scripting.Dictionary")
    counts.CompareMode = VBA.vbBinaryCompare
    For Each child In node.ChildNodes
        If child.NodeType = 1 Then
            localName = VBA.CStr(child.baseName)
            namespaceUri = VBA.CStr(child.namespaceURI)
            tag = namespaceUri & "|" & VBA.LCase$(localName)
            If Not counts.Exists(tag) Then counts(tag) = 0
            counts(tag) = counts(tag) + 1
            childPath = path & "/" & VBA.CStr(child.nodeName) & "[" & VBA.CStr(counts(tag)) & "]"
            value = ex_UiElementFactory.fn_Attribute(child, "name")
            If VBA.Len(value) > 0 Then childPath = childPath & "[@name='" & value & "']"
            Set childSchema = Nothing
            If schema.Children.Exists(tag) Then
                rule = schema.Children(tag)
                Set childSchema = rule(0)
                If childSchema Is Nothing Then
                    If Not ex_UiElementFactory.fn_TryGetSchema(localName, childSchema, diagnostic, namespaceUri) Then private_AddError errors, childPath, tag, diagnostic
                End If
            ElseIf (schema.VisualChildren And Not (namespaceUri = "urn:excelprototype:profiles" And localName = "page")) Or (schema.ControlChildren And namespaceUri = "urn:excelprototype:controls") Then
                If Not ex_UiElementFactory.fn_TryGetSchema(localName, childSchema, diagnostic, namespaceUri) Then private_AddError errors, childPath, tag, diagnostic
            Else
                private_AddError errors, childPath, tag, "Child element is not allowed here."
            End If
            If Not childSchema Is Nothing Then private_ValidateNode child, childSchema, childPath, errors
        ElseIf child.NodeType = 3 Or child.NodeType = 4 Then
            value = VBA.Replace(VBA.Replace(VBA.Replace(VBA.CStr(child.Text), VBA.vbCr, ""), VBA.vbLf, ""), VBA.vbTab, "")
            If VBA.Len(VBA.Trim$(value)) > 0 Then private_AddError errors, path, "text", "Text content is not allowed."
        End If
    Next child
    For Each key In schema.Children.Keys
        rule = schema.Children(key)
        value = "0"
        If counts.Exists(key) Then value = VBA.CStr(counts(key))
        If VBA.CLng(value) < rule(1) Or (rule(2) >= 0 And VBA.CLng(value) > rule(2)) Then _
            private_AddError errors, path, VBA.CStr(key), "Invalid child count: " & value
    Next key
End Sub

Private Function private_AttributeValid(ByVal value As String, ByVal rule As Variant) As Boolean
    Dim source As String
    Dim path As String
    Dim number As Double

    On Error GoTo InvalidValue
    If VBA.Left$(VBA.Trim$(value), 1) = "{" And rule(0) <> "styleblock" Then
        If Not rule(3) Then Exit Function
        private_AttributeValid = ex_UiBindingRuntime.fn_TryParseBinding(value, "__scope", source, path)
        Exit Function
    End If
    If rule(0) = "context" Then Exit Function
    If rule(4) > 0 Then
        If VBA.Len(value) > rule(4) Then Exit Function
    End If
    Select Case rule(0)
        Case "gridsize"
            If value = "auto" Or value = "*" Then
                private_AttributeValid = True
                Exit Function
            End If
            If VBA.Right$(value, 1) = "*" Then value = VBA.Left$(value, VBA.Len(value) - 1)
            If Not VBA.IsNumeric(value) Then Exit Function
            number = VBA.CDbl(value)
            If number < 1 Or number <> VBA.Fix(number) Or number > 1048576 Then Exit Function
        Case "positive", "nonnegative", "number"
            If Not VBA.IsNumeric(value) Then Exit Function
            number = VBA.CDbl(value)
            If rule(0) <> "number" Then
                If number <> VBA.Fix(number) Or number > 2147483647# Then Exit Function
                If rule(0) = "positive" And number < 1 Then Exit Function
                If rule(0) = "nonnegative" And number < 0 Then Exit Function
            End If
        Case "boolean"
            If value <> "true" And value <> "false" Then Exit Function
        Case "enum"
            If VBA.InStr(1, "|" & rule(2) & "|", "|" & value & "|", VBA.vbBinaryCompare) = 0 Then Exit Function
    End Select
    private_AttributeValid = True
InvalidValue:
End Function