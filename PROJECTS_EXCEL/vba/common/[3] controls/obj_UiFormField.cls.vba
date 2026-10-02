VERSION 1.0 CLASS
BEGIN
  MultiUse = -1
END
Attribute VB_Name = "obj_UiFormField"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Implements obj_IUiElement

Private m_panel As obj_UiPanel
Private m_context As obj_UiRenderContext
Private m_source As String
Private m_path As String
Private m_name As String
Private m_required As Boolean
Private m_checkbox As Boolean

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
Public Sub Dispose()
    If Not m_panel Is Nothing Then m_panel.DisposePanel
    Set m_panel = Nothing
    Set m_context = Nothing
End Sub

Public Function ConfigureField( _
    ByVal fieldNode As Object, _
    ByVal formNode As Object, _
    ByVal context As obj_UiRenderContext, _
    ByVal source As String, _
    ByRef diagnostic As String _
) As Boolean
    Dim labelNode As Object
    Dim editorNode As Object
    Dim panelNode As Object
    Dim nodeAttribute As Object
    Dim label As obj_IUiElement
    Dim editor As obj_IUiElement
    Dim kind As String
    Dim position As String
    Dim rawValue As String
    Dim booleanName As Variant
    Dim booleanValue As String

    m_name = ex_UiElementFactory.fn_Attribute(fieldNode, "name")
    kind = VBA.LCase$(ex_UiElementFactory.fn_Attribute(fieldNode, "type"))
    If VBA.Len(m_name) = 0 Or VBA.Len(m_name) > 25 Then
        diagnostic = "Field name must contain 1 to 25 characters."
        Exit Function
    End If
    If kind <> "text" And kind <> "select" And kind <> "checkbox" Then
        diagnostic = "Unsupported field type: " & kind
        Exit Function
    End If
    position = private_Inherit(fieldNode, formNode, "labelPosition", "left")
    If position <> "left" And position <> "top" Then
        diagnostic = "labelPosition must be left or top."
        Exit Function
    End If
    If VBA.Len(ex_UiElementFactory.fn_Attribute(fieldNode, "label")) = 0 Then
        diagnostic = "Field requires label: " & m_name
        Exit Function
    End If
    For Each booleanName In VBA.Array("required", "readOnly")
        booleanValue = VBA.LCase$(private_Inherit(fieldNode, formNode, VBA.CStr(booleanName), "false"))
        If booleanValue <> "true" And booleanValue <> "false" Then
            diagnostic = "Boolean field attribute required: " & VBA.CStr(booleanName)
            Exit Function
        End If
    Next booleanName
    Set m_context = context
    m_required = (VBA.LCase$(ex_UiElementFactory.fn_Attribute(fieldNode, "required")) = "true")
    m_checkbox = (kind = "checkbox")
    rawValue = ex_UiElementFactory.fn_Attribute(fieldNode, "value")
    If VBA.Len(rawValue) = 0 Then rawValue = "{Binding Path=" & m_name & "}"
    If Not ex_UiBindingRuntime.fn_TryParseBinding(rawValue, source, m_source, m_path, context.BindingContext) Then
        diagnostic = "Invalid field binding: " & m_name
        Exit Function
    End If
    ' Изолированные описания для адаптера существующих контролов; DOM страницы не изменяется.
    Set panelNode = fieldNode.cloneNode(False)
    panelNode.setAttribute "orientation", VBA.IIf(position = "left", "horizontal", "vertical")
    Set m_panel = New obj_UiPanel
    m_panel.Initialize "stack"
    If Not m_panel.ConfigurePanel(panelNode, context, source, diagnostic, False) Then Exit Function
    Set labelNode = fieldNode.cloneNode(False)
    labelNode.setAttribute "type", "Label"
    labelNode.setAttribute "name", m_name & "_label"
    labelNode.setAttribute "text", ex_UiElementFactory.fn_Attribute(fieldNode, "label")
    labelNode.setAttribute "columnSpan", private_Inherit(fieldNode, formNode, "labelColumnSpan", "2")
    labelNode.setAttribute "style", private_Inherit(fieldNode, formNode, "labelStyle", VBA.vbNullString)
    Set label = New obj_UiControlElement
    If Not label.Configure(labelNode, context, source, diagnostic) Then Exit Function
    m_panel.AddChild label
    Set editorNode = fieldNode.cloneNode(False)
    editorNode.setAttribute "type", VBA.IIf(kind = "select", "Select", "Input")
    editorNode.setAttribute "inputType", kind
    editorNode.setAttribute "name", m_name & "_input"
    editorNode.setAttribute "value", rawValue
    editorNode.setAttribute "columnSpan", private_Inherit(fieldNode, formNode, "fieldWidth", "4")
    editorNode.setAttribute "rowSpan", private_Inherit(fieldNode, formNode, "height", "1")
    editorNode.setAttribute "style", private_Inherit(fieldNode, formNode, "fieldStyle", VBA.vbNullString)
    editorNode.setAttribute "onChange", private_Inherit(fieldNode, formNode, "onChange", VBA.vbNullString)
    editorNode.setAttribute "readOnly", private_Inherit(fieldNode, formNode, "readOnly", "false")
    Set editor = New obj_UiControlElement
    If Not editor.Configure(editorNode, context, source, diagnostic) Then Exit Function
    m_panel.AddChild editor
    ConfigureField = True
End Function

' //
' // Private
' //
Private Function private_Inherit( _
    ByVal fieldNode As Object, _
    ByVal formNode As Object, _
    ByVal name As String, _
    ByVal defaultValue As String _
) As String
    private_Inherit = ex_UiElementFactory.fn_Attribute(fieldNode, name)
    If VBA.Len(private_Inherit) = 0 Then private_Inherit = ex_UiElementFactory.fn_Attribute(formNode, name)
    If VBA.Len(private_Inherit) = 0 Then private_Inherit = defaultValue
End Function

' //
' // Interface
' //
Private Function obj_IUiElement_Configure( _
    ByVal definition As Object, _
    ByVal context As obj_UiRenderContext, _
    ByVal source As String, _
    ByRef diagnostic As String _
) As Boolean
    diagnostic = "Field must be configured by its owning form."
End Function

Private Function obj_IUiElement_Measure( _
    ByRef rows As Long, _
    ByRef columns As Long, _
    ByRef diagnostic As String _
) As Boolean
    obj_IUiElement_Measure = m_panel.MeasurePanel(rows, columns, diagnostic)
End Function

Private Function obj_IUiElement_Arrange( _
    ByVal row As Long, _
    ByVal column As Long, _
    ByRef diagnostic As String _
) As Boolean
    obj_IUiElement_Arrange = m_panel.ArrangePanel(row, column, diagnostic)
End Function

Private Function obj_IUiElement_Render(ByRef diagnostic As String) As Boolean
    obj_IUiElement_Render = m_panel.RenderPanel(diagnostic)
End Function

Private Function obj_IUiElement_Validate(ByVal errors As Collection) As Boolean
    Dim value As Variant
    Dim sourceObject As Object
    Dim isObject As Boolean

    obj_IUiElement_Validate = True
    If Not m_required Then Exit Function
    If Not m_context.BindingContext.TryGetValue(m_source, m_path, value, sourceObject, isObject) Then
        obj_IUiElement_Validate = False
    ElseIf isObject Or VBA.IsNull(value) Or VBA.IsError(value) Then
        obj_IUiElement_Validate = False
    ElseIf m_checkbox Then
        obj_IUiElement_Validate = (VBA.VarType(value) = VBA.vbBoolean)
        If obj_IUiElement_Validate Then obj_IUiElement_Validate = VBA.CBool(value)
    Else
        obj_IUiElement_Validate = (VBA.Len(VBA.Trim$(VBA.CStr(value))) > 0)
    End If
    If Not obj_IUiElement_Validate Then errors.Add "Required field: " & m_name
End Function

Private Sub obj_IUiElement_Dispose()
    Me.Dispose
End Sub