VERSION 1.0 CLASS
BEGIN
  MultiUse = -1
END
Attribute VB_Name = "obj_UiControlElement"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Implements obj_IUiElement
Implements obj_IUiLayoutSlot

Private m_definition As Object
Private m_intrinsicRows As String
Private m_intrinsicColumns As String
Private m_context As obj_UiRenderContext
Private WithEvents m_bindingContext As obj_UiBindingContext
Private m_control As obj_IUiControl

' //
' // Lifecycle
' //
Private Sub Class_Initialize()
End Sub

Private Sub Class_Terminate()
    Me.Dispose
End Sub

' //
' // Interface
' //
Private Function obj_IUiElement_Configure( _
    ByVal definition As Object, _
    ByVal context As obj_UiRenderContext, _
    ByVal source As String, _
    ByRef diagnostic As String _
) As Boolean
    Dim attributeNode As Object
    Dim bindingSource As String
    Dim bindingPath As String

    Set m_definition = definition.cloneNode(True)
    m_intrinsicRows = ex_UiElementFactory.fn_Attribute(definition, "rowSpan")
    m_intrinsicColumns = ex_UiElementFactory.fn_Attribute(definition, "columnSpan")
    If VBA.CStr(definition.namespaceURI) = "urn:excelprototype:controls" Then _
        m_definition.setAttribute "type", VBA.CStr(definition.baseName)
    Set m_context = context
    Set m_bindingContext = context.BindingContext
    For Each attributeNode In m_definition.Attributes
        If attributeNode.nodeName <> "dataContext" And VBA.StrComp(VBA.Left$(VBA.Trim$(VBA.CStr(attributeNode.Text)), 9), _
                "{Binding ", VBA.vbTextCompare) = 0 Then
            If Not ex_UiBindingRuntime.fn_TryParseBinding(VBA.CStr(attributeNode.Text), _
                    source, bindingSource, bindingPath, context.BindingContext) Then
                diagnostic = "Invalid binding on control: " & ex_UiElementFactory.fn_Attribute(definition, "name")
                Exit Function
            End If
            attributeNode.Text = "{Binding Source=" & bindingSource & "; Path=" & bindingPath & "}"
        End If
    Next attributeNode
    Set m_control = ex_UiControlFactory.fn_Create(m_definition)
    If m_control Is Nothing Then
        diagnostic = "Cannot create control: " & ex_UiElementFactory.fn_Attribute(definition, "name")
        Exit Function
    End If
    If Not m_control.Configure(m_definition, context, source, diagnostic) Then
        diagnostic = "Cannot configure control: " & ex_UiElementFactory.fn_Attribute(definition, "name")
        Exit Function
    End If
    context.RegisterElement ex_UiElementFactory.fn_Attribute(definition, "name"), Me
    obj_IUiElement_Configure = True
End Function

Private Function obj_IUiElement_Measure( _
    ByRef rows As Long, _
    ByRef columns As Long, _
    ByRef diagnostic As String _
) As Boolean
    If VBA.Len(m_intrinsicRows) = 0 Then
        m_definition.removeAttribute "rowSpan"
    Else
        m_definition.setAttribute "rowSpan", m_intrinsicRows
    End If
    If VBA.Len(m_intrinsicColumns) = 0 Then
        m_definition.removeAttribute "columnSpan"
    Else
        m_definition.setAttribute "columnSpan", m_intrinsicColumns
    End If
    obj_IUiElement_Measure = m_control.Measure(rows, columns, diagnostic)
End Function

Private Function obj_IUiElement_Arrange( _
    ByVal row As Long, _
    ByVal column As Long, _
    ByRef diagnostic As String _
) As Boolean
    obj_IUiElement_Arrange = m_control.Arrange(row, column, diagnostic)
End Function

Private Function obj_IUiElement_Render(ByRef diagnostic As String) As Boolean
    obj_IUiElement_Render = m_control.Render(diagnostic)
    If Not obj_IUiElement_Render Then diagnostic = "Cannot render control: " & ex_UiElementFactory.fn_Attribute(m_definition, "name")
End Function

Private Function obj_IUiElement_Validate(ByVal errors As Collection) As Boolean
    obj_IUiElement_Validate = m_control.Validate(errors)
End Function

Private Sub obj_IUiElement_Dispose()
    Me.Dispose
End Sub

Private Sub obj_IUiLayoutSlot_SetSize(ByVal rows As Long, ByVal columns As Long)
    Select Case VBA.LCase$(ex_UiElementFactory.fn_Attribute(m_definition, "type"))
        Case "label", "button", "input", "select"
            m_definition.setAttribute "rowSpan", VBA.CStr(rows)
            m_definition.setAttribute "columnSpan", VBA.CStr(columns)
    End Select
End Sub

' //
' // API
' //
Public Sub Dispose()
    If Not m_control Is Nothing Then m_control.Dispose
    Set m_control = Nothing
    Set m_bindingContext = Nothing
    Set m_context = Nothing
    Set m_definition = Nothing
End Sub

' //
' // Private
' //
Private Sub m_bindingContext_ValueChanged(ByVal sourceName As String, ByVal bindingPath As String)
    private_bindingContext_ValueChanged sourceName, bindingPath
End Sub

Private Sub private_bindingContext_ValueChanged(ByVal sourceName As String, ByVal bindingPath As String)
    Dim nodeAttribute As Object
    Dim source As String
    Dim path As String
    Dim target As obj_IUiBindingTarget
    Dim previousEvents As Boolean

    If m_control Is Nothing Then Exit Sub
    For Each nodeAttribute In m_definition.Attributes
        If ex_UiBindingRuntime.fn_TryParseBinding(VBA.CStr(nodeAttribute.Text), _
                VBA.vbNullString, source, path) Then
            If VBA.StrComp(source, sourceName, VBA.vbTextCompare) = 0 Then
                If VBA.StrComp(path, bindingPath, VBA.vbTextCompare) = 0 Or _
                        VBA.StrComp(VBA.Left$(path, VBA.Len(bindingPath) + 1), bindingPath & ".", VBA.vbTextCompare) = 0 Then
                    If TypeOf m_control Is obj_IUiBindingTarget Then
                        Set target = m_control
                        previousEvents = Application.EnableEvents
                        On Error GoTo EH_REFRESH
                        Application.EnableEvents = False
                        target.RefreshBindings m_context
                        Application.EnableEvents = previousEvents
                    ElseIf nodeAttribute.nodeName <> "value" Then
                        m_context.InvalidateMeasure
                    End If
                    Exit Sub
                End If
            End If
        End If
    Next nodeAttribute
    Exit Sub
EH_REFRESH:
    Application.EnableEvents = previousEvents
    VBA.Err.Raise VBA.Err.Number, "Control.RefreshBindings", VBA.Err.Description
End Sub