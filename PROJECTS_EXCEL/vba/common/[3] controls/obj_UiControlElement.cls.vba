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

Private m_definition As Object
Private m_context As obj_UiRenderContext
Private WithEvents m_bindingContext As obj_UiBindingContext
Private m_control As obj_IUiControl
Private m_rows As Long
Private m_columns As Long
Private m_row As Long
Private m_column As Long

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
    If Not m_control Is Nothing Then m_control.Dispose
    Set m_control = Nothing
    Set m_bindingContext = Nothing
    Set m_context = Nothing
    Set m_definition = Nothing
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

    m_row = ex_UiElementFactory.fn_Long(definition, "row", 1)
    m_column = ex_UiElementFactory.fn_Long(definition, "column", 1)
    Set m_definition = definition.cloneNode(True)
    Set m_context = context
    Set m_bindingContext = context.BindingContext
    For Each attributeNode In m_definition.Attributes
        If VBA.StrComp(VBA.Left$(VBA.Trim$(VBA.CStr(attributeNode.Text)), 9), _
                "{Binding ", VBA.vbTextCompare) = 0 Then
            If Not ex_UiBindingRuntime.fn_TryParseBinding(VBA.CStr(attributeNode.Text), _
                    source, bindingSource, bindingPath, context.BindingContext) Then
                diagnostic = "Invalid binding on control: " & ex_UiElementFactory.fn_Attribute(definition, "name")
                Exit Function
            End If
            attributeNode.Text = "{Binding Source=" & bindingSource & "; Path=" & bindingPath & "}"
        End If
    Next attributeNode
    m_definition.setAttribute "row", "1"
    m_definition.setAttribute "column", "1"
    Set m_control = ex_UiControlFactory.fn_Create(m_definition)
    If m_control Is Nothing Then
        diagnostic = "Cannot create control: " & ex_UiElementFactory.fn_Attribute(definition, "name")
        Exit Function
    End If
    If Not m_control.Configure(m_definition) Then
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
    Dim range As Range

    If Not m_control.Configure(m_definition) Then Exit Function
    Set range = m_control.Measure(m_context)
    If range Is Nothing Then
        diagnostic = "Cannot measure control: " & ex_UiElementFactory.fn_Attribute(m_definition, "name")
        Exit Function
    End If
    rows = range.Rows.Count
    columns = range.Columns.Count
    If rows < 1 Or columns < 1 Then
        diagnostic = "Control spans must be positive."
        Exit Function
    End If
    m_rows = rows
    m_columns = columns
    rows = rows + m_row - 1
    columns = columns + m_column - 1
    obj_IUiElement_Measure = True
End Function

Private Function obj_IUiElement_Arrange( _
    ByVal row As Long, _
    ByVal column As Long, _
    ByRef diagnostic As String _
) As Boolean
    Dim range As Range

    m_definition.setAttribute "row", VBA.CStr(row + m_row - 1)
    m_definition.setAttribute "column", VBA.CStr(column + m_column - 1)
    If Not m_control.Configure(m_definition) Then Exit Function
    Set range = m_control.Measure(m_context)
    obj_IUiElement_Arrange = Not range Is Nothing
End Function

Private Function obj_IUiElement_Render(ByRef diagnostic As String) As Boolean
    obj_IUiElement_Render = m_control.Render(m_context)
    If Not obj_IUiElement_Render Then diagnostic = "Cannot render control: " & ex_UiElementFactory.fn_Attribute(m_definition, "name")
End Function

Private Function obj_IUiElement_Validate(ByVal errors As Collection) As Boolean
    obj_IUiElement_Validate = True
End Function

Private Sub obj_IUiElement_Dispose()
    Me.Dispose
End Sub

' //
' // Private
' //
Private Sub m_bindingContext_ValueChanged(ByVal sourceName As String, ByVal bindingPath As String)
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