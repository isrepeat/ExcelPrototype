VERSION 1.0 CLASS
BEGIN
  MultiUse = -1
END
Attribute VB_Name = "obj_UiForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Implements obj_IUiElement

Private m_panel As obj_UiPanel
Private m_name As String
Private m_context As obj_UiRenderContext
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
    If Not m_panel Is Nothing Then m_panel.DisposePanel
    Set m_panel = Nothing
    Set m_context = Nothing
End Sub

Public Function Validate(ByVal errors As Collection) As Boolean
    Validate = m_panel.ValidatePanel(errors)
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
    Dim node As Object
    Dim child As obj_IUiElement
    Dim field As obj_UiFormField

    m_name = ex_UiElementFactory.fn_Attribute(definition, "name")
    source = ex_UiElementFactory.fn_Attribute(definition, "source")
    If VBA.Len(m_name) = 0 Or VBA.Len(source) = 0 Then
        diagnostic = "Form requires name and source."
        Exit Function
    End If
    Set m_context = context
    m_row = ex_UiElementFactory.fn_Long(definition, "row", 1)
    m_column = ex_UiElementFactory.fn_Long(definition, "column", 1)
    Set m_panel = New obj_UiPanel
    m_panel.Initialize "stack"
    If Not m_panel.ConfigurePanel(definition, context, source, diagnostic, False) Then Exit Function
    For Each node In definition.ChildNodes
        If node.NodeType = 1 Then
            Select Case VBA.LCase$(VBA.CStr(node.baseName))
                Case "field"
                    Set field = New obj_UiFormField
                    If Not field.ConfigureField(node, definition, context, source, diagnostic) Then Exit Function
                    Set child = field
                Case "control"
                    Set child = ex_UiElementFactory.fn_Create(node, context, source, diagnostic)
                    If child Is Nothing Then Exit Function
                Case Else
                    diagnostic = "Form only accepts field and control children: " & m_name
                    Exit Function
            End Select
            m_panel.AddChild child
        End If
    Next node
    context.RegisterForm m_name, Me
    obj_IUiElement_Configure = True
End Function

Private Function obj_IUiElement_Measure( _
    ByRef rows As Long, _
    ByRef columns As Long, _
    ByRef diagnostic As String _
) As Boolean
    obj_IUiElement_Measure = m_panel.MeasurePanel(rows, columns, diagnostic)
    rows = rows + m_row - 1
    columns = columns + m_column - 1
End Function

Private Function obj_IUiElement_Arrange( _
    ByVal row As Long, _
    ByVal column As Long, _
    ByRef diagnostic As String _
) As Boolean
    obj_IUiElement_Arrange = m_panel.ArrangePanel(row + m_row - 1, column + m_column - 1, diagnostic)
End Function

Private Function obj_IUiElement_Render(ByRef diagnostic As String) As Boolean
    obj_IUiElement_Render = m_panel.RenderPanel(diagnostic)
End Function

Private Function obj_IUiElement_Validate(ByVal errors As Collection) As Boolean
    obj_IUiElement_Validate = Validate(errors)
End Function

Private Sub obj_IUiElement_Dispose()
    Me.Dispose
End Sub