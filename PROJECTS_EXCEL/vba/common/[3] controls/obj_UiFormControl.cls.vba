VERSION 1.0 CLASS
BEGIN
  MultiUse = -1
END
Attribute VB_Name = "obj_UiFormControl"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Implements obj_IUiControl

Private m_panel As obj_IUiElement
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
' // Interface
' //
Private Function obj_IUiControl_Initialize() As Boolean
    obj_IUiControl_Initialize = True
End Function

Private Sub obj_IUiControl_Dispose()
    Me.Dispose
End Sub

Private Function obj_IUiControl_Configure( _
    ByVal definition As Object, _
    ByVal context As obj_UiRenderContext, _
    ByVal source As String, _
    ByRef diagnostic As String _
) As Boolean
    Dim node As Object
    Dim child As obj_IUiElement
    Dim container As obj_IUiContainer
    Dim panelNode As Object

    m_name = ex_UiElementFactory.fn_Attribute(definition, "name")
    If VBA.Len(m_name) = 0 Or VBA.Len(source) = 0 Then
        diagnostic = "Form requires name and dataContext (local or inherited)."
        Exit Function
    End If
    Set m_context = context
    m_row = ex_UiElementFactory.fn_Long(definition, "row", 1)
    m_column = ex_UiElementFactory.fn_Long(definition, "column", 1)
    Set m_panel = New obj_UiStackPanelElement
    Set container = m_panel
    Set panelNode = definition.cloneNode(False)
    panelNode.setAttribute "row", "1"
    panelNode.setAttribute "column", "1"
    If Not m_panel.Configure(panelNode, context, source, diagnostic) Then Exit Function
    For Each node In definition.ChildNodes
        If node.NodeType = 1 Then
            Select Case VBA.CStr(node.namespaceURI) & "|" & VBA.LCase$(VBA.CStr(node.baseName))
                Case "urn:excelprototype:profiles|field"
                    Set child = New obj_UiFormField
                    If Not child.Configure(node, context, source, diagnostic) Then
                        child.Dispose
                        Exit Function
                    End If
                Case Else
                    If VBA.CStr(node.namespaceURI) <> "urn:excelprototype:controls" Then
                        diagnostic = "Form only accepts field and controls namespace children: " & m_name
                        Exit Function
                    End If
                    Set child = ex_UiElementFactory.fn_Create(node, context, source, diagnostic)
                    If child Is Nothing Then Exit Function
            End Select
            container.AddChild child
        End If
    Next node
    context.RegisterForm m_name, Me
    obj_IUiControl_Configure = True
End Function

Private Function obj_IUiControl_Measure( _
    ByRef rows As Long, _
    ByRef columns As Long, _
    ByRef diagnostic As String _
) As Boolean
    obj_IUiControl_Measure = m_panel.Measure(rows, columns, diagnostic)
    rows = rows + m_row - 1
    columns = columns + m_column - 1
End Function

Private Function obj_IUiControl_Arrange( _
    ByVal row As Long, _
    ByVal column As Long, _
    ByRef diagnostic As String _
) As Boolean
    obj_IUiControl_Arrange = m_panel.Arrange(row + m_row - 1, column + m_column - 1, diagnostic)
End Function

Private Function obj_IUiControl_Render(ByRef diagnostic As String) As Boolean
    obj_IUiControl_Render = m_panel.Render(diagnostic)
End Function

Private Function obj_IUiControl_Validate(ByVal errors As Collection) As Boolean
    obj_IUiControl_Validate = m_panel.Validate(errors)
End Function

' //
' // API
' //
Public Sub Dispose()
    If Not m_panel Is Nothing Then m_panel.Dispose
    Set m_panel = Nothing
    Set m_context = Nothing
End Sub