VERSION 1.0 CLASS
BEGIN
  MultiUse = -1
END
Attribute VB_Name = "obj_UiGridElement"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Implements obj_IUiElement
Implements obj_IUiContainer

Private m_children As Collection
Private m_definition As Object
Private m_context As obj_UiRenderContext
Private m_source As String
Private m_rows As Long
Private m_columns As Long
Private m_row As Long
Private m_column As Long

' //
' // Lifecycle
' //
Private Sub Class_Initialize()
    Set m_children = New Collection
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
    obj_IUiElement_Configure = private_ConfigurePanel(definition, context, source, diagnostic)
End Function

Private Function obj_IUiElement_Measure( _
    ByRef rows As Long, _
    ByRef columns As Long, _
    ByRef diagnostic As String _
) As Boolean
    obj_IUiElement_Measure = private_MeasurePanel(rows, columns, diagnostic)
    rows = rows + m_row - 1
    columns = columns + m_column - 1
End Function

Private Function obj_IUiElement_Arrange( _
    ByVal row As Long, _
    ByVal column As Long, _
    ByRef diagnostic As String _
) As Boolean
    row = row + m_row - 1
    column = column + m_column - 1
    obj_IUiElement_Arrange = private_ArrangePanel(row, column, diagnostic)
End Function

Private Function obj_IUiElement_Render(ByRef diagnostic As String) As Boolean
    obj_IUiElement_Render = private_RenderPanel(diagnostic)
End Function

Private Function obj_IUiElement_Validate(ByVal errors As Collection) As Boolean
    obj_IUiElement_Validate = private_ValidatePanel(errors)
End Function

Private Sub obj_IUiElement_Dispose()
    Me.Dispose
End Sub

Private Sub obj_IUiContainer_AddChild(ByVal child As obj_IUiElement)
    m_children.Add child
End Sub

' //
' // API
' //
Public Sub Dispose()
    private_DisposePanel
End Sub

' //
' // Private
' //
Private Function private_ConfigurePanel( _
    ByVal definition As Object, _
    ByVal context As obj_UiRenderContext, _
    ByVal source As String, _
    ByRef diagnostic As String _
) As Boolean
    Dim node As Object
    Dim child As obj_IUiElement

    Set m_definition = definition
    m_row = ex_UiElementFactory.fn_Long(definition, "row", 1)
    m_column = ex_UiElementFactory.fn_Long(definition, "column", 1)
    Set m_context = context
    m_source = source
    If VBA.Len(ex_UiElementFactory.fn_Attribute(definition, "source")) > 0 Then _
        m_source = ex_UiElementFactory.fn_Attribute(definition, "source")
        For Each node In definition.ChildNodes
            If node.NodeType = 1 Then
                If VBA.LCase$(VBA.CStr(node.baseName)) <> "styles" Then
                    Set child = ex_UiElementFactory.fn_Create(node, context, m_source, diagnostic)
                    If child Is Nothing Then Exit Function
                    m_children.Add child
                End If
            End If
        Next node
    private_ConfigurePanel = True
End Function

Private Function private_MeasurePanel( _
    ByRef rows As Long, _
    ByRef columns As Long, _
    ByRef diagnostic As String _
) As Boolean
    Dim child As obj_IUiElement
    Dim height As Long
    Dim width As Long

    rows = 0
    columns = 0
    For Each child In m_children
        If Not child.Measure(height, width, diagnostic) Then Exit Function
        If height > rows Then rows = height
        If width > columns Then columns = width
    Next child
    If rows = 0 Then rows = 1
    If columns = 0 Then columns = 1
    m_rows = rows
    m_columns = columns
    private_MeasurePanel = True
End Function

Private Function private_ArrangePanel( _
    ByVal row As Long, _
    ByVal column As Long, _
    ByRef diagnostic As String _
) As Boolean
    Dim child As obj_IUiElement
    Dim height As Long
    Dim width As Long
    Dim nextRow As Long
    Dim nextColumn As Long

    nextRow = row
    nextColumn = column
    For Each child In m_children
        If Not child.Measure(height, width, diagnostic) Then Exit Function
        If Not child.Arrange(nextRow, nextColumn, diagnostic) Then Exit Function
    Next child
    private_ArrangePanel = True
End Function

Private Function private_RenderPanel(ByRef diagnostic As String) As Boolean
    Dim child As obj_IUiElement

    For Each child In m_children
        If Not child.Render(diagnostic) Then Exit Function
    Next child
    private_RenderPanel = True
End Function

Private Function private_ValidatePanel(ByVal errors As Collection) As Boolean
    Dim child As obj_IUiElement

    private_ValidatePanel = True
    For Each child In m_children
        If Not child.Validate(errors) Then private_ValidatePanel = False
    Next child
End Function

Private Sub private_DisposePanel()
    Dim child As obj_IUiElement

    If Not m_children Is Nothing Then
        For Each child In m_children
            child.Dispose
        Next child
    End If
    Set m_children = Nothing
    Set m_definition = Nothing
    Set m_context = Nothing
End Sub