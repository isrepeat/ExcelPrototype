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

Private m_isInitialized As Boolean
Private m_isDisposed As Boolean

Private m_axisRows As obj_UiGridAxis
Private m_axisColumns As obj_UiGridAxis
Private m_slots As Collection
Private m_trackLayout As Boolean
Private m_children As Collection
Private m_definition As Object
Private m_context As obj_UiRenderContext
Private m_source As String
Private m_rows As Long
Private m_columns As Long
Private m_row As Long
Private m_column As Long
Private m_arrangedRow As Long
Private m_arrangedColumn As Long

' //
' // Lifecycle
' //
Private Sub Class_Initialize()
    Set m_children = New Collection
    Set m_slots = New Collection
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
    If m_isDisposed Or m_isInitialized Then
        diagnostic = "Element is already initialized or disposed."
        Exit Function
    End If
    obj_IUiElement_Configure = private_ConfigurePanel(definition, context, source, diagnostic)
    m_isInitialized = obj_IUiElement_Configure
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
    If m_isDisposed Then
        Exit Sub
    End If
    m_isDisposed = True
    m_isInitialized = False
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
    Set m_axisRows = New obj_UiGridAxis
    Set m_axisColumns = New obj_UiGridAxis
    If Not m_axisRows.Initialize(definition, "grid.rowDefinitions", "rowDefinition", "rowSpan", diagnostic) Then Exit Function
    If Not m_axisColumns.Initialize(definition, "grid.columnDefinitions", "columnDefinition", "columnSpan", diagnostic) Then Exit Function
    m_trackLayout = m_axisRows.HasDefinitions Or m_axisColumns.HasDefinitions
    For Each node In definition.ChildNodes
        If node.NodeType = 1 Then
            If Not (VBA.CStr(node.namespaceURI) = "urn:excelprototype:profiles" And _
                    (VBA.LCase$(VBA.CStr(node.baseName)) = "styles" Or _
                     VBA.LCase$(VBA.CStr(node.baseName)) = "grid.rowdefinitions" Or _
                     VBA.LCase$(VBA.CStr(node.baseName)) = "grid.columndefinitions")) Then
                Dim childDefinition As Object
                Dim slotRow As Long, slotColumn As Long, slotRows As Long, slotColumns As Long
                Set childDefinition = node.cloneNode(True)
                slotRow = ex_UiElementFactory.fn_Long(node, "row", 1)
                slotColumn = ex_UiElementFactory.fn_Long(node, "column", 1)
                slotRows = ex_UiElementFactory.fn_Long(node, "rowSpan", 1)
                slotColumns = ex_UiElementFactory.fn_Long(node, "columnSpan", 1)
                If m_trackLayout Then
                    If Not m_axisRows.Contains(slotRow, slotRows) Or Not m_axisColumns.Contains(slotColumn, slotColumns) Then
                        diagnostic = "Grid child exceeds row or column definitions: " & VBA.CStr(node.nodeName)
                        Exit Function
                    End If
                    childDefinition.setAttribute "row", "1"
                    childDefinition.setAttribute "column", "1"
                    childDefinition.removeAttribute "rowSpan"
                    childDefinition.removeAttribute "columnSpan"
                End If
                Set child = ex_UiElementFactory.fn_Create(childDefinition, context, m_source, diagnostic)
                If child Is Nothing Then Exit Function
                m_children.Add child
                m_slots.Add VBA.Array(slotRow, slotColumn, slotRows, slotColumns)
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

    If m_trackLayout Then
        Dim index As Long, pass As Long
        Dim slot As Variant
        m_axisRows.Reset
        m_axisColumns.Reset
        For pass = 1 To 2
            For index = 1 To m_children.Count
                Set child = m_children(index)
                slot = m_slots(index)
                If Not child.Measure(height, width, diagnostic) Then Exit Function
                If (pass = 1 And slot(2) = 1) Or (pass = 2 And slot(2) > 1) Then m_axisRows.Grow slot(0), slot(2), height
                If (pass = 1 And slot(3) = 1) Or (pass = 2 And slot(3) > 1) Then m_axisColumns.Grow slot(1), slot(3), width
            Next index
        Next pass
        If Not m_axisRows.Resolve(diagnostic) Then Exit Function
        If Not m_axisColumns.Resolve(diagnostic) Then Exit Function
        rows = m_axisRows.Total
        columns = m_axisColumns.Total
        m_rows = rows
        m_columns = columns
        private_MeasurePanel = True
        Exit Function
    End If
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

    m_arrangedRow = row
    m_arrangedColumn = column
    If m_trackLayout Then
        Dim index As Long
        Dim slot As Variant
        Dim layoutSlot As obj_IUiLayoutSlot
        For index = 1 To m_children.Count
            Set child = m_children(index)
            slot = m_slots(index)
            If TypeOf child Is obj_IUiLayoutSlot Then
                Set layoutSlot = child
                layoutSlot.SetSize m_axisRows.Extent(slot(0), slot(2)), m_axisColumns.Extent(slot(1), slot(3))
            End If
            If Not child.Arrange(row + m_axisRows.Offset(slot(0)), column + m_axisColumns.Offset(slot(1)), diagnostic) Then Exit Function
        Next index
        private_ArrangePanel = True
        Exit Function
    End If
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

    If m_rows > 0 And m_columns > 0 Then
        m_context.Styles.RegisterPart m_definition, m_context.TargetWorksheet.Cells(m_arrangedRow, m_arrangedColumn).Resize(m_rows, m_columns), "layout"
    End If
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
    Set m_slots = Nothing
    If Not m_axisRows Is Nothing Then m_axisRows.Dispose
    If Not m_axisColumns Is Nothing Then m_axisColumns.Dispose
    Set m_axisRows = Nothing
    Set m_axisColumns = Nothing
    Set m_definition = Nothing
    Set m_context = Nothing
End Sub