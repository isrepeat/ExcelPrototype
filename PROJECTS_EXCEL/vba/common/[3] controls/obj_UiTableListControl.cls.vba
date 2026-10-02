VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "obj_UiTableListControl"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Implements obj_IUiControl

Private m_isInitialized As Boolean
Private m_isDisposed As Boolean

Private m_context As obj_UiRenderContext
Private m_definition As Object
Private m_base As obj_UiControlBase
Private m_children As Collection
Private m_sizes As Collection
Private m_gapRows As Long

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
    If m_isDisposed Or m_isInitialized Then
        Exit Function
    End If
    Set m_base = New obj_UiControlBase
    Set m_children = New Collection
    Set m_sizes = New Collection
    obj_IUiControl_Initialize = m_base.Initialize()
    m_isInitialized = obj_IUiControl_Initialize
End Function

Private Sub obj_IUiControl_Dispose()
    Me.Dispose
End Sub

Private Function obj_IUiControl_Configure(ByVal definition As Object, ByVal context As obj_UiRenderContext, ByVal source As String, ByRef diagnostic As String) As Boolean
    Set m_context = context
    Set m_definition = definition.cloneNode(True)
    m_base.ConfigurePosition definition
    m_gapRows = 1
    If VBA.Len(ex_UiElementFactory.fn_Attribute(definition, "gapRows")) > 0 Then _
        m_gapRows = VBA.CLng(ex_UiElementFactory.fn_Attribute(definition, "gapRows"))
    obj_IUiControl_Configure = True
End Function

Private Function obj_IUiControl_Measure(ByRef rows As Long, ByRef columns As Long, ByRef diagnostic As String) As Boolean
    Dim value As Variant, sourceObject As Object, isObject As Boolean
    Dim tableSource As obj_IUiTableSource
    Dim child As obj_IUiControl
    Dim target As obj_IUiTableTarget
    Dim definition As Object
    Dim raw As String
    Dim index As Long, childRows As Long, childColumns As Long

    private_ClearChildren
    raw = ex_UiElementFactory.fn_Attribute(m_definition, "itemsSource")
    If Not ex_UiBindingRuntime.fn_TryResolveValue(raw, m_context.BindingContext, value, sourceObject, isObject) Then
        diagnostic = "Cannot resolve tableList source: " & raw
        Exit Function
    End If
    If Not isObject Then Exit Function
    If Not TypeOf sourceObject Is obj_IUiTableSource Then
        diagnostic = "tableList requires obj_IUiTableSource"
        Exit Function
    End If
    Set tableSource = sourceObject
    rows = 0
    columns = 1
    For index = 1 To tableSource.TableCount
        Set definition = m_definition.cloneNode(False)
        definition.setAttribute "type", "Table"
        definition.setAttribute "row", "1"
        definition.setAttribute "column", "1"
        definition.removeAttribute "itemsSource"
        definition.removeAttribute "gapRows"
        Set child = ex_UiControlFactory.fn_Create(definition)
        If child Is Nothing Then Exit Function
        m_children.Add child
        Set target = child
        target.SetTable tableSource.GetTable(index)
        If Not child.Configure(definition, m_context, "", diagnostic) Then Exit Function
        If Not child.Measure(childRows, childColumns, diagnostic) Then Exit Function
        m_sizes.Add childRows
        rows = rows + childRows
        If index < tableSource.TableCount Then rows = rows + m_gapRows
        If childColumns > columns Then columns = childColumns
    Next index
    If rows = 0 Then rows = 1
    rows = rows + ex_UiElementFactory.fn_Long(m_definition, "row", 1) - 1
    columns = columns + ex_UiElementFactory.fn_Long(m_definition, "column", 1) - 1
    obj_IUiControl_Measure = True
End Function

Private Function obj_IUiControl_Arrange(ByVal row As Long, ByVal column As Long, ByRef diagnostic As String) As Boolean
    Dim child As obj_IUiControl
    Dim index As Long

    row = row + ex_UiElementFactory.fn_Long(m_definition, "row", 1) - 1
    column = column + ex_UiElementFactory.fn_Long(m_definition, "column", 1) - 1
    For index = 1 To m_children.Count
        Set child = m_children(index)
        If Not child.Arrange(row, column, diagnostic) Then Exit Function
        row = row + VBA.CLng(m_sizes(index)) + m_gapRows
    Next index
    obj_IUiControl_Arrange = True
End Function

Private Function obj_IUiControl_Render(ByRef diagnostic As String) As Boolean
    Dim child As obj_IUiControl

    For Each child In m_children
        If Not child.Render(diagnostic) Then Exit Function
    Next child
    obj_IUiControl_Render = True
End Function

Private Function obj_IUiControl_Validate(ByVal errors As Collection) As Boolean
    Dim child As obj_IUiControl
    Dim valid As Boolean

    valid = True
    For Each child In m_children
        If Not child.Validate(errors) Then valid = False
    Next child
    obj_IUiControl_Validate = valid
End Function

' //
' // API
' //
Public Sub Dispose()
    If m_isDisposed Then
        Exit Sub
    End If
    m_isDisposed = True
    m_isInitialized = False
    private_ClearChildren
    Set m_children = Nothing
    Set m_sizes = Nothing
    If Not m_base Is Nothing Then m_base.Dispose
    Set m_base = Nothing
    Set m_context = Nothing
    Set m_definition = Nothing
End Sub

' //
' // Private
' //
Private Sub private_ClearChildren()
    Dim child As obj_IUiControl

    If Not m_children Is Nothing Then
        For Each child In m_children
            child.Dispose
        Next child
    End If
    Set m_children = New Collection
    Set m_sizes = New Collection
End Sub