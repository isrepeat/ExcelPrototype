VERSION 1.0 CLASS
BEGIN
  MultiUse = -1
END
Attribute VB_Name = "obj_UiItemsControl"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Implements obj_IUiControl

Private Const ITEM_BINDING_PREFIX As String = "{Binding Path=Item."

Private m_isInitialized As Boolean
Private m_isDisposed As Boolean

Private m_context As obj_UiRenderContext
Private m_definition As Object
Private m_children As Collection
Private m_sizes As Collection
Private m_widths As Collection
Private m_areas As Collection
Private m_names As Collection
Private m_row As Long
Private m_column As Long
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
    obj_IUiControl_Initialize = Me.Initialize()
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
    Set m_context = context
    Set m_definition = definition.cloneNode(True)
    m_row = ex_UiElementFactory.fn_Long(definition, "row", 1)
    m_column = ex_UiElementFactory.fn_Long(definition, "column", 1)
    m_gapRows = 0
    If VBA.Len(ex_UiElementFactory.fn_Attribute(definition, "gapRows")) > 0 Then
        m_gapRows = VBA.CLng(ex_UiElementFactory.fn_Attribute(definition, "gapRows"))
        If m_gapRows < 0 Then Err.Raise 5, , "Items gapRows must be nonnegative."
    End If
    If definition.SelectNodes("*").Length <> 1 Then
        diagnostic = "Items requires exactly one control template."
        Exit Function
    End If
    obj_IUiControl_Configure = True
End Function

Private Function obj_IUiControl_Measure( _
    ByRef rows As Long, _
    ByRef columns As Long, _
    ByRef diagnostic As String _
) As Boolean
    Dim value As Variant
    Dim source As Object
    Dim isObject As Boolean
    Dim items As Collection
    Dim item As Object
    Dim definition As Object
    Dim attributeNode As Object
    Dim child As obj_IUiControl
    Dim index As Long
    Dim childRows As Long
    Dim childColumns As Long
    Dim name As String
    Dim raw As String
    Dim field As String

    On Error GoTo EH
    private_ClearChildren
    raw = ex_UiElementFactory.fn_Attribute(m_definition, "itemsSource")
    If Not ex_UiBindingRuntime.fn_TryResolveValue(raw, m_context.BindingContext, value, source, isObject) Then
        Err.Raise 5, , "Cannot resolve Items source: " & raw
    End If
    If Not isObject Then Err.Raise 5, , "Items requires a Collection of records."
    If Not TypeOf source Is Collection Then Err.Raise 5, , "Items requires a Collection of records."
    Set items = source
    rows = 0
    columns = 0
    For index = 1 To items.Count
        Set item = items(index)
        Set definition = m_definition.SelectSingleNode("*").cloneNode(True)
        name = ex_UiElementFactory.fn_Attribute(m_definition, "name") & "_" & VBA.CStr(index)
        If VBA.Len("btn_" & name) > 31 Then Err.Raise 5, , "Repeated control name exceeds the Excel shape limit: " & name
        definition.setAttribute "name", name
        definition.setAttribute "type", VBA.CStr(definition.baseName)
        definition.setAttribute "row", "1"
        definition.setAttribute "column", "1"
        For Each attributeNode In definition.Attributes
            raw = VBA.Trim$(VBA.CStr(attributeNode.Text))
            If VBA.StrComp(VBA.Left$(raw, VBA.Len(ITEM_BINDING_PREFIX)), ITEM_BINDING_PREFIX, VBA.vbTextCompare) = 0 Then
                field = VBA.Mid$(raw, VBA.Len(ITEM_BINDING_PREFIX) + 1, VBA.Len(raw) - VBA.Len(ITEM_BINDING_PREFIX) - 1)
                If Not item.Exists(field) Then Err.Raise 5, , "Repeated item field is missing: " & field
                attributeNode.Text = VBA.CStr(item(field))
            End If
        Next attributeNode
        Set child = ex_UiControlFactory.fn_Create(definition)
        If child Is Nothing Then Err.Raise 5, , "Cannot create repeated control: " & name
        m_children.Add child
        m_names.Add name
        If Not child.Configure(definition, m_context, VBA.vbNullString, diagnostic) Then Exit Function
        If Not child.Measure(childRows, childColumns, diagnostic) Then Exit Function
        m_sizes.Add childRows
        m_widths.Add childColumns
        rows = rows + childRows
        If index < items.Count Then rows = rows + m_gapRows
        If childColumns > columns Then columns = childColumns
    Next index
    rows = rows + m_row - 1
    columns = columns + m_column - 1
    obj_IUiControl_Measure = True
    Exit Function
EH:
    diagnostic = VBA.Err.Description
End Function

Private Function obj_IUiControl_Arrange( _
    ByVal row As Long, _
    ByVal column As Long, _
    ByRef diagnostic As String _
) As Boolean
    Dim index As Long
    Dim child As obj_IUiControl

    row = row + m_row - 1
    column = column + m_column - 1
    Set m_areas = New Collection
    For index = 1 To m_children.Count
        Set child = m_children(index)
        If Not child.Arrange(row, column, diagnostic) Then Exit Function
        m_areas.Add m_context.TargetWorksheet.Cells(row, column).Resize(VBA.CLng(m_sizes(index)), VBA.CLng(m_widths(index)))
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

    obj_IUiControl_Validate = True
    For Each child In m_children
        If Not child.Validate(errors) Then obj_IUiControl_Validate = False
    Next child
End Function

' //
' // API
' //
Public Function Initialize() As Boolean
    If m_isInitialized Or m_isDisposed Then Exit Function
    Set m_children = New Collection
    Set m_sizes = New Collection
    Set m_widths = New Collection
    Set m_areas = New Collection
    Set m_names = New Collection
    m_isInitialized = True
    Initialize = True
End Function

Public Sub Dispose()
    If m_isDisposed Then Exit Sub
    m_isDisposed = True
    m_isInitialized = False
    private_ClearChildren
    Set m_context = Nothing
    Set m_definition = Nothing
End Sub

Public Function TryGetItemArea( _
    ByVal index As Long, _
    ByRef area As Range _
) As Boolean
    Set area = Nothing
    If Not m_isInitialized Or m_isDisposed Then Exit Function
    If index < 1 Or index > m_areas.Count Then Exit Function
    Set area = m_areas(index)
    TryGetItemArea = True
End Function

' //
' // Private
' //
Private Sub private_ClearChildren()
    Dim child As obj_IUiControl
    Dim name As Variant
    Dim index As Long
    Dim shape As Shape
    Dim owned As Object

    Set owned = VBA.CreateObject("Scripting.Dictionary")
    If Not m_names Is Nothing And Not m_context Is Nothing Then
        For Each name In m_names
            owned("btn_" & name) = True
            m_context.Router.UnregisterShape "btn_" & name
            m_context.UnregisterShapeLayout "btn_" & name
            m_context.Styles.ForgetControl VBA.CStr(name)
        Next name
        For index = m_context.TargetWorksheet.Shapes.Count To 1 Step -1
            Set shape = m_context.TargetWorksheet.Shapes(index)
            If owned.Exists(shape.Name) Then shape.Delete
        Next index
    End If
    If Not m_children Is Nothing Then
        For Each child In m_children
            child.Dispose
        Next child
    End If
    Set m_children = New Collection
    Set m_sizes = New Collection
    Set m_widths = New Collection
    Set m_areas = New Collection
    Set m_names = New Collection
End Sub