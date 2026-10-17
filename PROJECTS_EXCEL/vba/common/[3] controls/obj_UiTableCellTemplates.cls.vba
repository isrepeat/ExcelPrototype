VERSION 1.0 CLASS
BEGIN
  MultiUse = -1
END
Attribute VB_Name = "obj_UiTableCellTemplates"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Implements obj_IUiEventHandler

Private m_isInitialized As Boolean
Private m_isDisposed As Boolean

Private m_context As obj_UiRenderContext
Private m_entries As Collection
Private m_routes As Object

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
Private Function obj_IUiEventHandler_HandleEvent( _
    ByVal kind As String, _
    ByVal payload As Variant _
) As Boolean
    Dim route As Collection
    Dim command As obj_UiCommand
    Dim item As Object

    If kind <> "click" Or m_isDisposed Then Exit Function
    If Not m_routes.Exists(VBA.CStr(payload)) Then Exit Function
    Set route = m_routes(VBA.CStr(payload))
    Set command = route(1)
    Set item = route(2)
    obj_IUiEventHandler_HandleEvent = command.ExecuteWithPayload(item)
End Function

' //
' // API
' //
Public Function Initialize(ByVal context As obj_UiRenderContext) As Boolean
    If m_isInitialized Or m_isDisposed Or context Is Nothing Then Exit Function
    Set m_context = context
    Set m_entries = New Collection
    Set m_routes = VBA.CreateObject("Scripting.Dictionary")
    m_isInitialized = True
    Initialize = True
End Function

Public Sub Dispose()
    If m_isDisposed Then Exit Sub
    Me.Clear
    m_isDisposed = True
    m_isInitialized = False
    Set m_context = Nothing
End Sub

Public Sub Clear()
    Dim entry As Object
    Dim child As obj_IUiControl
    Dim area As Range
    Dim shape As Shape
    Dim index As Long
    Dim owned As Object

    If m_entries Is Nothing Or m_context Is Nothing Then Exit Sub
    Set owned = VBA.CreateObject("Scripting.Dictionary")
    For Each entry In m_entries
        Set child = entry("Control")
        child.Dispose
        Set area = entry("Area")
        area.NumberFormat = entry("NumberFormat")
        m_context.BindingContext.RemoveValue "TableCells", VBA.CStr(entry("Name"))
        m_context.Styles.ForgetControl VBA.CStr(entry("Name"))
        owned("btn_" & entry("Name")) = True
        owned("chk_" & entry("Name")) = True
    Next entry
    For index = m_context.TargetWorksheet.Shapes.Count To 1 Step -1
        Set shape = m_context.TargetWorksheet.Shapes(index)
        If owned.Exists(shape.Name) Then
            m_context.Router.UnregisterShape shape.Name
            m_context.UnregisterShapeLayout shape.Name
            shape.Delete
        End If
    Next index
    Set m_entries = New Collection
    Set m_routes = VBA.CreateObject("Scripting.Dictionary")
End Sub

Public Function CanRefreshBindings( _
    ByVal table As obj_UiRawTable, _
    ByVal source As obj_IUiTableSource _
) As Boolean
    Dim entry As Object
    Dim currentItem As Object
    Dim previousItem As Object
    Dim child As obj_IUiControl

    If Not m_isInitialized Or m_isDisposed Then
        Exit Function
    End If
    For Each entry In m_entries
        If entry("RowIndex") > table.RowCount Or entry("ColumnIndex") > table.ColumnCount Then
            Exit Function
        End If
        Set currentItem = private_RowItem(table, source, entry("RowIndex"))
        Set previousItem = entry("Item")
        If Not currentItem Is previousItem Then
            Exit Function
        End If
        Set child = entry("Control")
        If Not TypeOf child Is obj_IUiBindingTarget Then
            Exit Function
        End If
    Next entry
    CanRefreshBindings = True
End Function

Public Function TryRefreshBindings( _
    ByVal table As obj_UiRawTable, _
    ByRef diagnostic As String _
) As Boolean
    Dim entry As Object
    Dim child As obj_IUiControl
    Dim target As obj_IUiBindingTarget
    Dim started As Double
    Dim scope As Object
    Dim cell As Object
    Dim value As Variant
    Dim isObject As Boolean

    On Error GoTo EH
    If Not m_isInitialized Or m_isDisposed Then
        Err.Raise 5, , "Cell template renderer is not initialized."
    End If
    started = VBA.Timer
    For Each entry In m_entries
        If Not m_context.BindingContext.TryGetValue("TableCells", VBA.CStr(entry("Name")), value, scope, isObject) Then
            Err.Raise 5, , "Cell template binding scope is missing."
        End If
        If Not isObject Or scope Is Nothing Then
            Err.Raise 5, , "Cell template binding scope is invalid."
        End If
        Set cell = scope("Cell")
        cell("Value") = table.ValueAt(entry("RowIndex"), entry("ColumnIndex"))
        Set child = entry("Control")
        If Not TypeOf child Is obj_IUiBindingTarget Then
            Err.Raise 5, , "Cell template does not support binding updates."
        End If
        Set target = child
        If Not target.RefreshBindings(m_context) Then
            Err.Raise 5, , "Cannot refresh cell template bindings: " & entry("Name")
        End If
    Next entry
    TryRefreshBindings = True
    ex_Core.fn_Diagnostic_WritePerf "Control.Table.RefreshTemplateBindings", started
    Exit Function
EH:
    diagnostic = "Table.RefreshTemplateBindings: " & VBA.Err.Description
End Function

Public Function TryRender( _
    ByVal definition As Object, _
    ByVal table As obj_UiRawTable, _
    ByVal source As obj_IUiTableSource, _
    ByVal body As Range, _
    ByRef diagnostic As String _
) As Boolean
    Dim templates As Object
    Dim template As Object
    Dim used As Object
    Dim columnIndex As Long
    Dim rowIndex As Long
    Dim columnAlias As String
    Dim mode As String
    Dim item As Object
    Dim errorText As String

    On Error GoTo EH
    If Not m_isInitialized Or m_isDisposed Then Err.Raise 5, , "Cell template renderer is not initialized."
    Me.Clear
    Set used = VBA.CreateObject("Scripting.Dictionary")
    Set templates = definition.SelectNodes("*[local-name()='table.cellTemplates']/*[local-name()='cellTemplate']")
    For Each template In templates
        columnAlias = ex_UiElementFactory.fn_Attribute(template, "columnAlias")
        columnIndex = private_FindColumn(table, source, columnAlias)
        If used.Exists(columnIndex) Then Err.Raise 5, , "Duplicate cell template for column: " & columnAlias
        used.Add columnIndex, True
        mode = ex_UiElementFactory.fn_Attribute(template, "mode")
        If VBA.Len(mode) = 0 Then mode = "replace"
        If mode <> "replace" And mode <> "overlay" Then Err.Raise 5, , "Invalid cell template mode: " & mode
        If template.SelectNodes("*").Length <> 1 Then Err.Raise 5, , "Cell template requires exactly one control."
        For rowIndex = 1 To table.RowCount
            If body Is Nothing Then Err.Raise 5, , "Cell template body range is missing."
            Set item = private_RowItem(table, source, rowIndex)
            If item Is Nothing Then Err.Raise 5, , "Cell template row record is missing."
            If Not private_RenderCell(definition, template, item, table.ValueAt(rowIndex, columnIndex), _
                    body.Cells(rowIndex, columnIndex).MergeArea, rowIndex, columnIndex, mode, diagnostic) Then Exit Function
        Next rowIndex
    Next template
    TryRender = True
    Exit Function
EH:
    errorText = VBA.Err.Description
    On Error Resume Next
    Me.Clear
    If VBA.Err.Number <> 0 Then errorText = errorText & " | Cleanup: " & VBA.Err.Description
    On Error GoTo 0
    diagnostic = "Table.CellTemplates: " & errorText
End Function

' //
' // Private
' //
Private Function private_FindColumn( _
    ByVal table As obj_UiRawTable, _
    ByVal source As obj_IUiTableSource, _
    ByVal columnAlias As String _
) As Long
    Dim metadata As obj_IUiColumnAliases
    Dim index As Long
    Dim matches As Boolean

    If TypeOf source Is obj_IUiColumnAliases Then Set metadata = source
    For index = 1 To table.ColumnCount
        matches = (VBA.CStr(table.HeaderAt(index)) = columnAlias)
        If Not metadata Is Nothing Then matches = matches Or (metadata.ColumnAliasAt(index) = columnAlias)
        If matches Then
            If private_FindColumn > 0 Then Err.Raise 5, , "Ambiguous cell template column: " & columnAlias
            private_FindColumn = index
        End If
    Next index
    If private_FindColumn = 0 Then Err.Raise 5, , "Cell template column was not found: " & columnAlias
End Function

Private Function private_RowItem( _
    ByVal table As obj_UiRawTable, _
    ByVal source As obj_IUiTableSource, _
    ByVal rowIndex As Long _
) As Object
    Dim selection As obj_IUiSelectionSource
    Dim records As obj_IUiRowRecords
    Dim metadata As obj_IUiColumnAliases
    Dim item As Object
    Dim index As Long
    Dim columnAlias As String

    If TypeOf source Is obj_IUiRowRecords Then
        Set records = source
        Set item = records.GetRowItem(rowIndex)
        If Not item Is Nothing Then
            Set private_RowItem = item
            Exit Function
        End If
    End If
    If TypeOf source Is obj_IUiSelectionSource Then
        Set selection = source
        Set private_RowItem = selection.GetItemAt(rowIndex)
        Exit Function
    End If
    If TypeOf source Is obj_IUiColumnAliases Then Set metadata = source
    Set item = VBA.CreateObject("Scripting.Dictionary")
    For index = 1 To table.ColumnCount
        columnAlias = VBA.CStr(table.HeaderAt(index))
        If Not item.Exists(columnAlias) Then item.Add columnAlias, table.ValueAt(rowIndex, index)
        If Not metadata Is Nothing Then
            columnAlias = metadata.ColumnAliasAt(index)
            If VBA.Len(columnAlias) > 0 Then
                If Not item.Exists(columnAlias) Then item.Add columnAlias, table.ValueAt(rowIndex, index)
            End If
        End If
    Next index
    Set private_RowItem = item
End Function

Private Function private_RenderCell( _
    ByVal owner As Object, _
    ByVal template As Object, _
    ByVal item As Object, _
    ByVal value As Variant, _
    ByVal area As Range, _
    ByVal rowIndex As Long, _
    ByVal columnIndex As Long, _
    ByVal mode As String, _
    ByRef diagnostic As String _
) As Boolean
    Dim node As Object
    Dim attributeNode As Object
    Dim scope As Object
    Dim cell As Object
    Dim entry As Object
    Dim child As obj_IUiControl
    Dim command As obj_UiCommand
    Dim route As Collection
    Dim name As String
    Dim kind As String
    Dim raw As String
    Dim path As String
    Dim rows As Long
    Dim columns As Long

    name = ex_UiElementFactory.fn_Attribute(owner, "name") & "_r" & VBA.CStr(rowIndex) & "c" & VBA.CStr(columnIndex)
    If VBA.Len("btn_" & name) > 31 Then Err.Raise 5, , "Cell template shape name exceeds the Excel limit: " & name
    Set scope = VBA.CreateObject("Scripting.Dictionary")
    Set cell = VBA.CreateObject("Scripting.Dictionary")
    cell.Add "Value", value
    scope.Add "Item", item
    scope.Add "Cell", cell
    Set node = template.SelectSingleNode("*").cloneNode(True)
    kind = VBA.LCase$(VBA.CStr(node.baseName))
    If kind <> "button" And kind <> "label" And kind <> "checkbox" Then Err.Raise 5, , "Cell templates support button, label and checkbox controls."
    If mode = "overlay" And kind <> "button" And kind <> "checkbox" Then Err.Raise 5, , "Overlay templates require a shape control."
    node.setAttribute "type", kind
    node.setAttribute "name", name
    node.setAttribute "row", "1"
    node.setAttribute "column", "1"
    node.setAttribute "rowSpan", VBA.CStr(area.Rows.Count)
    node.setAttribute "columnSpan", VBA.CStr(area.Columns.Count)
    If kind = "button" And VBA.Len(ex_UiElementFactory.fn_Attribute(node, "shapePlacement")) = 0 Then node.setAttribute "shapePlacement", "moveAndSize"
    For Each attributeNode In node.Attributes
        raw = VBA.Trim$(VBA.CStr(attributeNode.Text))
        If VBA.Left$(raw, 14) = "{Binding Path=" And VBA.Right$(raw, 1) = "}" Then
            path = VBA.Mid$(raw, 15, VBA.Len(raw) - 15)
            If path = "Item" Or VBA.Left$(path, 5) = "Item." Or VBA.Left$(path, 5) = "Cell." Then
                attributeNode.Text = "{Binding Source=TableCells; Path=" & name & "." & path & "}"
            End If
        End If
    Next attributeNode
    Set child = ex_UiControlFactory.fn_Create(node)
    If child Is Nothing Then Err.Raise 5, , "Cannot create cell template control: " & kind
    Set entry = VBA.CreateObject("Scripting.Dictionary")
    entry.Add "Name", name
    entry.Add "Control", child
    entry.Add "Area", area
    entry.Add "NumberFormat", area.NumberFormat
    entry.Add "RowIndex", rowIndex
    entry.Add "ColumnIndex", columnIndex
    entry.Add "Item", item
    m_entries.Add entry
    If Not m_context.BindingContext.SetObject("TableCells", name, scope) Then Err.Raise 5, , "Cannot initialize cell template context."
    If Not child.Configure(node, m_context, VBA.vbNullString, diagnostic) Then Exit Function
    If Not child.Measure(rows, columns, diagnostic) Then Exit Function
    If Not child.Arrange(area.Row, area.Column, diagnostic) Then Exit Function
    If Not child.Render(diagnostic) Then Exit Function
    If kind = "button" Then
        If Not ex_UiBindingRuntime.fn_TryResolveCommand(ex_UiElementFactory.fn_Attribute(node, "command"), _
                m_context.BindingContext, command) Then Exit Function
        Set route = New Collection
        route.Add command
        route.Add item
        m_routes.Add "btn_" & name, route
        m_context.Router.RegisterShape "btn_" & name, Me
    End If
    If mode = "replace" And (kind = "button" Or kind = "checkbox") Then area.NumberFormat = ";;;"
    private_RenderCell = True
End Function