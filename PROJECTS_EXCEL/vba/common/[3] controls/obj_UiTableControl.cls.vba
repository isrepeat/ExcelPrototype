VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "obj_UiTableControl"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Implements obj_IUiControl
Implements obj_IUiTableTarget
Implements obj_IUiEventHandler

Private m_isInitialized As Boolean
Private m_isDisposed As Boolean

Private m_renderContext As obj_UiRenderContext
Private m_uiControlBase As obj_UiControlBase
Private m_source As obj_IUiTableSource
Private m_targetRange As Range
Private m_sourceRaw As String
Private m_table As obj_UiRawTable
Private m_showHeaders As Boolean
Private m_selectionSource As obj_IUiSelectionSource
Private m_selectCommand As obj_UiCommand
Private m_selectedSource As String
Private m_selectedPath As String
Private m_selectedStyle As String
Private m_bodyRange As Range
Private m_selectedRow As Long
Private m_tablePlan As Object
Private m_sourceOverride As obj_IUiTableSource
Private m_selectionMode As String
Private m_renderedSchema As String
Private m_renderedValues As Variant
Private m_renderedSelection As obj_IUiSelectionSource
Private m_cellTemplates As obj_UiTableCellTemplates

' //
' // Lifecycle
' //
Private Sub Class_Initialize()
    Set m_uiControlBase = New obj_UiControlBase
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
    If m_uiControlBase Is Nothing Then Set m_uiControlBase = New obj_UiControlBase
    obj_IUiControl_Initialize = m_uiControlBase.Initialize()
    m_isInitialized = obj_IUiControl_Initialize
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
    Set m_renderContext = context
    m_uiControlBase.ConfigurePosition definition
    obj_IUiControl_Configure = private_Configure(definition)
    If Not obj_IUiControl_Configure Then diagnostic = "Cannot configure control: " & ex_UiElementFactory.fn_Attribute(definition, "name")
End Function

Private Function obj_IUiControl_Measure( _
    ByRef rows As Long, _
    ByRef columns As Long, _
    ByRef diagnostic As String _
) As Boolean
    Dim target As Range

    m_uiControlBase.SetPosition 1, 1
    If Not private_Configure(m_uiControlBase.ControlNode) Then Exit Function
    Set target = private_Measure(m_renderContext)
    If target Is Nothing Then
        diagnostic = "Cannot measure control: " & m_uiControlBase.ControlName
        Exit Function
    End If
    m_uiControlBase.GetSize target, rows, columns
    obj_IUiControl_Measure = True
End Function

Private Function obj_IUiControl_Arrange( _
    ByVal row As Long, _
    ByVal column As Long, _
    ByRef diagnostic As String _
) As Boolean
    Dim target As Range

    m_uiControlBase.ArrangePosition row, column
    If Not private_Configure(m_uiControlBase.ControlNode) Then Exit Function
    Set target = private_Measure(m_renderContext)
    obj_IUiControl_Arrange = Not target Is Nothing
    If Not target Is Nothing Then
        Set m_tablePlan = ex_UiTables.fn_CreatePlan(m_uiControlBase.ControlNode, target, m_table, m_showHeaders)
        m_renderContext.RegisterTablePlan m_tablePlan
    End If
    If target Is Nothing Then diagnostic = "Cannot arrange control: " & m_uiControlBase.ControlName
End Function

Private Function obj_IUiControl_Render(ByRef diagnostic As String) As Boolean
    obj_IUiControl_Render = private_Render(m_renderContext)
    If Not obj_IUiControl_Render Then diagnostic = "Cannot render control: " & m_uiControlBase.ControlName
End Function

Private Function obj_IUiControl_Validate(ByVal errors As Collection) As Boolean
    obj_IUiControl_Validate = True
End Function

Private Function obj_IUiEventHandler_HandleEvent( _
    ByVal kind As String, _
    ByVal payload As Variant _
) As Boolean
    Dim target As Range
    Dim item As Object
    Dim rowIndex As Long
    If kind <> "selection" And kind <> "confirmSelection" Then Exit Function
    ' Confirmation mode ignores navigation without reading records or changing styles.
    If kind = "selection" And m_selectionMode = "Confirm" Then Exit Function
    If m_bodyRange Is Nothing Then Exit Function
    If kind = "confirmSelection" And m_selectionMode <> "Confirm" Then Exit Function
    If m_selectionSource Is Nothing Then Exit Function
    Set target = payload
    If target Is Nothing Then Exit Function
    If Not target.Parent Is m_bodyRange.Parent Then Exit Function
    If Application.Intersect(target, m_bodyRange) Is Nothing Then Exit Function
    rowIndex = target.Row - m_bodyRange.Row + 1
    Set item = m_selectionSource.GetItemAt(rowIndex)
    If item Is Nothing Then Exit Function
    If m_selectedRow > 0 Then
        m_renderContext.Styles.RestoreControlArea m_bodyRange.Rows(m_selectedRow), _
            m_uiControlBase.ControlNode, m_renderContext.BindingContext
    End If
    m_selectedRow = rowIndex
    private_StyleSelectedRow rowIndex
    If VBA.Len(m_selectedSource) > 0 Then
        If Not m_renderContext.BindingContext.TrySetPathObject(m_selectedSource, m_selectedPath, item) Then Exit Function
    End If
    If Not m_selectCommand Is Nothing Then
        If Not m_selectCommand.ExecuteWithPayload(item) Then Exit Function
    End If
    obj_IUiEventHandler_HandleEvent = True
End Function

Private Sub obj_IUiTableTarget_SetTable(ByVal table As obj_UiRawTable)
    Set m_table = table
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
    If Not m_cellTemplates Is Nothing Then m_cellTemplates.Dispose
    Set m_cellTemplates = Nothing
    If Not m_renderContext Is Nothing Then m_renderContext.Router.UnregisterSelection m_uiControlBase.ControlName
    Set m_selectionSource = Nothing
    Set m_selectCommand = Nothing
    Set m_bodyRange = Nothing
    Set m_source = Nothing
    Set m_table = Nothing
    Set m_targetRange = Nothing
    Set m_tablePlan = Nothing
    Set m_sourceOverride = Nothing
    Set m_renderedSelection = Nothing
    m_renderedValues = Empty
    m_renderedSchema = VBA.vbNullString
    If Not m_uiControlBase Is Nothing Then m_uiControlBase.Dispose
    Set m_uiControlBase = Nothing
    Set m_renderContext = Nothing
End Sub

Public Sub SetSource(ByVal source As obj_IUiTableSource)
    If Not m_isInitialized Or m_isDisposed Then Exit Sub
    Set m_sourceOverride = source
    Set m_table = Nothing
End Sub

Public Function TryRefreshValues( _
    ByRef refreshed As Boolean, _
    ByRef diagnostic As String _
) As Boolean
    Dim table As obj_UiRawTable
    Dim selection As obj_IUiSelectionSource
    Dim headerPresentation As obj_IUiHeaderPresentation
    Dim columnLayout As obj_IUiColumnLayout
    Dim item As Object
    Dim previousItem As Object
    Dim buffer As Variant
    Dim columnValues As Variant
    Dim rowIndex As Long
    Dim columnIndex As Long
    Dim targetRow As Long
    Dim changed As Boolean
    Dim startedAt As Double
    Dim selectedValue As Variant
    Dim selected As Object
    Dim selectedIsObject As Boolean

    On Error GoTo EH
    refreshed = False
    TryRefreshValues = True
    If m_uiControlBase.ControlNode.SelectNodes("*[local-name()='table.cellTemplates']").Length > 0 Then Exit Function
    If m_targetRange Is Nothing Or m_sourceOverride Is Nothing Then Exit Function
    If Not IsArray(m_renderedValues) Then Exit Function
    If m_sourceOverride.TableCount <> 1 Then Exit Function
    Set table = m_sourceOverride.GetTable(1)
    If table Is Nothing Then Exit Function
    If private_SourceSchema(m_sourceOverride, table) <> m_renderedSchema Then Exit Function
    If Not TypeOf m_sourceOverride Is obj_IUiSelectionSource Then Exit Function
    If m_renderedSelection Is Nothing Then Exit Function
    Set selection = m_sourceOverride
    For rowIndex = 1 To table.RowCount
        Set item = selection.GetItemAt(rowIndex)
        Set previousItem = m_renderedSelection.GetItemAt(rowIndex)
        If Not item Is previousItem Then Exit Function
    Next rowIndex
    If VBA.Len(m_selectedSource) > 0 Then
        If Not m_renderContext.BindingContext.TryGetValue( _
                m_selectedSource, m_selectedPath, selectedValue, selected, selectedIsObject) Then
            VBA.Err.Raise 5, , "Cannot resolve the current table selection."
        End If
        If selectedIsObject Then
            If m_selectedRow <= 0 Then Exit Function
            Set item = selection.GetItemAt(m_selectedRow)
            If Not item Is selected Then Exit Function
        ElseIf m_selectedRow > 0 Then
            Exit Function
        End If
    End If

    startedAt = VBA.Timer
    ReDim buffer(1 To m_targetRange.Rows.Count, 1 To m_targetRange.Columns.Count)
    targetRow = 1
    If VBA.Len(table.Title) > 0 Then
        buffer(targetRow, 1) = table.Title
        targetRow = targetRow + 1
    End If
    If TypeOf m_sourceOverride Is obj_IUiHeaderPresentation Then Set headerPresentation = m_sourceOverride
    If TypeOf m_sourceOverride Is obj_IUiColumnLayout Then Set columnLayout = m_sourceOverride
    If m_showHeaders And IsArray(table.Headers) Then
        For columnIndex = 1 To table.ColumnCount
            If headerPresentation Is Nothing Then
                buffer(targetRow, columnIndex) = table.HeaderAt(columnIndex)
            Else
                buffer(targetRow, columnIndex) = headerPresentation.HeaderTextAt(columnIndex)
            End If
        Next columnIndex
        targetRow = targetRow + 1
    End If
    For rowIndex = 1 To table.RowCount
        For columnIndex = 1 To table.ColumnCount
            buffer(targetRow, columnIndex) = table.ValueAt(rowIndex, columnIndex)
        Next columnIndex
        targetRow = targetRow + 1
    Next rowIndex

    ' Keep styles, merged cells and selection routes; write changed columns only.
    For columnIndex = 1 To table.ColumnCount
        If Not columnLayout Is Nothing Then
            If columnLayout.ColumnSpanAt(columnIndex) = 0 Then GoTo NextValueColumn
        End If
        changed = False
        For rowIndex = 1 To UBound(buffer, 1)
            If Not private_SameCellValue(buffer(rowIndex, columnIndex), m_renderedValues(rowIndex, columnIndex)) Then
                changed = True
                Exit For
            End If
        Next rowIndex
        If changed Then
            If private_SourceColumnSpan(m_sourceOverride, columnIndex) > 1 Then
                For rowIndex = 1 To UBound(buffer, 1)
                    If Not private_SameCellValue(buffer(rowIndex, columnIndex), m_renderedValues(rowIndex, columnIndex)) Then
                        m_targetRange.Cells(rowIndex, columnIndex).Value2 = buffer(rowIndex, columnIndex)
                    End If
                Next rowIndex
            Else
                ReDim columnValues(1 To UBound(buffer, 1), 1 To 1)
                For rowIndex = 1 To UBound(buffer, 1)
                    columnValues(rowIndex, 1) = buffer(rowIndex, columnIndex)
                Next rowIndex
                m_targetRange.Columns(columnIndex).Value2 = columnValues
            End If
        End If
NextValueColumn:
    Next columnIndex
    Set m_source = m_sourceOverride
    Set m_table = table
    Set m_selectionSource = selection
    Set m_renderedSelection = selection
    m_renderedValues = buffer
    refreshed = True
    ex_Core.fn_Diagnostic_WritePerf "Control.Table.RefreshValues", startedAt
    Exit Function
EH:
    diagnostic = "Table.RefreshValues: " & VBA.Err.Description
    TryRefreshValues = False
End Function

' //
' // Private
' //
Private Function private_Configure(ByVal controlNode As Object) As Boolean
    Dim modeRaw As String

    If Not m_uiControlBase.Configure(controlNode) Then Exit Function
    modeRaw = private_ReadAttribute(controlNode, "selectionMode")
    m_selectionMode = "Focus"
    If VBA.Len(modeRaw) > 0 Then
        If Not ex_UiBindingRuntime.fn_TryResolveText(modeRaw, m_renderContext.BindingContext, m_selectionMode) Then Exit Function
    End If
    If m_selectionMode <> "Focus" And m_selectionMode <> "Confirm" Then
        Err.Raise 5, , "Invalid table selectionMode: " & m_selectionMode
    End If
    Set m_source = Nothing
    If VBA.Len(m_sourceRaw) > 0 Then Set m_table = Nothing
    m_sourceRaw = private_ReadAttribute(controlNode, "itemsSource")
    If VBA.Len(m_sourceRaw) = 0 And m_table Is Nothing Then Exit Function
    m_showHeaders = private_ReadBoolean(controlNode, "showHeaders", True)
    m_selectedStyle = private_ReadAttribute(controlNode, "selectedStyle")
    m_selectedSource = VBA.vbNullString
    m_selectedPath = VBA.vbNullString
    If VBA.Len(private_ReadAttribute(controlNode, "selectedItem")) > 0 Then
        If Not ex_UiBindingRuntime.fn_TryParseBinding(private_ReadAttribute(controlNode, "selectedItem"), _
                VBA.vbNullString, m_selectedSource, m_selectedPath, m_renderContext.BindingContext) Then Exit Function
        If VBA.Len(m_selectedPath) = 0 Then Exit Function
    End If
    Set m_selectCommand = Nothing
    If VBA.Len(private_ReadAttribute(controlNode, "onSelect")) > 0 Then
        If Not ex_UiBindingRuntime.fn_TryResolveCommand(private_ReadAttribute(controlNode, "onSelect"), _
                m_renderContext.BindingContext, m_selectCommand) Then Exit Function
    End If
    private_Configure = True
End Function

Private Function private_Measure(ByVal uiRenderContext As obj_UiRenderContext) As Range
    Dim tableIndex As Long, rows As Long, columns As Long
    Dim rawTable As obj_UiRawTable
    Dim value As Variant, sourceObject As Object, isObject As Boolean

    If uiRenderContext Is Nothing Then Exit Function
    If m_table Is Nothing Then
        If Not m_sourceOverride Is Nothing Then
            Set sourceObject = m_sourceOverride
            isObject = True
        Else
            If Not ex_UiBindingRuntime.fn_TryResolveValue( _
                    m_sourceRaw, uiRenderContext.BindingContext, value, sourceObject, isObject) Then Exit Function
        End If
        If Not isObject Or Not TypeOf sourceObject Is obj_IUiTableSource Then Exit Function
        Set m_source = sourceObject
        Set m_selectionSource = Nothing
        If TypeOf sourceObject Is obj_IUiSelectionSource Then Set m_selectionSource = sourceObject
        If Not m_selectCommand Is Nothing Or VBA.Len(m_selectedSource) > 0 Then
            If m_selectionSource Is Nothing Then
                ex_WindowsUi.fn_ShowMessage "Selectable table requires obj_IUiSelectionSource: " & m_uiControlBase.ControlName, vbExclamation, "Table"
                Exit Function
            End If
        End If
        If m_source.TableCount <> 1 Then Exit Function
        Set m_table = m_source.GetTable(1)
    End If
    If m_table Is Nothing Then Exit Function
    Set rawTable = m_table
    rows = rawTable.RowCount
    If rows = 0 And private_ReadAttribute(m_uiControlBase.ControlNode, "representation") = "Smart" Then rows = 1
    If VBA.Len(rawTable.Title) > 0 Then rows = rows + 1
    If m_showHeaders And IsArray(rawTable.Headers) Then rows = rows + 1
    columns = rawTable.ColumnCount
    If columns <= 0 Then Exit Function
    If private_ReadAttribute(m_uiControlBase.ControlNode, "representation") = "Smart" Then
        For tableIndex = 1 To columns
            If private_ColumnSpanAt(tableIndex) <> 1 Then
                Err.Raise 5, , "Smart tables do not support column spans"
            End If
        Next tableIndex
    End If
    If rows <= 0 Then rows = 1
    Set m_targetRange = uiRenderContext.TargetWorksheet.Cells(1, 1).Offset( _
        private_ReadLong(m_uiControlBase.ControlNode, "row", 1) - 1, _
        private_ReadLong(m_uiControlBase.ControlNode, "column", 1) - 1).Resize(rows, columns)
    Set private_Measure = m_targetRange
End Function

Private Function private_Render(ByVal uiRenderContext As obj_UiRenderContext) As Boolean
    Dim startedAt As Double
    Dim tableIndex As Long, rowIndex As Long, columnIndex As Long, targetRow As Long
    Dim rawTable As obj_UiRawTable, buffer As Variant
    Dim selected As Object
    Dim item As Object
    Dim selectedValue As Variant
    Dim selectedIsObject As Boolean
    Dim rowTags As String
    Dim columnMetadata As obj_IUiColumnMetadata
    Dim sourceAlias As String
    Dim columnSpan As Long
    Dim fieldRange As Range
    Dim columnLayout As obj_IUiColumnLayout
    Dim headerPresentation As obj_IUiHeaderPresentation
    Dim templateDiagnostic As String

    startedAt = VBA.Timer
    If m_targetRange Is Nothing Then Set m_targetRange = private_Measure(uiRenderContext)
    If m_targetRange Is Nothing Then Exit Function
    If m_table Is Nothing Then
        ex_Core.fn_Diagnostic_WriteLog "UI_TABLE_RENDER_SKIPPED_NO_SOURCE | Source=" & m_sourceRaw
        Exit Function
    End If
    If Not m_cellTemplates Is Nothing Then m_cellTemplates.Clear
    ReDim buffer(1 To m_targetRange.Rows.Count, 1 To m_targetRange.Columns.Count)
    targetRow = 1
    Set rawTable = m_table
    If Not m_source Is Nothing Then
        If TypeOf m_source Is obj_IUiHeaderPresentation Then Set headerPresentation = m_source
    End If
        If VBA.Len(rawTable.Title) > 0 Then buffer(targetRow, 1) = rawTable.Title: targetRow = targetRow + 1
        If m_showHeaders And IsArray(rawTable.Headers) Then
            For columnIndex = 1 To rawTable.ColumnCount
                If headerPresentation Is Nothing Then
                    buffer(targetRow, columnIndex) = rawTable.HeaderAt(columnIndex)
                Else
                    buffer(targetRow, columnIndex) = headerPresentation.HeaderTextAt(columnIndex)
                End If
            Next columnIndex
            targetRow = targetRow + 1
        End If
        For rowIndex = 1 To rawTable.RowCount
            For columnIndex = 1 To rawTable.ColumnCount: buffer(targetRow, columnIndex) = rawTable.ValueAt(rowIndex, columnIndex): Next columnIndex
            targetRow = targetRow + 1
        Next rowIndex
    uiRenderContext.Router.UnregisterSelection m_uiControlBase.ControlName
    Set m_bodyRange = Nothing
    m_selectedRow = 0
    If rawTable.RowCount > 0 Then
        targetRow = 1
        If VBA.Len(rawTable.Title) > 0 Then targetRow = targetRow + 1
        If m_showHeaders And IsArray(rawTable.Headers) Then targetRow = targetRow + 1
        Set m_bodyRange = m_targetRange.Cells(targetRow, 1).Resize(rawTable.RowCount, rawTable.ColumnCount)
        If Not m_selectionSource Is Nothing Then
            uiRenderContext.Router.RegisterSelection m_uiControlBase.ControlName, m_bodyRange, Me
        End If
    End If
    If Not m_source Is Nothing Then
        If TypeOf m_source Is obj_IUiColumnLayout Then Set columnLayout = m_source
    End If
    If Not columnLayout Is Nothing Then
        m_targetRange.UnMerge
    End If
    m_targetRange.Value2 = buffer
    If Not columnLayout Is Nothing Then
        targetRow = 1
        If VBA.Len(rawTable.Title) > 0 Then targetRow = targetRow + 1
        For columnIndex = 1 To rawTable.ColumnCount
            columnSpan = columnLayout.ColumnSpanAt(columnIndex)
            If columnSpan > 1 Then
                If columnIndex + columnSpan - 1 > rawTable.ColumnCount Then Err.Raise 5, , "Table columnSpan exceeds table width"
                For rowIndex = targetRow To m_targetRange.Rows.Count
                    m_targetRange.Cells(rowIndex, columnIndex).Resize(1, columnSpan).Merge
                Next rowIndex
            End If
        Next columnIndex
    End If
    If Not m_tablePlan Is Nothing Then ex_UiTables.fn_Apply m_tablePlan
    uiRenderContext.Styles.ApplyControlStyle m_targetRange, Nothing, m_uiControlBase.ControlNode, uiRenderContext.BindingContext
    targetRow = 1
    If VBA.Len(rawTable.Title) > 0 Then
        uiRenderContext.Styles.RegisterPart m_uiControlBase.ControlNode, m_targetRange.Rows(targetRow), "section"
        targetRow = targetRow + 1
    End If
    If m_showHeaders And IsArray(rawTable.Headers) Then
        uiRenderContext.Styles.RegisterPart m_uiControlBase.ControlNode, m_targetRange.Rows(targetRow), "header"
        For columnIndex = 1 To rawTable.ColumnCount
            columnSpan = private_ColumnSpanAt(columnIndex)
            If columnSpan > 0 Then
                Set fieldRange = m_targetRange.Cells(targetRow, columnIndex).Resize(1, columnSpan)
                uiRenderContext.Styles.RegisterPart m_uiControlBase.ControlNode, fieldRange, "header", rawTable.HeaderAt(columnIndex)
            End If
        Next columnIndex
    End If
    If Not m_bodyRange Is Nothing Then
        uiRenderContext.Styles.RegisterPart m_uiControlBase.ControlNode, m_bodyRange, "rows"
        For rowIndex = 1 To rawTable.RowCount
            rowTags = "tableRow odd"
            If rowIndex Mod 2 = 0 Then rowTags = "tableRow even"
            uiRenderContext.Styles.RegisterPart m_uiControlBase.ControlNode, m_bodyRange.Rows(rowIndex), "row", tags:=rowTags
        Next rowIndex
        For columnIndex = 1 To rawTable.ColumnCount
            columnSpan = private_ColumnSpanAt(columnIndex)
            If columnSpan > 0 Then
                Set fieldRange = m_bodyRange.Columns(columnIndex).Resize(rawTable.RowCount, columnSpan)
                uiRenderContext.Styles.RegisterPart m_uiControlBase.ControlNode, fieldRange, "rows", rawTable.HeaderAt(columnIndex)
                If Not m_source Is Nothing Then
                    If TypeOf m_source Is obj_IUiColumnMetadata Then
                        Set columnMetadata = m_source
                        sourceAlias = columnMetadata.SourceAliasAt(columnIndex)
                        If VBA.Len(sourceAlias) > 0 Then
                            uiRenderContext.Styles.RegisterPart m_uiControlBase.ControlNode, fieldRange, _
                                "rows", sourceAlias:=sourceAlias
                            For rowIndex = 1 To rawTable.RowCount
                                rowTags = "tableRow odd"
                                If rowIndex Mod 2 = 0 Then rowTags = "tableRow even"
                                uiRenderContext.Styles.RegisterPart m_uiControlBase.ControlNode, m_bodyRange.Cells(rowIndex, columnIndex).Resize(1, columnSpan), _
                                    "rows", sourceAlias:=sourceAlias, tags:=rowTags
                            Next rowIndex
                        End If
                    End If
                End If
            End If
        Next columnIndex
    End If
    If Not m_selectionSource Is Nothing And Not m_bodyRange Is Nothing And VBA.Len(m_selectedSource) > 0 Then
        If uiRenderContext.BindingContext.TryGetValue(m_selectedSource, m_selectedPath, _
                selectedValue, selected, selectedIsObject) Then
            If selectedIsObject Then
                For rowIndex = 1 To rawTable.RowCount
                    Set item = m_selectionSource.GetItemAt(rowIndex)
                    If item Is selected Then
                        m_selectedRow = rowIndex
                        private_StyleSelectedRow rowIndex
                        Exit For
                    End If
                Next rowIndex
            End If
        End If
    End If
    m_renderedSchema = private_SourceSchema(m_source, rawTable)
    m_renderedValues = buffer
    Set m_renderedSelection = m_selectionSource
    If m_uiControlBase.ControlNode.SelectNodes("*[local-name()='table.cellTemplates']").Length > 0 Then
        If m_cellTemplates Is Nothing Then
            Set m_cellTemplates = New obj_UiTableCellTemplates
            If Not m_cellTemplates.Initialize(uiRenderContext) Then Err.Raise 5, , "Cannot initialize table cell templates."
        End If
        If Not m_cellTemplates.TryRender(m_uiControlBase.ControlNode, rawTable, m_source, m_bodyRange, templateDiagnostic) Then
            Err.Raise 5, , templateDiagnostic
        End If
    End If
    private_Render = True
    ex_Core.fn_Diagnostic_WritePerf "Control.Table.Render", startedAt
End Function

' //
' // Private
' //
Private Function private_SourceSchema( _
    ByVal source As obj_IUiTableSource, _
    ByVal table As obj_UiRawTable _
) As String
    Dim columnIndex As Long
    Dim metadata As obj_IUiColumnMetadata
    Dim key As String
    Dim name As String

    key = VBA.CStr(table.RowCount) & "|" & VBA.CStr(table.ColumnCount) & "|" & VBA.CStr(m_showHeaders)
    key = key & "|" & VBA.CStr(VBA.Len(table.Title)) & ":" & table.Title
    If Not source Is Nothing Then
        If TypeOf source Is obj_IUiColumnMetadata Then Set metadata = source
    End If
    For columnIndex = 1 To table.ColumnCount
        name = VBA.CStr(table.HeaderAt(columnIndex))
        key = key & "|" & VBA.CStr(VBA.Len(name)) & ":" & name
        key = key & "|" & VBA.CStr(private_SourceColumnSpan(source, columnIndex))
        If Not metadata Is Nothing Then
            name = metadata.SourceAliasAt(columnIndex)
            key = key & "|" & VBA.CStr(VBA.Len(name)) & ":" & name
        End If
    Next columnIndex
    private_SourceSchema = key
End Function

Private Function private_SourceColumnSpan( _
    ByVal source As obj_IUiTableSource, _
    ByVal columnIndex As Long _
) As Long
    Dim layout As obj_IUiColumnLayout

    private_SourceColumnSpan = 1
    If source Is Nothing Then Exit Function
    If TypeOf source Is obj_IUiColumnLayout Then
        Set layout = source
        private_SourceColumnSpan = layout.ColumnSpanAt(columnIndex)
    End If
End Function

Private Function private_SameCellValue(ByVal first As Variant, ByVal second As Variant) As Boolean
    If VBA.IsError(first) Or VBA.IsError(second) Then Exit Function
    If VBA.VarType(first) <> VBA.VarType(second) Then Exit Function
    If VBA.IsNull(first) Then
        private_SameCellValue = True
    Else
        private_SameCellValue = (first = second)
    End If
End Function

Private Function private_ColumnSpanAt(ByVal columnIndex As Long) As Long
    Dim columnLayout As obj_IUiColumnLayout

    private_ColumnSpanAt = 1
    If m_source Is Nothing Then Exit Function
    If TypeOf m_source Is obj_IUiColumnLayout Then
        Set columnLayout = m_source
        private_ColumnSpanAt = columnLayout.ColumnSpanAt(columnIndex)
    End If
End Function

Private Sub private_StyleSelectedRow(ByVal rowIndex As Long)
    Dim node As Object

    If VBA.Len(m_selectedStyle) = 0 Or m_bodyRange Is Nothing Then Exit Sub
    Set node = m_uiControlBase.ControlNode.cloneNode(False)
    node.setAttribute "style", m_selectedStyle
    m_renderContext.Styles.ApplyOverlay m_bodyRange.Rows(rowIndex), node, m_renderContext.BindingContext
End Sub

Private Function private_ReadAttribute(ByVal node As Object, ByVal name As String) As String
    Dim value As Variant: value = node.getAttribute(name)

    If Not VBA.IsNull(value) And Not VBA.IsEmpty(value) Then private_ReadAttribute = VBA.CStr(value)
End Function

Private Function private_ReadLong( _
    ByVal node As Object, _
    ByVal name As String, _
    ByVal defaultValue As Long _
) As Long
    Dim value As String: value = private_ReadAttribute(node, name)

    If VBA.IsNumeric(value) Then private_ReadLong = VBA.CLng(value) Else private_ReadLong = defaultValue
End Function

Private Function private_ReadBoolean( _
    ByVal node As Object, _
    ByVal name As String, _
    ByVal defaultValue As Boolean _
) As Boolean
    Dim value As String: value = VBA.LCase$(private_ReadAttribute(node, name))

    If value = "true" Then
        private_ReadBoolean = True
    ElseIf value = "false" Then
        private_ReadBoolean = False
    Else
        private_ReadBoolean = defaultValue
    End If
End Function