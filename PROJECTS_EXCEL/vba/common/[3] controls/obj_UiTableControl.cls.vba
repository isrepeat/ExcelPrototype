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

Private m_uiControlBase As obj_UiControlBase
Private m_source As obj_IUiTableSource
Private m_targetRange As Range
Private m_sourceRaw As String
Private m_gapRows As Long
Private m_showHeaders As Boolean
Private m_isDisposed As Boolean

' //
' // Lifecycle
' //
Private Sub Class_Initialize()
    Set m_uiControlBase = New obj_UiControlBase
End Sub

Private Sub Class_Terminate()
    obj_IUiControl_Dispose
End Sub

' //
' // Interface
' //
Private Function obj_IUiControl_Initialize() As Boolean
    m_isDisposed = False
    If m_uiControlBase Is Nothing Then Set m_uiControlBase = New obj_UiControlBase
    obj_IUiControl_Initialize = m_uiControlBase.Initialize()
End Function

Private Sub obj_IUiControl_Dispose()
    If m_isDisposed Then Exit Sub
    m_isDisposed = True
    Set m_source = Nothing
    Set m_targetRange = Nothing
    If Not m_uiControlBase Is Nothing Then m_uiControlBase.Dispose
    Set m_uiControlBase = Nothing
End Sub

Private Function obj_IUiControl_Configure(ByVal controlNode As Object) As Boolean
    If Not m_uiControlBase.Configure(controlNode) Then Exit Function
    Set m_source = Nothing
    m_sourceRaw = private_ReadAttribute(controlNode, "source")
    If VBA.Len(m_sourceRaw) = 0 Then m_sourceRaw = private_ReadAttribute(controlNode, "itemsSource")
    If VBA.Len(m_sourceRaw) = 0 Then Exit Function
    m_gapRows = private_ReadLong(controlNode, "gapRows", 1)
    m_showHeaders = private_ReadBoolean(controlNode, "showHeaders", True)
    obj_IUiControl_Configure = True
End Function

Private Function obj_IUiControl_Measure(ByVal uiRenderContext As obj_UiRenderContext) As Range
    Dim tableIndex As Long, rows As Long, columns As Long
    Dim rawTable As obj_UiRawTable
    Dim value As Variant, sourceObject As Object, isObject As Boolean

    If uiRenderContext Is Nothing Then Exit Function
    If m_source Is Nothing Then
        If Not ex_UiBindingRuntime.fn_TryResolveValue( _
                m_sourceRaw, uiRenderContext.BindingContext, value, sourceObject, isObject) Then Exit Function
        If Not isObject Or Not TypeOf sourceObject Is obj_IUiTableSource Then Exit Function
        Set m_source = sourceObject
    End If
    If m_source.TableCount = 0 Then
        Set m_targetRange = uiRenderContext.TargetWorksheet.Cells(1, 1).Offset( _
            private_ReadLong(m_uiControlBase.ControlNode, "row", 1) - 1, _
            private_ReadLong(m_uiControlBase.ControlNode, "column", 1) - 1)
        Set obj_IUiControl_Measure = m_targetRange
        Exit Function
    End If
    For tableIndex = 1 To m_source.TableCount
        Set rawTable = m_source.GetTable(tableIndex)
        If rawTable Is Nothing Then Exit Function
        rows = rows + rawTable.RowCount
        If VBA.Len(rawTable.Title) > 0 Then rows = rows + 1
        If m_showHeaders And IsArray(rawTable.Headers) Then rows = rows + 1
        If tableIndex < m_source.TableCount Then rows = rows + m_gapRows
        If rawTable.ColumnCount > columns Then columns = rawTable.ColumnCount
    Next tableIndex
    If rows <= 0 Or columns <= 0 Then Exit Function
    Set m_targetRange = uiRenderContext.TargetWorksheet.Cells(1, 1).Offset( _
        private_ReadLong(m_uiControlBase.ControlNode, "row", 1) - 1, _
        private_ReadLong(m_uiControlBase.ControlNode, "column", 1) - 1).Resize(rows, columns)
    Set obj_IUiControl_Measure = m_targetRange
End Function

Private Function obj_IUiControl_Render(ByVal uiRenderContext As obj_UiRenderContext) As Boolean
    Dim startedAt As Double
    Dim tableIndex As Long, rowIndex As Long, columnIndex As Long, targetRow As Long
    Dim rawTable As obj_UiRawTable, buffer As Variant

    startedAt = VBA.Timer
    If m_targetRange Is Nothing Then Set m_targetRange = obj_IUiControl_Measure(uiRenderContext)
    If m_targetRange Is Nothing Then Exit Function
    If m_source Is Nothing Then
        ex_Core.fn_Diagnostic_WriteLog "UI_TABLE_RENDER_SKIPPED_NO_SOURCE | Source=" & m_sourceRaw
        Exit Function
    End If
    ReDim buffer(1 To m_targetRange.Rows.Count, 1 To m_targetRange.Columns.Count)
    targetRow = 1
    For tableIndex = 1 To m_source.TableCount
        Set rawTable = m_source.GetTable(tableIndex)
        If VBA.Len(rawTable.Title) > 0 Then buffer(targetRow, 1) = rawTable.Title: targetRow = targetRow + 1
        If m_showHeaders And IsArray(rawTable.Headers) Then
            For columnIndex = 1 To rawTable.ColumnCount: buffer(targetRow, columnIndex) = rawTable.HeaderAt(columnIndex): Next columnIndex
            targetRow = targetRow + 1
        End If
        For rowIndex = 1 To rawTable.RowCount
            For columnIndex = 1 To rawTable.ColumnCount: buffer(targetRow, columnIndex) = rawTable.ValueAt(rowIndex, columnIndex): Next columnIndex
            targetRow = targetRow + 1
        Next rowIndex
        targetRow = targetRow + m_gapRows
    Next tableIndex
    m_targetRange.ClearContents
    m_targetRange.Value2 = buffer
    obj_IUiControl_Render = True
    ex_Core.fn_Diagnostic_WritePerf "Control.Table.Render", startedAt
End Function

Private Function obj_IUiControl_HandleCellChange(ByVal target As Range) As Boolean
End Function

' //
' // Private
' //
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