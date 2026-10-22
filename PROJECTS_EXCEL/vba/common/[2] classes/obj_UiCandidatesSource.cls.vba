VERSION 1.0 CLASS
BEGIN
  MultiUse = -1
END
Attribute VB_Name = "obj_UiCandidatesSource"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Implements obj_IUiTableSource
Implements obj_IUiSelectionSource
Implements obj_IUiColumnMetadata
Implements obj_IUiColumnLayout
Implements obj_IUiHeaderPresentation

Private m_isInitialized As Boolean
Private m_isDisposed As Boolean

Private m_table As obj_UiRawTable
Private m_selectionSource As obj_IUiSelectionSource
Private m_sourceAliases As Object
Private m_columnSpans As Object
Private m_rowsPerItem As Long
Private m_detailCells As Collection

' //
' // Lifecycle
' //
Private Sub Class_Initialize()
End Sub

Private Sub Class_Terminate()
    Me.Dispose
End Sub

' //
' // Properties
' //
Private Property Get obj_IUiTableSource_TableCount() As Long
    If m_isInitialized And Not m_isDisposed Then obj_IUiTableSource_TableCount = 1
End Property

' //
' // Interface
' //
Private Function obj_IUiTableSource_GetTable(ByVal index As Long) As Object
    If Not m_isInitialized Or m_isDisposed Or index <> 1 Then Exit Function
    Set obj_IUiTableSource_GetTable = m_table
End Function

Private Function obj_IUiSelectionSource_GetItemAt(ByVal rowIndex As Long) As Object
    If Not m_isInitialized Or m_isDisposed Then Exit Function
    If rowIndex < 1 Or rowIndex > m_table.RowCount Then Exit Function
    Set obj_IUiSelectionSource_GetItemAt = m_selectionSource.GetItemAt((rowIndex - 1) \ m_rowsPerItem + 1)
End Function

Private Function obj_IUiColumnMetadata_SourceAliasAt(ByVal columnIndex As Long) As String
    Dim header As String

    If Not m_isInitialized Or m_isDisposed Then Exit Function
    If m_sourceAliases Is Nothing Then Exit Function
    header = VBA.CStr(m_table.HeaderAt(columnIndex))
    If m_sourceAliases.Exists(header) Then obj_IUiColumnMetadata_SourceAliasAt = m_sourceAliases(header)
End Function

Private Function obj_IUiColumnLayout_ColumnSpanAt(ByVal columnIndex As Long) As Long
    If Not m_isInitialized Or m_isDisposed Then Exit Function
    If columnIndex < 1 Or columnIndex > m_table.ColumnCount Then Exit Function
    obj_IUiColumnLayout_ColumnSpanAt = 1
    If Not m_columnSpans Is Nothing Then
        If m_columnSpans.Exists(columnIndex) Then obj_IUiColumnLayout_ColumnSpanAt = m_columnSpans(columnIndex)
    End If
End Function

Private Function obj_IUiHeaderPresentation_HeaderTextAt(ByVal columnIndex As Long) As String
    Dim rowIndex As Long
    Dim value As Variant
    Dim hasValue As Boolean

    If Not m_isInitialized Or m_isDisposed Then Exit Function
    If columnIndex < 1 Or columnIndex > m_table.ColumnCount Then Exit Function
    For rowIndex = 1 To m_table.RowCount Step m_rowsPerItem
        value = m_table.ValueAt(rowIndex, columnIndex)
        hasValue = False
        If VBA.IsError(value) Then
            hasValue = True
        ElseIf Not VBA.IsNull(value) And Not VBA.IsEmpty(value) Then
            hasValue = (VBA.Len(VBA.Trim$(VBA.CStr(value))) > 0)
        End If
        If hasValue Then
            obj_IUiHeaderPresentation_HeaderTextAt = VBA.CStr(m_table.HeaderAt(columnIndex))
            Exit Function
        End If
    Next rowIndex
End Function

' //
' // API
' //
Public Function Initialize( _
    ByVal table As obj_UiRawTable, _
    ByVal selectionSource As obj_IUiSelectionSource, _
    Optional ByVal sourceAliases As Object, _
    Optional ByVal columnSpans As Object, _
    Optional ByVal rowsPerItem As Long = 1, _
    Optional ByVal detailCells As Collection _
) As Boolean
    If m_isInitialized Or m_isDisposed Then Exit Function
    If table Is Nothing Or selectionSource Is Nothing Then Exit Function
    If rowsPerItem < 1 Then Exit Function
    Set m_table = table
    Set m_selectionSource = selectionSource
    Set m_sourceAliases = sourceAliases
    Set m_columnSpans = columnSpans
    m_rowsPerItem = rowsPerItem
    Set m_detailCells = detailCells
    m_isInitialized = True
    Initialize = True
End Function

Public Sub Dispose()
    If m_isDisposed Then Exit Sub
    m_isDisposed = True
    m_isInitialized = False
    If Not m_table Is Nothing Then m_table.Dispose
    Set m_table = Nothing
    Set m_selectionSource = Nothing
    Set m_sourceAliases = Nothing
    Set m_columnSpans = Nothing
    Set m_detailCells = Nothing
End Sub

Public Sub AppendDetailSourceRanges(ByVal body As Range, ByVal ranges As Object)
    Dim definition As Object
    Dim sourceAlias As String
    Dim rowIndex As Long
    Dim area As Range

    If Not m_isInitialized Or m_isDisposed Then Exit Sub
    If m_detailCells Is Nothing Then Exit Sub
    For Each definition In m_detailCells
        If definition.Exists("sourceAlias") Then
            sourceAlias = CStr(definition("sourceAlias"))
            If Len(sourceAlias) > 0 Then
                For rowIndex = 1 To m_table.RowCount Step m_rowsPerItem
                    Set area = body.Cells(rowIndex + CLng(definition("lane")), CLng(definition("column")))
                    If ranges.Exists(sourceAlias) Then
                        Set ranges(sourceAlias) = Application.Union(ranges(sourceAlias), area)
                    Else
                        ranges.Add sourceAlias, area
                    End If
                Next rowIndex
            End If
        End If
    Next definition
End Sub

Public Sub GetRowGroup( _
    ByVal rowIndex As Long, _
    ByRef firstRow As Long, _
    ByRef rowCount As Long _
)
    If Not m_isInitialized Or m_isDisposed Then Err.Raise 5, , "Candidate source is not initialized."
    If rowIndex < 1 Or rowIndex > m_table.RowCount Then Err.Raise 5, , "Candidate row is outside the table."
    firstRow = ((rowIndex - 1) \ m_rowsPerItem) * m_rowsPerItem + 1
    rowCount = m_rowsPerItem
End Sub