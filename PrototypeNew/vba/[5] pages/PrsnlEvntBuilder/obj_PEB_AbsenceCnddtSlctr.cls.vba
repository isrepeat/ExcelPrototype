VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "obj_PEB_AbsenceCnddtSlctr"
Option Explicit

Implements obj_ILookupCandidateSelector

Private m_CorridorDateColumnAlias As String
Private m_OrderDateColumnAlias As String
Private m_MinDate As Date
Private m_MaxDate As Date
Private m_ReferenceDate As Date
Private m_IsInitialized As Boolean

Public Function Initialize( _
    ByVal corridorDateColumnAlias As String, _
    ByVal orderDateColumnAlias As String, _
    ByVal minDate As Date, _
    ByVal maxDate As Date, _
    ByVal referenceDate As Date _
) As Boolean
    corridorDateColumnAlias = VBA.Trim$(corridorDateColumnAlias)
    orderDateColumnAlias = VBA.Trim$(orderDateColumnAlias)
    If VBA.Len(corridorDateColumnAlias) = 0 Or _
        VBA.Len(orderDateColumnAlias) = 0 Then Exit Function
    If minDate <= 0 Then Exit Function
    If maxDate < minDate Then Exit Function
    If referenceDate < minDate Or referenceDate > maxDate Then Exit Function

    m_CorridorDateColumnAlias = corridorDateColumnAlias
    m_OrderDateColumnAlias = orderDateColumnAlias
    m_MinDate = VBA.DateValue(minDate)
    m_MaxDate = VBA.DateValue(maxDate)
    m_ReferenceDate = VBA.DateValue(referenceDate)
    m_IsInitialized = True
    Initialize = True
End Function

Private Function obj_ILookupCandidateSelector_TrySelectCandidateRow( _
    ByVal candidateTable As obj_TableDynamic, _
    ByVal candidateRow As obj_Row, _
    ByVal extensionTable As obj_TableDynamic, _
    ByVal matchingRowIndexes As Collection, _
    ByRef outSelectedRowIndex As Long _
) As Boolean
    Dim corridorDateColumnIndex As Long
    Dim orderDateColumnIndex As Long
    Dim rowIndexObj As Variant
    Dim rowIndex As Long
    Dim extensionRow As obj_Row
    Dim corridorDate As Date
    Dim orderDate As Date
    Dim selectedOrderDate As Date
    Dim rawCorridorDateText As String
    Dim rawOrderDateText As String

    outSelectedRowIndex = 0
    If Not m_IsInitialized Then Exit Function
    If extensionTable Is Nothing Then Exit Function
    If matchingRowIndexes Is Nothing Then Exit Function
    If Not extensionTable.TryGetColumnIndexByAlias( _
        m_CorridorDateColumnAlias, corridorDateColumnIndex) Then
        If Not extensionTable.TryGetColumnIndexByName( _
            m_CorridorDateColumnAlias, corridorDateColumnIndex) Then Exit Function
    End If
    If Not extensionTable.TryGetColumnIndexByAlias( _
        m_OrderDateColumnAlias, orderDateColumnIndex) Then
        If Not extensionTable.TryGetColumnIndexByName( _
            m_OrderDateColumnAlias, orderDateColumnIndex) Then Exit Function
    End If

    For Each rowIndexObj In matchingRowIndexes
        rowIndex = VBA.CLng(rowIndexObj)
        If rowIndex <= 0 Or rowIndex > extensionTable.RowCount Then GoTo ContinueRow
        Set extensionRow = extensionTable.Rows.Item(rowIndex)
        rawCorridorDateText = extensionRow.GetCellValue(corridorDateColumnIndex)
        If Not ex_Helpers.fn_TryResolveDateWithContext( _
            rawCorridorDateText, m_ReferenceDate, corridorDate) Then
            GoTo ContinueRow
        End If
        If VBA.DateValue(corridorDate) < m_MinDate Or _
            VBA.DateValue(corridorDate) > m_MaxDate Then
            GoTo ContinueRow
        End If
        rawOrderDateText = extensionRow.GetCellValue(orderDateColumnIndex)
        If Not ex_Helpers.fn_TryResolveDateWithContext( _
            rawOrderDateText, m_ReferenceDate, orderDate) Then GoTo ContinueRow
        If VBA.DateValue(orderDate) > m_ReferenceDate Then GoTo ContinueRow
        If outSelectedRowIndex = 0 Or _
            VBA.DateValue(orderDate) > selectedOrderDate Then
            outSelectedRowIndex = rowIndex
            selectedOrderDate = VBA.DateValue(orderDate)
        End If
ContinueRow:
    Next rowIndexObj

    obj_ILookupCandidateSelector_TrySelectCandidateRow = True
End Function
