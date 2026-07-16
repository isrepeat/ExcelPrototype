VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "obj_PEB_AbsenceCnddtSlctr"
Option Explicit

Implements obj_ILookupCandidateSelector

Private m_DateColumnAlias As String
Private m_MinDate As Date
Private m_MaxDate As Date
Private m_IsInitialized As Boolean

Public Function Initialize( _
    ByVal dateColumnAlias As String, _
    ByVal minDate As Date, _
    ByVal maxDate As Date _
) As Boolean
    dateColumnAlias = VBA.Trim$(dateColumnAlias)
    If VBA.Len(dateColumnAlias) = 0 Then Exit Function
    If minDate <= 0 Then Exit Function
    If maxDate < minDate Then Exit Function

    m_DateColumnAlias = dateColumnAlias
    m_MinDate = VBA.DateValue(minDate)
    m_MaxDate = VBA.DateValue(maxDate)
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
    Dim dateColumnIndex As Long
    Dim rowIndexObj As Variant
    Dim rowIndex As Long
    Dim extensionRow As obj_Row
    Dim actualDate As Date
    Dim rawDateText As String

    outSelectedRowIndex = 0
    If Not m_IsInitialized Then Exit Function
    If extensionTable Is Nothing Then Exit Function
    If matchingRowIndexes Is Nothing Then Exit Function
    If Not extensionTable.TryGetColumnIndexByAlias(m_DateColumnAlias, dateColumnIndex) Then
        If Not extensionTable.TryGetColumnIndexByName(m_DateColumnAlias, dateColumnIndex) Then Exit Function
    End If

    For Each rowIndexObj In matchingRowIndexes
        rowIndex = VBA.CLng(rowIndexObj)
        If rowIndex <= 0 Or rowIndex > extensionTable.RowCount Then GoTo ContinueRow
        Set extensionRow = extensionTable.Rows.Item(rowIndex)
        rawDateText = extensionRow.GetCellValue(dateColumnIndex)
        If Not ex_Helpers.fn_TryResolveDateWithContext(rawDateText, m_MinDate, actualDate) Then
            GoTo ContinueRow
        End If
        If VBA.DateValue(actualDate) < m_MinDate Or VBA.DateValue(actualDate) > m_MaxDate Then
            GoTo ContinueRow
        End If

        outSelectedRowIndex = rowIndex
        Exit For
ContinueRow:
    Next rowIndexObj

    obj_ILookupCandidateSelector_TrySelectCandidateRow = True
End Function
