Attribute VB_Name = "ex_TableQuery"
Option Explicit

Public Enum en_QueryOp
    QueryEquals = 1
    QueryNotEquals
    QueryContains
    QueryStartsWith
    QueryEndsWith
    QueryIsEmpty
    QueryIsNotEmpty
    QueryGreater
    QueryGreaterOrEqual
    QueryLess
    QueryLessOrEqual
End Enum

Public Enum en_QueryType
    QueryText = 1
    QueryNumber
    QueryDate
    QueryBoolean
End Enum

Public Enum en_QueryBackend
    QueryAuto = 0
    QueryExcel = 1
    QueryAdo = 2
End Enum

' --------------------------------------
' namespace API {
' --------------------------------------
Public Function fn_ReadColumns( _
    ByVal headers As Variant, _
    ByVal query As obj_TableQuery _
) As Collection
    Dim requested As Collection
    Dim names As New Collection
    Dim seen As Object
    Dim item As Variant
    Dim ordering As obj_QueryOrder
    Dim index As Long
    Dim i As Long

    Set seen = VBA.CreateObject("Scripting.Dictionary")
    Set requested = query.Columns
    If requested.Count = 0 Then
        For i = LBound(headers) To UBound(headers)
            requested.Add headers(i)
        Next i
    End If
    query.Filter.CollectColumns requested
    For Each ordering In query.Orders
        requested.Add ordering.ColumnName
    Next ordering
    For Each item In requested
        index = fn_HeaderIndex(headers, VBA.CStr(item))
        If Not seen.Exists(VBA.CStr(index)) Then
            names.Add VBA.CStr(headers(LBound(headers) + index - 1))
            seen.Add VBA.CStr(index), True
        End If
    Next item
    query.Filter.Validate headers
    Set fn_ReadColumns = names
End Function

Public Function fn_HeaderIndex( _
    ByVal headers As Variant, _
    ByVal name As String _
) As Long
    Dim i As Long

    For i = LBound(headers) To UBound(headers)
        If VBA.StrComp(VBA.Trim$(VBA.CStr(headers(i))), VBA.Trim$(name), VBA.vbTextCompare) = 0 Then
            If fn_HeaderIndex > 0 Then
                Err.Raise VBA.vbObjectError + 2100, , "Ambiguous column: " & name
            End If
            fn_HeaderIndex = i - LBound(headers) + 1
        End If
    Next i
    If fn_HeaderIndex = 0 Then
        Err.Raise VBA.vbObjectError + 2101, , "Column not found: " & name
    End If
End Function

Public Function fn_NormalizeText(ByVal value As String) As String
    value = VBA.Replace$(value, VBA.ChrW$(160), " ")
    value = VBA.Replace$(VBA.Replace$(VBA.Replace$(value, VBA.vbCr, ""), VBA.vbLf, ""), VBA.vbTab, "")
    fn_NormalizeText = VBA.Trim$(value)
End Function

Public Function fn_Compare( _
    ByVal actual As Variant, _
    ByVal expected As Variant, _
    ByVal kind As en_QueryType, _
    ByVal normalize As Boolean, _
    ByVal caseSensitive As Boolean _
) As Long
    Dim a As Variant
    Dim b As Variant
    Dim mode As VbCompareMethod

    If VBA.IsError(actual) Or VBA.IsNull(actual) Or VBA.IsEmpty(actual) Then
        Err.Raise VBA.vbObjectError + 2102, , "Cannot compare an empty or error value."
    End If
    Select Case kind
        Case QueryText
            a = VBA.CStr(actual)
            b = VBA.CStr(expected)
            If normalize Then
                a = fn_NormalizeText(a)
                b = fn_NormalizeText(b)
            End If
            mode = VBA.vbTextCompare
            If caseSensitive Then
                mode = VBA.vbBinaryCompare
            End If
            fn_Compare = VBA.StrComp(a, b, mode)
            Exit Function
        Case QueryNumber
            If VBA.VarType(actual) = VBA.vbBoolean Or Not VBA.IsNumeric(actual) Or Not VBA.IsNumeric(expected) Then
                Err.Raise VBA.vbObjectError + 2103, , "Expected a number."
            End If
            a = VBA.CDbl(actual)
            b = VBA.CDbl(expected)
        Case QueryDate
            If VBA.IsNumeric(actual) Then
                a = VBA.CDbl(actual)
            Else
                If Not VBA.IsDate(actual) Then
                    Err.Raise VBA.vbObjectError + 2104, , "Expected a date, received: " & VBA.CStr(actual)
                End If
                a = VBA.CDbl(VBA.CDate(actual))
            End If
            b = VBA.CDbl(VBA.CDate(expected))
        Case QueryBoolean
            If VBA.VarType(actual) <> VBA.vbBoolean Or VBA.VarType(expected) <> VBA.vbBoolean Then
                Err.Raise VBA.vbObjectError + 2105, , "Expected a Boolean."
            End If
            a = actual
            b = expected
        Case Else
            Err.Raise VBA.vbObjectError + 2106, , "Unsupported comparison type."
    End Select
    If a < b Then
        fn_Compare = -1
    End If
    If a > b Then
        fn_Compare = 1
    End If
End Function

Public Function fn_TryApply( _
    ByVal dataTable As obj_DataTable, _
    ByVal query As obj_TableQuery, _
    ByRef output As obj_DataTable, _
    ByRef diagnostic As String _
) As Boolean
    Dim columns As Collection
    Dim headers As Variant
    Dim values As Variant
    Dim result As Variant
    Dim resultHeaders As Variant
    Dim selected() As Long
    Dim indexes() As Long
    Dim count As Long
    Dim i As Long
    Dim j As Long
    Dim row As Long
    Dim ordering As obj_QueryOrder
    Dim col As Long
    Dim group As obj_QueryGroup

    On Error GoTo EH
    Set output = Nothing
    headers = dataTable.Headers
    values = dataTable.Values
    Set columns = query.Columns
    If columns.Count = 0 Then
        For i = LBound(headers) To UBound(headers)
            columns.Add VBA.CStr(headers(i))
        Next i
    End If
    ReDim selected(1 To columns.Count)
    ReDim resultHeaders(0 To columns.Count - 1)
    For i = 1 To columns.Count
        selected(i) = fn_HeaderIndex(headers, VBA.CStr(columns(i)))
        resultHeaders(i - 1) = columns(i)
    Next i
    Set group = query.Filter
    group.Validate headers
    For Each ordering In query.Orders
        col = fn_HeaderIndex(headers, ordering.ColumnName)
        If ordering.ValueType < QueryText Or ordering.ValueType > QueryBoolean Then
            Err.Raise VBA.vbObjectError + 2106, , "Unsupported order type."
        End If
    Next ordering
    If dataTable.RowCount > 0 Then
        ReDim indexes(1 To dataTable.RowCount)
    End If
    For row = 1 To dataTable.RowCount
        If group.Matches(values, row, headers) Then
            count = count + 1
            indexes(count) = row
        End If
    Next row
    If query.Orders.Count > 0 And count > 1 Then
        private_Sort indexes, count, values, headers, query.Orders
    End If
    If query.Limit > 0 And count > query.Limit Then
        count = query.Limit
    End If
    If count > 0 Then
        ReDim result(1 To count, 1 To columns.Count)
        For i = 1 To count
            For j = 1 To columns.Count
                result(i, j) = values(indexes(i), selected(j))
            Next j
        Next i
    End If
    Set output = New obj_DataTable
    fn_TryApply = output.Initialize(resultHeaders, result, count, diagnostic)
    Exit Function
EH:
    diagnostic = "Query evaluation: " & Err.Description
    Set output = Nothing
End Function
' --------------------------------------
' } // namespace API
' --------------------------------------

' //
' // Private
' //
Private Sub private_Sort( _
    ByRef indexes() As Long, _
    ByVal count As Long, _
    ByVal values As Variant, _
    ByVal headers As Variant, _
    ByVal orders As Collection _
)
    Dim buffer() As Long
    Dim columns() As Long
    Dim width As Long
    Dim left As Long
    Dim middle As Long
    Dim last As Long
    Dim a As Long
    Dim b As Long
    Dim k As Long
    Dim i As Long
    Dim comparison As Long
    Dim ordering As obj_QueryOrder

    ReDim buffer(1 To count)
    ReDim columns(1 To orders.Count)
    For i = 1 To orders.Count
        columns(i) = fn_HeaderIndex(headers, orders(i).ColumnName)
    Next i
    ' Стабильная сортировка слиянием: O(n log n), без перестановки исходных значений.
    width = 1
    Do While width < count
        For left = 1 To count Step width * 2
            middle = left + width - 1
            last = left + width * 2 - 1
            If middle > count Then
                middle = count
            End If
            If last > count Then
                last = count
            End If
            a = left
            b = middle + 1
            For k = left To last
                If a > middle Then
                    buffer(k) = indexes(b)
                    b = b + 1
                ElseIf b > last Then
                    buffer(k) = indexes(a)
                    a = a + 1
                Else
                    comparison = 0
                    For i = 1 To orders.Count
                        Set ordering = orders(i)
                        comparison = private_CompareOrder(values(indexes(a), columns(i)), values(indexes(b), columns(i)), ordering)
                        If comparison <> 0 Then
                            Exit For
                        End If
                    Next i
                    If comparison <= 0 Then
                        buffer(k) = indexes(a)
                        a = a + 1
                    Else
                        buffer(k) = indexes(b)
                        b = b + 1
                    End If
                End If
            Next k
        Next left
        For k = 1 To count
            indexes(k) = buffer(k)
        Next k
        width = width * 2
    Loop
End Sub

' //
' // Private
' //
Private Function private_CompareOrder( _
    ByVal a As Variant, _
    ByVal b As Variant, _
    ByVal ordering As obj_QueryOrder _
) As Long
    Dim emptyA As Boolean
    Dim emptyB As Boolean

    If VBA.IsError(a) Or VBA.IsError(b) Then
        Err.Raise VBA.vbObjectError + 2107, , "Cannot sort Excel error values."
    End If
    emptyA = VBA.IsNull(a) Or VBA.IsEmpty(a)
    emptyB = VBA.IsNull(b) Or VBA.IsEmpty(b)
    If emptyA And emptyB Then
        Exit Function
    End If
    If emptyA Then
        private_CompareOrder = -1
    ElseIf emptyB Then
        private_CompareOrder = 1
    Else
        private_CompareOrder = fn_Compare(a, b, ordering.ValueType, ordering.NormalizeText, ordering.CaseSensitive)
    End If
    If ordering.Descending Then
        private_CompareOrder = -private_CompareOrder
    End If
End Function