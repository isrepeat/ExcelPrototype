VERSION 1.0 CLASS
BEGIN
  MultiUse = -1
END
Attribute VB_Name = "obj_QueryCondition"
Option Explicit

Private m_isInitialized As Boolean
Private m_isDisposed As Boolean

Public ColumnName As String
Public Operation As en_QueryOp
Public Value As Variant
Public ValueType As en_QueryType
Public NormalizeText As Boolean
Public CaseSensitive As Boolean
Private m_columnIndex As Long

' //
' // Lifecycle
' //
Private Sub Class_Initialize()
    Operation = QueryEquals
    ValueType = QueryText
End Sub

Private Sub Class_Terminate()
    Me.Dispose
End Sub

' //
' // API
' //
Public Function Initialize() As Boolean
    If m_isDisposed Or m_isInitialized Then
        Exit Function
    End If
    m_isInitialized = True
    Initialize = True
End Function

Public Sub Dispose()
    If m_isDisposed Then
        Exit Sub
    End If
    m_isDisposed = True
    m_isInitialized = False
    ColumnName = VBA.vbNullString
    Operation = QueryEquals
    Value = Empty
    ValueType = QueryText
    NormalizeText = False
    CaseSensitive = False
    m_columnIndex = 0
End Sub

Public Sub Validate(ByVal headers As Variant)
    Dim index As Long
    Dim comparison As Long

    If m_isDisposed Or Not m_isInitialized Then
        Err.Raise VBA.vbObjectError + 2165, , "Object is not initialized or is disposed."
    End If
    index = ex_TableQuery.fn_HeaderIndex(headers, ColumnName)
    m_columnIndex = index
    If Operation < QueryEquals Or Operation > QueryLessOrEqual Then
        Err.Raise VBA.vbObjectError + 2110, , "Unsupported condition operator."
    End If
    If ValueType < QueryText Or ValueType > QueryBoolean Then
        Err.Raise VBA.vbObjectError + 2111, , "Unsupported condition type."
    End If
    If Operation = QueryIsEmpty Or Operation = QueryIsNotEmpty Then
        Exit Sub
    End If
    If VBA.IsError(Value) Or VBA.IsNull(Value) Or VBA.IsEmpty(Value) Or VBA.IsObject(Value) Or VBA.IsArray(Value) Then
        Err.Raise VBA.vbObjectError + 2112, , "Condition requires a scalar value."
    End If
    If Operation >= QueryContains And Operation <= QueryEndsWith And ValueType <> QueryText Then
        Err.Raise VBA.vbObjectError + 2113, , "Text operator requires QueryText."
    End If
    comparison = ex_TableQuery.fn_Compare(Value, Value, ValueType, NormalizeText, CaseSensitive)
End Sub

Public Function MatchesRow( _
    ByVal values As Variant, _
    ByVal row As Long _
) As Boolean
    If m_isDisposed Or Not m_isInitialized Then
        Err.Raise VBA.vbObjectError + 2165, , "Object is not initialized or is disposed."
    End If
    MatchesRow = Me.Matches(values(row, m_columnIndex))
End Function

Public Function Matches(ByVal actual As Variant) As Boolean
    Dim blank As Boolean
    Dim a As String
    Dim b As String
    Dim mode As VbCompareMethod
    Dim comparison As Long

    If m_isDisposed Or Not m_isInitialized Then
        Err.Raise VBA.vbObjectError + 2165, , "Object is not initialized or is disposed."
    End If
    If VBA.IsError(actual) Then
        Err.Raise VBA.vbObjectError + 2114, , "Excel error in condition column: " & ColumnName
    End If
    blank = VBA.IsEmpty(actual) Or VBA.IsNull(actual)
    If Not blank And VBA.VarType(actual) = VBA.vbString Then
        a = VBA.CStr(actual)
        If NormalizeText Then
            a = ex_TableQuery.fn_NormalizeText(a)
        End If
        blank = (VBA.Len(a) = 0)
    End If
    If Operation = QueryIsEmpty Then
        Matches = blank
        Exit Function
    End If
    If Operation = QueryIsNotEmpty Then
        Matches = Not blank
        Exit Function
    End If
    If blank Then
        Exit Function
    End If
    If Operation >= QueryContains And Operation <= QueryEndsWith Then
        a = VBA.CStr(actual)
        b = VBA.CStr(Value)
        If NormalizeText Then
            a = ex_TableQuery.fn_NormalizeText(a)
            b = ex_TableQuery.fn_NormalizeText(b)
        End If
        mode = VBA.vbTextCompare
        If CaseSensitive Then
            mode = VBA.vbBinaryCompare
        End If
        Select Case Operation
            Case QueryContains
                Matches = VBA.InStr(1, a, b, mode) > 0
            Case QueryStartsWith
                Matches = VBA.StrComp(VBA.Left$(a, VBA.Len(b)), b, mode) = 0
            Case QueryEndsWith
                Matches = VBA.StrComp(VBA.Right$(a, VBA.Len(b)), b, mode) = 0
        End Select
        Exit Function
    End If
    comparison = ex_TableQuery.fn_Compare(actual, Value, ValueType, NormalizeText, CaseSensitive)
    Select Case Operation
        Case QueryEquals
            Matches = comparison = 0
        Case QueryNotEquals
            Matches = comparison <> 0
        Case QueryGreater
            Matches = comparison > 0
        Case QueryGreaterOrEqual
            Matches = comparison >= 0
        Case QueryLess
            Matches = comparison < 0
        Case QueryLessOrEqual
            Matches = comparison <= 0
    End Select
End Function