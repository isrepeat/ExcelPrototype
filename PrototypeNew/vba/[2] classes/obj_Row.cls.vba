VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "obj_Row"
Option Explicit
#Const LOGGING_DEBUG_ENABLED = True
#Const LOGGING_VERBOSE_ENABLED = False

Implements obj_IClonable

Private m_Cells() As obj_Cell
Private m_CellCount As Long
Private m_Desc As String
Private m_Index As Long
Private m_IsDisposed As Boolean
Private m_TableSchema As obj_DynamicTableSchema

Private Sub Class_Initialize()
#If LOGGING_VERBOSE_ENABLED Then
    ex_Core.fn_Diagnostic_LogVerbose "lifecycle:" & VBA.TypeName(Me) & ".Class_Initialize"
#End If
    m_CellCount = 0
    m_Desc = VBA.vbNullString
    m_Index = 0
End Sub

Private Sub Class_Terminate()
#If LOGGING_VERBOSE_ENABLED Then
    ex_Core.fn_Diagnostic_LogVerbose "lifecycle:" & VBA.TypeName(Me) & ".Class_Terminate"
#End If
    ' В VBA Class_Terminate вызывается в момент, когда рантайм уже
    ' освобождает объект. Не запускаем здесь Dispose, потому что ручное
    ' Erase массива obj_Cell внутри деструктора может остановить выполнение
    ' во время обычного переприсваивания obj_Row в SQL row loop.
    m_IsDisposed = True
End Sub

' //
' // Interface
' //
Private Function obj_IClonable_Clone(Optional ByVal targetColumnCount As Long = 0) As Object
    Set obj_IClonable_Clone = Me.Clone(targetColumnCount)
End Function

' //
' // Properties
' //
Public Property Get CellCount() As Long
    CellCount = m_CellCount
End Property

Public Property Get Desc() As String
    Desc = m_Desc
End Property

Public Property Let Desc(ByVal valueText As String)
    m_Desc = VBA.CStr(valueText)
End Property

Public Property Get Index() As Long
    Index = m_Index
End Property

Public Property Let Index(ByVal value As Long)
    m_Index = value
End Property

Public Property Get CellValues() As Collection
    Dim result As Collection
    Dim i As Long

    Set result = New Collection
    For i = 1 To m_CellCount
        result.Add Me.GetCellValue(i)
    Next i

    Set CellValues = result
End Property

Public Property Get Cells() As Collection
    Dim result As Collection
    Dim i As Long
    Dim cellObj As obj_Cell

    Set result = New Collection
    For i = 1 To m_CellCount
        Set cellObj = private_GetCellObject(i)
        If Not cellObj Is Nothing Then result.Add cellObj
    Next i

    Set Cells = result
End Property

' //
' // API
' //
Public Function Initialize() As Boolean
#If LOGGING_VERBOSE_ENABLED Then
    ex_Core.fn_Diagnostic_LogVerbose "lifecycle:" & VBA.TypeName(Me) & ".Initialize"
#End If
    Initialize = True
End Function

Public Sub Dispose()
#If LOGGING_VERBOSE_ENABLED Then
    ex_Core.fn_Diagnostic_LogVerbose "lifecycle:" & VBA.TypeName(Me) & ".Dispose"
#End If
    If m_IsDisposed Then Exit Sub
    m_IsDisposed = True
    On Error Resume Next
    Erase m_Cells
    Set m_TableSchema = Nothing
    m_CellCount = 0
    m_Desc = VBA.vbNullString
    m_Index = 0
    On Error GoTo 0
End Sub

Public Sub PushCellRaw( _
    ByVal value As Variant, _
    Optional ByVal descText As String = VBA.vbNullString _
)
    Dim cellObj As obj_Cell

    Set cellObj = New obj_Cell
    cellObj.Value = VBA.CStr(value)
    cellObj.Desc = VBA.CStr(descText)
    Call Me.PushCell(cellObj)
End Sub

Public Function PushCell(ByVal cell As obj_Cell) As Boolean
    If cell Is Nothing Then Exit Function
    If Not private_EnsureCapacity(m_CellCount + 1) Then Exit Function

    m_CellCount = m_CellCount + 1
    Set m_Cells(m_CellCount) = private_CloneCell(cell)
    PushCell = Not m_Cells(m_CellCount) Is Nothing
End Function

Public Function InsertCellAt( _
    ByVal oneBasedIndex As Long, _
    Optional ByVal value As Variant = VBA.vbNullString, _
    Optional ByVal cellDesc As String = VBA.vbNullString _
) As Boolean
    Dim i As Long
    Dim insertedCell As obj_Cell

    If oneBasedIndex <= 0 Then Exit Function
    If oneBasedIndex > m_CellCount + 1 Then Exit Function
    If Not private_EnsureCapacity(m_CellCount + 1) Then Exit Function

    For i = m_CellCount To oneBasedIndex Step -1
        Set m_Cells(i + 1) = m_Cells(i)
    Next i

    Set insertedCell = New obj_Cell
    insertedCell.Value = VBA.CStr(value)
    insertedCell.Desc = VBA.CStr(cellDesc)
    Set m_Cells(oneBasedIndex) = insertedCell

    m_CellCount = m_CellCount + 1
    InsertCellAt = True
End Function

Public Function SetCellRaw( _
    ByVal oneBasedIndex As Long, _
    ByVal value As Variant, _
    Optional ByVal descText As String = VBA.vbNullString _
) As Boolean
    Dim cellObj As obj_Cell

    Set cellObj = New obj_Cell
    cellObj.Value = VBA.CStr(value)
    cellObj.Desc = VBA.CStr(descText)

    SetCellRaw = Me.AssignCellAt(oneBasedIndex, cellObj)
End Function

Public Function AddCellTag(ByVal oneBasedIndex As Long, ByVal tagName As String) As Boolean
    Dim cellObj As obj_Cell

    Set cellObj = private_GetCellObject(oneBasedIndex)
    If cellObj Is Nothing Then Exit Function
    AddCellTag = cellObj.AddTag(tagName)
End Function

Public Function AssignCellAt( _
    ByVal oneBasedIndex As Long, _
    ByVal cell As obj_Cell _
    ) As Boolean
    If oneBasedIndex <= 0 Then
#If LOGGING_DEBUG_ENABLED Then
        ex_Core.fn_Diagnostic_LogError "obj_Row: cell index must be greater than zero."
#End If
        Exit Function
    End If
    If cell Is Nothing Then Exit Function

    If oneBasedIndex > m_CellCount Then
        If Not private_EnsureCapacity(oneBasedIndex) Then Exit Function
        m_CellCount = oneBasedIndex
    End If

    Set m_Cells(oneBasedIndex) = private_CloneCell(cell)
    AssignCellAt = Not m_Cells(oneBasedIndex) Is Nothing
End Function

Public Function GetCellValue(ByVal oneBasedIndex As Long) As String
    Dim cellObj As obj_Cell

    Set cellObj = private_GetCellObject(oneBasedIndex)
    If cellObj Is Nothing Then Exit Function

    GetCellValue = cellObj.Value
End Function

' Binds shared table schema without copying aliases into every row.
Public Sub BindTableSchema(ByVal tableSchema As obj_DynamicTableSchema)
    Set m_TableSchema = tableSchema
End Sub

Public Sub DetachTableSchema()
    Set m_TableSchema = Nothing
End Sub

' Exposes schema lookup for callers that need the physical cell position.
Public Function TryGetColumnIndex( _
    ByVal aliasOrName As String, _
    ByRef outIndex As Long _
) As Boolean
    outIndex = 0
    If m_TableSchema Is Nothing Then
#If LOGGING_DEBUG_ENABLED Then
        ex_Core.fn_Diagnostic_LogError "obj_Row: table schema is not bound; cannot resolve column '" & aliasOrName & "'."
#End If
        Exit Function
    End If
    TryGetColumnIndex = m_TableSchema.TryGetColumnIndex(aliasOrName, outIndex)
End Function

' Returns the cell object by column alias/name without exposing an index.
Public Function TryGetCellByColumn( _
    ByVal aliasOrName As String, _
    ByRef outCell As obj_Cell _
) As Boolean
    Dim columnIndex As Long

    Set outCell = Nothing
    If Not Me.TryGetColumnIndex(aliasOrName, columnIndex) Then Exit Function
    Set outCell = private_GetCellObject(columnIndex)
    TryGetCellByColumn = Not outCell Is Nothing
End Function

' Resolves a column alias first and its display name second.
Public Function TryGetCellValueByColumn( _
    ByVal aliasOrName As String, _
    ByRef outValue As String _
) As Boolean
    Dim columnIndex As Long
    Dim cellObj As obj_Cell

    outValue = VBA.vbNullString
    If Not Me.TryGetColumnIndex(aliasOrName, columnIndex) Then
#If LOGGING_DEBUG_ENABLED Then
        ex_Core.fn_Diagnostic_LogError "obj_Row: column alias/name '" & aliasOrName & "' was not found."
#End If
        VBA.MsgBox "PrototypeNew: row column alias/name '" & aliasOrName & "' was not found or the row is not bound to a table schema.", _
            VBA.vbExclamation, "PrototypeNew / row"
        Exit Function
    End If

    Set cellObj = private_GetCellObject(columnIndex)
    If cellObj Is Nothing Then
#If LOGGING_DEBUG_ENABLED Then
        ex_Core.fn_Diagnostic_LogError "obj_Row: resolved column '" & aliasOrName & "' has no cell at index " & VBA.CStr(columnIndex) & "."
#End If
        VBA.MsgBox "PrototypeNew: row column '" & aliasOrName & "' exists in the table schema, but the row has no corresponding cell.", _
            VBA.vbExclamation, "PrototypeNew / row"
        Exit Function
    End If

    outValue = cellObj.Value
    TryGetCellValueByColumn = True
End Function

Public Function TryFindCellIndexByDesc(ByVal descToken As String, ByRef outIndex As Long) As Boolean
    Dim i As Long
    Dim cellObj As obj_Cell
    Dim normalizedToken As String
    Dim cellDesc As String

    outIndex = 0
    normalizedToken = VBA.Trim$(VBA.CStr(descToken))
    If VBA.Len(normalizedToken) = 0 Then Exit Function

    For i = 1 To m_CellCount
        Set cellObj = private_GetCellObject(i)
        If cellObj Is Nothing Then GoTo ContinueCell
        cellDesc = VBA.Trim$(cellObj.Desc)
        If VBA.Len(cellDesc) = 0 Then GoTo ContinueCell
        If VBA.InStr(1, cellDesc, normalizedToken, VBA.vbTextCompare) > 0 Then
            outIndex = i
            TryFindCellIndexByDesc = True
            Exit Function
        End If
ContinueCell:
    Next i
End Function

Public Function TryGetCellByDesc(ByVal descToken As String, ByRef outCell As obj_Cell) As Boolean
    Dim cellIndex As Long

    Set outCell = Nothing
    If Not Me.TryFindCellIndexByDesc(descToken, cellIndex) Then Exit Function
    If cellIndex <= 0 Then Exit Function

    Set outCell = private_GetCellObject(cellIndex)
    TryGetCellByDesc = Not outCell Is Nothing
End Function

Public Function TryGetCellAt(ByVal oneBasedIndex As Long, ByRef outCell As obj_Cell) As Boolean
    Set outCell = Nothing
    Set outCell = private_GetCellObject(oneBasedIndex)
    TryGetCellAt = Not outCell Is Nothing
End Function

Public Function IsCellVirtual(ByVal oneBasedIndex As Long) As Boolean
    Dim cellObj As obj_Cell

    Set cellObj = private_GetCellObject(oneBasedIndex)
    If cellObj Is Nothing Then Exit Function

    IsCellVirtual = cellObj.IsVirtual
End Function

Public Sub CopyToMatrixRow(ByRef targetMatrix As Variant, ByVal matrixRow As Long, ByVal columnCount As Long)
    Dim i As Long
    Dim maxCols As Long

    If matrixRow <= 0 Then Exit Sub
    If columnCount <= 0 Then Exit Sub

    maxCols = columnCount
    If m_CellCount < maxCols Then maxCols = m_CellCount

    For i = 1 To maxCols
        targetMatrix(matrixRow, i) = Me.GetCellValue(i)
    Next i
End Sub

Public Function Clone(Optional ByVal targetColumnCount As Long = 0) As Object
    Dim result As obj_Row
    Dim i As Long
    Dim cellObj As obj_Cell

    Set result = New obj_Row
    For i = 1 To m_CellCount
        Set cellObj = private_GetCellObject(i)
        If Not cellObj Is Nothing Then
            If Not result.PushCell(cellObj) Then Exit Function
        Else
            result.PushCellRaw VBA.vbNullString
        End If
    Next i

    If targetColumnCount > result.CellCount Then
        For i = result.CellCount + 1 To targetColumnCount
            result.PushCellRaw VBA.vbNullString
        Next i
    End If

    result.Desc = m_Desc
    result.Index = m_Index
    result.BindTableSchema m_TableSchema

    Set Clone = result
End Function

' //
' // Internal
' //
Private Function private_EnsureCapacity(ByVal requiredCount As Long) As Boolean
    If requiredCount <= 0 Then
        private_EnsureCapacity = True
        Exit Function
    End If

    If requiredCount <= m_CellCount Then
        private_EnsureCapacity = True
        Exit Function
    End If

    If m_CellCount = 0 Then
        ReDim m_Cells(1 To requiredCount)
    Else
        ReDim Preserve m_Cells(1 To requiredCount)
    End If

    private_EnsureCapacity = True
End Function

Private Function private_GetCellObject(ByVal oneBasedIndex As Long) As obj_Cell
    If oneBasedIndex <= 0 Then Exit Function
    If oneBasedIndex > m_CellCount Then Exit Function
    Set private_GetCellObject = m_Cells(oneBasedIndex)
End Function

Private Function private_CloneCell(ByVal sourceCell As obj_Cell) As obj_Cell
    If sourceCell Is Nothing Then Exit Function
    Set private_CloneCell = sourceCell.Clone
End Function
